"""
Motor de extracción estructurada: PDF -> IR (representación intermedia) -> HTML editable.

Objetivo: reconstruir el documento como objetos editables (títulos, párrafos,
listas, tablas, imágenes) conservando estilos inline, alineación, orden de
lectura, encabezados/pies y geometría de página.
"""
from __future__ import annotations

import base64
import collections
import html
import os
import re
import shutil
import subprocess
import tempfile
from typing import Any

import pymupdf as fitz

BULLET_CHARS = "\u2022\u25e6\u25aa\u00b7\u25cf\u25cb\u25a0\u25a1\u2013\u2014\u2043\u2219"
BULLET_RE = re.compile(r"^\s*[" + BULLET_CHARS + r"]\s+")
NUM_RE = re.compile(r"^\s*(?:\d{1,3}|[a-zA-Z]|[ivxIVX]{1,4})[.)]\s+")
PAGENUM_RE = re.compile(
    r"(p[áa]gina|page|pág\.?|seite|pagina)?\s*\d+\s*(de|of|/|von)?\s*\d*", re.I
)

MIN_IMG_SIDE = 24  # px, descarta viñetas/reglas decorativas


# --------------------------------------------------------------------------- #
# utilidades
# --------------------------------------------------------------------------- #
def _is_bold(flags: int) -> bool:
    return bool(flags & 16)


def _is_italic(flags: int) -> bool:
    return bool(flags & 2)


def _is_mono(flags: int) -> bool:
    return bool(flags & 8)


def _is_super(flags: int) -> bool:
    return bool(flags & 1)


def _color_hex(c: int) -> str:
    if c is None:
        return ""
    return "#%06x" % (c & 0xFFFFFF)


def _clean_font(name: str) -> str:
    name = re.sub(r"^[A-Z]{6}\+", "", name or "")
    return name.split("-")[0].split(",")[0].strip() or "Calibri"


def _norm_for_repeat(text: str) -> str:
    t = re.sub(r"\d+", "#", text.strip().lower())
    return re.sub(r"\s+", " ", t)


def _looks_like_pagenum(text: str) -> bool:
    t = text.strip()
    if not t or len(t) > 60:
        return False
    return bool(re.fullmatch(r"[-–—\s]*(p[áa]gina|page|pág\.?)?\s*\d+\s*([-–—/]|de|of)\s*\d*\s*[-–—\s]*", t, re.I))


# --------------------------------------------------------------------------- #
# estilos inline
# --------------------------------------------------------------------------- #
def _runs_from_line(line: dict) -> list[dict]:
    runs: list[dict] = []
    for sp in line.get("spans", []):
        txt = sp.get("text", "")
        if not txt:
            continue
        run = {
            "t": txt,
            "b": _is_bold(sp.get("flags", 0)),
            "i": _is_italic(sp.get("flags", 0)),
            "size": round(float(sp.get("size", 0)), 1),
            "font": _clean_font(sp.get("font", "")),
            "color": _color_hex(sp.get("color")),
            "sup": _is_super(sp.get("flags", 0)),
            "mono": _is_mono(sp.get("flags", 0)),
            "bbox": sp.get("bbox"),
        }
        if runs and all(
            runs[-1][k] == run[k] for k in ("b", "i", "size", "font", "color", "sup", "mono")
        ):
            runs[-1]["t"] += run["t"]
        else:
            runs.append(run)
    return runs


def _dominant(runs: list[dict]) -> dict:
    best, best_len = None, -1
    for r in runs:
        n = len(r["t"].strip())
        if n > best_len:
            best, best_len = r, n
    return best or {"size": 0, "font": "", "b": False, "i": False, "color": "#000000", "mono": False}


def _link_for(run_bbox, links):
    if not run_bbox or not links:
        return None
    cx = (run_bbox[0] + run_bbox[2]) / 2
    cy = (run_bbox[1] + run_bbox[3]) / 2
    for lnk in links:
        r = lnk.get("from")
        if r and r.x0 - 1 <= cx <= r.x1 + 1 and r.y0 - 1 <= cy <= r.y1 + 1:
            return lnk.get("uri")
    return None


def _runs_to_html(runs: list[dict], links: list[dict] | None = None) -> str:
    out = []
    for r in runs:
        txt = r["t"].replace("\u00a0", " ")
        txt = html.escape(txt)
        if not txt:
            continue
        if r.get("sup"):
            txt = f"<sup>{txt}</sup>"
        if r.get("mono"):
            txt = f"<code>{txt}</code>"
        if r.get("b"):
            txt = f"<strong>{txt}</strong>"
        if r.get("i"):
            txt = f"<em>{txt}</em>"
        # color real del texto en el PDF (títulos de color, enlaces, gris de cuerpo…)
        color = (r.get("color") or "").lower()
        if color and color != "#000000":
            txt = f'<span style="color:{color}">{txt}</span>'
        uri = _link_for(r.get("bbox"), links)
        if uri and not uri.startswith(("javascript:", "data:")):
            txt = f'<a href="{html.escape(uri, quote=True)}">{txt}</a>'
        out.append(txt)
    return "".join(out)


# --------------------------------------------------------------------------- #
# detección de tablas
# --------------------------------------------------------------------------- #
def _cell_text(page, rect) -> str:
    """Texto de una celda conservando los espacios entre palabras y las líneas."""
    if rect is None:
        return ""
    try:
        words = page.get_text("words", clip=fitz.Rect(rect))
    except Exception:
        return ""
    if not words:
        return ""
    r = fitz.Rect(rect)
    inside = []
    for w in words:
        cx, cy = (w[0] + w[2]) / 2, (w[1] + w[3]) / 2
        if r.x0 - 1 <= cx <= r.x1 + 1 and r.y0 - 1 <= cy <= r.y1 + 1:
            inside.append(w)
    if not inside:
        return ""
    lines: dict[tuple, list] = collections.defaultdict(list)
    for w in inside:
        lines[(w[5], w[6])].append(w)
    out = []
    for key in sorted(lines):
        ws = sorted(lines[key], key=lambda w: w[0])
        out.append(" ".join(w[4] for w in ws))
    return " ".join(out).strip()


def _table_to_block(tab, page=None) -> dict:
    try:
        raw_cells = tab.extract()
    except Exception:
        return {}
    if page is not None:
        rows = []
        try:
            for row in tab.rows:
                rows.append([_cell_text(page, c) for c in row.cells])
        except Exception:
            rows = raw_cells
    else:
        rows = raw_cells
    clean: list[list[str]] = []
    for row in rows or []:
        clean.append([(c or "").strip() for c in row])
    if not clean or len(clean) < 1:
        return {}
    ncols = max(len(r) for r in clean)
    for r in clean:
        while len(r) < ncols:
            r.append("")
    # descarta "tablas" de una sola celda con texto (falsos positivos)
    if ncols < 2 and len(clean) < 3:
        return {}
    header_fill = ""
    if page is not None:
        try:
            hdr_rect = None
            try:
                cells0 = [c for c in tab.rows[0].cells if c]
                if cells0:
                    hdr_rect = fitz.Rect(cells0[0])
                    for c in cells0[1:]:
                        hdr_rect |= fitz.Rect(c)
            except Exception:
                hdr_rect = None
            if hdr_rect is None:
                hdr_rect = fitz.Rect(tab.bbox[0], tab.bbox[1], tab.bbox[2],
                                     tab.bbox[1] + max(8, (tab.bbox[3] - tab.bbox[1]) / max(1, len(clean))))
            best_cov = 0.0
            for dr in page.get_drawings():
                if not dr.get("fill"):
                    continue
                r = fitz.Rect(dr["rect"])
                inter = r & hdr_rect
                cov = abs(inter) / abs(hdr_rect) if abs(hdr_rect) else 0
                if cov > 0.55 and cov > best_cov:
                    hexcol = _rgb_hex(dr["fill"])
                    if hexcol and hexcol != "#ffffff":
                        best_cov, header_fill = cov, hexcol
        except Exception:
            header_fill = ""
    block = {
        "type": "table",
        "rows": clean,
        "bbox": [round(v, 2) for v in tab.bbox],
        "header": True,
    }
    if header_fill:
        block["header_fill"] = header_fill
    return block


# --------------------------------------------------------------------------- #
# imágenes
# --------------------------------------------------------------------------- #
def _image_block(page, bbox, doc, raw: bytes | None = None, ext: str = "") -> dict | None:
    x0, y0, x1, y1 = bbox
    w, h = x1 - x0, y1 - y0
    if w < MIN_IMG_SIDE or h < MIN_IMG_SIDE:
        return None
    data = None
    mime = "image/jpeg"
    # 1) bytes originales incrustados en el PDF (sin recomprimir ni renderizar)
    if raw and len(raw) <= 900_000:
        ext = (ext or "").lower()
        mime = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png",
                "jp2": "image/jp2", "jpx": "image/jp2"}.get(ext, "")
        if mime and ext not in ("jp2", "jpx"):
            data = raw
    # 2) renderizado ligero como respaldo
    if data is None:
        try:
            zoom = 130 / 72
            pix = page.get_pixmap(clip=fitz.Rect(bbox), matrix=fitz.Matrix(zoom, zoom))
            if pix.width > 1000:
                zoom2 = 1000 / pix.width
                pix = page.get_pixmap(
                    clip=fitz.Rect(bbox),
                    matrix=fitz.Matrix(zoom * zoom2, zoom * zoom2),
                )
            if pix.n >= 4 or pix.alpha:
                pix = fitz.Pixmap(fitz.csRGB, pix)
            try:
                data = pix.tobytes("jpeg", jpg_quality=72)
                mime = "image/jpeg"
            except Exception:
                data = pix.tobytes("png")
                mime = "image/png"
        except Exception:
            return None
    b64 = base64.b64encode(data).decode("ascii")
    return {
        "type": "image",
        "src": f"data:{mime};base64," + b64,
        "bbox": [round(v, 2) for v in bbox],
        "w": round(w, 2),
        "h": round(h, 2),
        "alt": "",
    }


# --------------------------------------------------------------------------- #
# agrupación de líneas -> párrafos
# --------------------------------------------------------------------------- #
def _group_lines(lines: list[dict], body_size: float, heading_sizes: dict, links: list[dict] | None = None) -> list[dict]:
    """Convierte líneas de un bloque de texto en párrafos / títulos / ítems."""
    out: list[dict] = []
    cur: dict | None = None

    def flush():
        nonlocal cur
        if cur and cur["runs"]:
            cur["text"] = "".join(r["t"] for r in cur["runs"]).strip()
            if cur["text"]:
                out.append(cur)
        cur = None

    pending_bullet = False
    for line in lines:
        runs = _runs_from_line(line)
        if not runs:
            continue
        dom = _dominant(runs)
        raw = "".join(r["t"] for r in runs)
        stripped = raw.strip()

        # viñeta suelta en su propia línea (frecuente en PDF) -> marca el siguiente texto como ítem
        if stripped and all(ch in BULLET_CHARS for ch in stripped):
            flush()
            pending_bullet = True
            continue
        if pending_bullet and stripped:
            flush()
            out.append({
                "type": "li", "ordered": False, "runs": runs,
                "text": stripped, "align": "left",
                "bbox": [round(v, 2) for v in line["bbox"]],
            })
            pending_bullet = False
            continue

        # --- clasificación -------------------------------------------------
        size = dom["size"]
        level = 0
        for lvl, sz in heading_sizes.items():
            if abs(size - sz) <= 0.6:
                level = lvl
                break
        is_heading = level > 0 and len(stripped) < 220

        item_match = BULLET_RE.match(raw) or NUM_RE.match(raw)
        is_item = bool(item_match) and not is_heading

        if is_heading:
            flush()
            out.append(
                {
                    "type": "heading",
                    "level": level,
                    "runs": runs,
                    "text": stripped,
                    "align": _align(line, runs),
                    "bbox": [round(v, 2) for v in line["bbox"]],
                }
            )
            continue

        # estilo distinto al párrafo en curso -> nuevo párrafo
        if cur is not None and (
            abs(dom["size"] - cur["size"]) > 0.8 or dom["b"] != cur["bold"]
        ):
            flush()

        if is_item:
            flush()
            marker = item_match.group(0)
            ordered = bool(NUM_RE.match(raw)) and not BULLET_RE.match(raw)
            first = dict(runs[0])
            first["t"] = first["t"][len(marker):] if first["t"].startswith(marker) else re.sub(
                r"^\s*[" + BULLET_CHARS + r"]\s+|^\s*(?:\d{1,3}|[a-zA-Z]|[ivxIVX]{1,4})[.)]\s+",
                "",
                first["t"],
            )
            body_runs = ([first] if first["t"] else []) + runs[1:]
            cur = {
                "type": "li",
                "ordered": ordered,
                "runs": body_runs,
                "size": dom["size"],
                "bold": dom["b"],
                "align": "left",
                "bbox": [round(v, 2) for v in line["bbox"]],
            }
            continue

        if cur is None:
            cur = {
                "type": "p",
                "runs": runs,
                "size": dom["size"],
                "bold": dom["b"],
                "align": _align(line, runs),
                "bbox": [round(v, 2) for v in line["bbox"]],
            }
        else:
            # une líneas del mismo párrafo
            prev = cur["runs"][-1]["t"] if cur["runs"] else ""
            if prev.endswith("-") and prev.strip():
                cur["runs"][-1]["t"] = prev[:-1]
            elif prev.endswith((" ", "\u00a0")):
                pass
            else:
                cur["runs"].append({"t": " ", "size": cur["size"], "font": "", "b": False,
                                    "i": False, "color": "#000000", "sup": False, "mono": False})
            cur["runs"].extend(runs)
            cur["bbox"][2] = max(cur["bbox"][2], round(line["bbox"][2], 2))
            cur["bbox"][3] = max(cur["bbox"][3], round(line["bbox"][3], 2))
    flush()

    for b in out:
        b.pop("size", None)
        b.pop("bold", None)
        if links:
            b["_links"] = links
    return out


def _align(line: dict, runs: list[dict]) -> str:
    return "left"


def _align_with_page(line: dict, page_rect) -> str:
    """Alineación aproximada comparando con los márgenes de la página."""
    x0, _, x1, _ = line["bbox"]
    pw = page_rect.width
    left_gap = x0
    right_gap = pw - x1
    width = x1 - x0
    if width < pw * 0.55:
        if abs(left_gap - right_gap) < pw * 0.06 and left_gap > pw * 0.15:
            return "center"
        if right_gap < pw * 0.12 and left_gap > pw * 0.3:
            return "right"
    return "left"


# --------------------------------------------------------------------------- #
# columnas
# --------------------------------------------------------------------------- #
def _split_columns(blocks: list[dict], page_w: float) -> list[list[dict]] | None:
    """Detecta maqueta a dos columnas mediante un canal vertical vacío."""
    if len(blocks) < 6:
        return None
    spans = sorted((b["bbox"][0], b["bbox"][2]) for b in blocks)
    # busca un hueco vertical a ambos lados del centro
    best = None
    for x in range(int(page_w * 0.35), int(page_w * 0.65)):
        left = [b for b in blocks if b["bbox"][2] <= x + 1]
        right = [b for b in blocks if b["bbox"][0] >= x - 1]
        crossing = [b for b in blocks if b["bbox"][0] < x - 1 and b["bbox"][2] > x + 1]
        if len(left) >= 3 and len(right) >= 3 and not crossing:
            score = len(left) + len(right)
            if best is None or score > best[0]:
                best = (score, x, left, right)
    if not best:
        return None
    _, x, left, right = best
    left = sorted(left, key=lambda b: (round(b["bbox"][1], 1), b["bbox"][0]))
    right = sorted(right, key=lambda b: (round(b["bbox"][1], 1), b["bbox"][0]))
    return [left, right]


# --------------------------------------------------------------------------- #
# página
# --------------------------------------------------------------------------- #
def _rgb_hex(fill) -> str:
    try:
        r, g, b = fill
        return "#%02x%02x%02x" % (round(r * 255), round(g * 255), round(b * 255))
    except Exception:
        return ""


def _page_blocks(page, doc, body_size: float, heading_sizes: dict, warn: list[str], pno: int):
    page_rect = page.rect
    try:
        links = [l for l in page.get_links() if l.get("uri")]
    except Exception:
        links = []
    try:
        fills = []
        for dr in page.get_drawings():
            if dr.get("fill"):
                r = fitz.Rect(dr["rect"])
                if abs(r) > 200:
                    fills.append((r, _rgb_hex(dr["fill"])))
    except Exception:
        fills = []
    tables = []
    try:
        finder = page.find_tables()
        tables = list(finder.tables)
    except Exception:
        tables = []
    table_boxes = [fitz.Rect(t.bbox) for t in tables]

    d = page.get_text("dict")
    text_blocks, image_blocks = [], []
    for blk in d.get("blocks", []):
        bbox = fitz.Rect(blk["bbox"])
        if any(_contains(bbox, tb) > 0.6 for tb in table_boxes):
            continue
        if blk.get("type") == 1:
            ib = _image_block(page, blk["bbox"], doc, blk.get("image"), blk.get("ext", ""))
            if ib:
                image_blocks.append(ib)
        else:
            if not blk.get("lines"):
                continue
            text_blocks.append(
                {
                    "type": "text",
                    "bbox": [round(v, 2) for v in blk["bbox"]],
                    "lines": blk["lines"],
                }
            )

    raw: list[dict] = []
    for tb in text_blocks:
        raw.extend(_group_lines(tb["lines"], body_size, heading_sizes, links))
    for t in tables:
        b = _table_to_block(t, page)
        if b:
            raw.append(b)
    raw.extend(image_blocks)

    raw.sort(key=lambda b: (round(b["bbox"][1], 1), round(b["bbox"][0], 1)))

    cols = _split_columns([b for b in raw if b["type"] in ("p", "heading", "li")], page_rect.width)
    if cols:
        warn.append(
            f"Página {pno}: se ha detectado una maqueta a varias columnas; "
            "las columnas se han reconstruido en orden de lectura (la columna "
            "izquierda completa y después la derecha)."
        )
        ordered = []
        for col in cols:
            ordered.extend(col)
        # añade tablas e imágenes por posición vertical mezclándolos
        ordered.extend([b for b in raw if b["type"] in ("table", "image")])
        ordered.sort(key=lambda b: (round(b["bbox"][1], 1), round(b["bbox"][0], 1)))
        raw = ordered

    # alineaciones (+ justificado: varias líneas que llenan el ancho de la columna)
    text_widths = [b["bbox"][2] - b["bbox"][0] for b in raw if b["type"] in ("p", "li")]
    content_w = max(text_widths) if text_widths else 0
    for b in raw:
        if b["type"] in ("p", "heading"):
            b["align"] = _align_with_page({"bbox": b["bbox"]}, page_rect)
        if b["type"] == "p" and content_w:
            bw = b["bbox"][2] - b["bbox"][0]
            bh = b["bbox"][3] - b["bbox"][1]
            if b["align"] == "left" and bw >= content_w * 0.93 and bh >= 34:
                b["align"] = "justify"

    # fondo de color tras el texto (cajas/recuadros de color del PDF)
    for b in raw:
        if b["type"] not in ("p", "heading", "li"):
            continue
        bb = fitz.Rect(b["bbox"])
        best = None
        for r, hexcol in fills:
            if not hexcol or hexcol in ("#ffffff", "#fefefe"):
                continue
            if _contains(bb, r) >= 0.9 and abs(r) <= abs(bb) * 2.6 and abs(bb) > 0:
                if best is None or abs(r) < abs(best[0]):
                    best = (r, hexcol)
        if best:
            b["bg"] = best[1]

    return raw


def _contains(inner, outer) -> float:
    """Fracción del área de ``inner`` contenida en ``outer``."""
    i = fitz.Rect(inner) & fitz.Rect(outer)
    if i.is_empty:
        return 0.0
    a = abs(fitz.Rect(inner))
    return abs(i) / a if a else 0.0


# --------------------------------------------------------------------------- #
# encabezados y pies
# --------------------------------------------------------------------------- #
def _header_footer(doc) -> tuple[dict | None, dict | None, set[int]]:
    """Devuelve (header, footer, páginas_afectadas) detectando texto repetido."""
    tops: dict[str, list] = collections.defaultdict(list)
    bottoms: dict[str, list] = collections.defaultdict(list)
    top_bbox: dict[str, tuple] = {}
    bot_bbox: dict[str, tuple] = {}
    for i, page in enumerate(doc):
        h = page.rect.height
        d = page.get_text("dict")
        for blk in d.get("blocks", []):
            if blk.get("type") != 0:
                continue
            y0, y1 = blk["bbox"][1], blk["bbox"][3]
            text = " ".join(
                "".join(sp.get("text", "") for sp in ln.get("spans", []))
                for ln in blk.get("lines", [])
            ).strip()
            if not text:
                continue
            key = _norm_for_repeat(text)
            if len(key) < 2:
                continue
            if y1 < h * 0.10:
                tops[key].append(i)
                top_bbox.setdefault(key, (text, blk["bbox"]))
            elif y0 > h * 0.90:
                bottoms[key].append(i)
                bot_bbox.setdefault(key, (text, blk["bbox"]))

    n = len(doc)
    threshold = 2 if n >= 3 else 1

    def pick(d: dict, boxes: dict):
        cands = [(k, v) for k, v in d.items() if len(set(v)) >= threshold]
        if not cands:
            return None, set()
        cands.sort(key=lambda kv: -len(set(kv[1])))
        key, pages = cands[0]
        text, bbox = boxes[key]
        # varias líneas sueltas en la misma zona -> únelas por orden vertical
        group = sorted(
            [(boxes[k][1], boxes[k][0]) for k, v in d.items() if len(set(v)) >= threshold],
            key=lambda t: t[0][1],
        )
        text = " · ".join(t for _, t in group[:3])
        return {"text": text, "bbox": bbox, "pages": sorted(set(pages))}, set(pages)

    hdr, hp = pick(tops, top_bbox)
    ftr, fp = pick(bottoms, bot_bbox)
    for info in (hdr, ftr):
        if info:
            _hf_style(doc, info)
    used = hp | fp
    return hdr, ftr, used


def _hf_style(doc, info: dict):
    """Añade color del texto y color de fondo (banda) del encabezado/pie."""
    try:
        page = doc[info["pages"][0]]
        target = fitz.Rect(info["bbox"])
        colors = collections.Counter()
        d = page.get_text("dict")
        for blk in d.get("blocks", []):
            if blk.get("type") != 0:
                continue
            br = fitz.Rect(blk["bbox"])
            cx, cy = (br.x0 + br.x1) / 2, (br.y0 + br.y1) / 2
            if not (target.x0 - 8 <= cx <= target.x1 + 8 and target.y0 - 8 <= cy <= target.y1 + 8):
                continue
            for ln in blk.get("lines", []):
                for sp in ln.get("spans", []):
                    if sp.get("text", "").strip():
                        colors[_color_hex(sp.get("color"))] += len(sp["text"])
        if colors:
            info["color"] = colors.most_common(1)[0][0]
        best_cov = 0.0
        for dr in page.get_drawings():
            if not dr.get("fill"):
                continue
            r = fitz.Rect(dr["rect"])
            inter = r & target
            cov = abs(inter) / abs(target) if abs(target) else 0
            if cov > 0.35 and cov > best_cov:
                hexcol = _rgb_hex(dr["fill"])
                if hexcol:
                    best_cov, info["fill"] = cov, hexcol
    except Exception:
        pass


# --------------------------------------------------------------------------- #
# HTML
# --------------------------------------------------------------------------- #
def _blocks_to_html(blocks: list[dict], i: int = 0) -> str:
    out: list[str] = []
    i = 0
    n = len(blocks)
    while i < n:
        b = blocks[i]
        t = b["type"]
        links = b.get("_links")
        if t == "heading":
            lvl = min(4, max(1, b["level"]))
            styles = []
            if b["align"] != "left":
                styles.append(f'text-align:{b["align"]}')
            if b.get("bg"):
                styles.append(f'background:{b["bg"]}')
            style = f' style="{";".join(styles)}"' if styles else ""
            cls = ' class="callout"' if b.get("bg") else ""
            out.append(f'<h{lvl}{cls}{style}>{_runs_to_html(b["runs"], links)}</h{lvl}>')
            i += 1
        elif t == "p":
            styles = []
            if b["align"] != "left":
                styles.append(f'text-align:{b["align"]}')
            if b.get("bg"):
                styles.append(f'background:{b["bg"]}')
            style = f' style="{";".join(styles)}"' if styles else ""
            cls = ' class="callout"' if b.get("bg") else ""
            out.append(f'<p{cls}{style}>{_runs_to_html(b["runs"], links)}</p>')
            i += 1
        elif t == "li":
            ordered = b["ordered"]
            tag = "ol" if ordered else "ul"
            out.append(f"<{tag}>")
            while i < n and blocks[i]["type"] == "li" and blocks[i]["ordered"] == ordered:
                out.append(f'<li>{_runs_to_html(blocks[i]["runs"], links)}</li>')
                i += 1
            out.append(f"</{tag}>")
        elif t == "table":
            out.append('<table class="doc-table">')
            rows = b["rows"]
            thstyle = f' style="background:{b["header_fill"]}"' if b.get("header_fill") else ""
            if b.get("header") and rows:
                out.append("<thead><tr>" + "".join(
                    f"<th{thstyle}>{html.escape(c)}</th>" for c in rows[0]) + "</tr></thead>")
                body = rows[1:]
            else:
                body = rows
            out.append("<tbody>")
            for r in body:
                out.append("<tr>" + "".join(f"<td>{html.escape(c)}</td>" for c in r) + "</tr>")
            out.append("</tbody></table>")
            i += 1
        elif t == "image":
            out.append(
                f'<figure><img src="{b["src"]}" alt="{html.escape(b.get("alt", ""))}" /></figure>'
            )
            i += 1
        else:
            i += 1
    return "".join(out)


def _ocr_page_blocks(page) -> list[dict]:
    """OCR con Tesseract (si está instalado) sobre una página escaneada."""
    if not shutil.which("tesseract"):
        return []
    try:
        pix = page.get_pixmap(dpi=300)
        with tempfile.TemporaryDirectory() as td:
            png = os.path.join(td, "page.png")
            pix.save(png)
            proc = subprocess.run(
                ["tesseract", png, "stdout", "-l", "spa+eng", "--psm", "6"],
                capture_output=True, text=True, timeout=90,
            )
        text = (proc.stdout or "").strip()
    except Exception:
        return []
    if len(text) < 20:
        return []
    blocks: list[dict] = []
    for para in re.split(r"\n\s*\n", text):
        clean = " ".join(para.split())
        if not clean:
            continue
        run = {"t": clean, "b": False, "i": False, "size": 11.0, "font": "Calibri",
               "color": "#000000", "sup": False, "mono": False}
        blocks.append({"type": "p", "runs": [run], "text": clean, "align": "left",
                       "bbox": [0, 0, 0, 0]})
    return blocks


# --------------------------------------------------------------------------- #
# API principal
# --------------------------------------------------------------------------- #
def analyze(pdf_bytes: bytes, filename: str = "documento.pdf") -> dict:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    warnings: list[str] = []
    if doc.needs_pass:
        raise ValueError("El PDF está protegido con contraseña.")

    # fuentes y tamaños globales
    sizes = collections.Counter()
    fonts = collections.Counter()
    total_chars = 0
    empty_pages: list[int] = []
    for pno, page in enumerate(doc, 1):
        d = page.get_text("dict")
        page_chars = 0
        for blk in d.get("blocks", []):
            if blk.get("type") != 0:
                continue
            for ln in blk.get("lines", []):
                for sp in ln.get("spans", []):
                    txt = sp.get("text", "").strip()
                    if not txt:
                        continue
                    page_chars += len(txt)
                    sizes[round(float(sp.get("size", 0)), 1)] += len(txt)
                    fonts[_clean_font(sp.get("font", ""))] += len(txt)
        total_chars += page_chars
        if page_chars < 12 and page.get_images():
            empty_pages.append(pno)

    body_size = sizes.most_common(1)[0][0] if sizes else 11.0
    body_font = fonts.most_common(1)[0][0] if fonts else "Calibri"

    # tamaños considerados títulos
    bigger = sorted(
        [s for s, c in sizes.items() if s >= body_size * 1.09 and c > 0], reverse=True
    )[:4]
    heading_sizes = {i + 1: s for i, s in enumerate(bigger)}

    hdr, ftr, hf_pages = _header_footer(doc)

    pages = []
    for pno, page in enumerate(doc, 1):
        rect = page.rect
        blocks = _page_blocks(page, doc, body_size, heading_sizes, warnings, pno)
        blocks = [
            b for b in blocks
            if not (b["bbox"][3] < rect.height * 0.10 or b["bbox"][1] > rect.height * 0.90)
        ]
        pages.append(
            {
                "n": pno,
                "w": round(rect.width, 2),
                "h": round(rect.height, 2),
                "blocks": blocks,
            }
        )

    # --- OCR opcional para páginas escaneadas -------------------------------
    ocr_pages: list[int] = []
    if empty_pages and shutil.which("tesseract"):
        for pno in empty_pages:
            ocr_blocks = _ocr_page_blocks(doc[pno - 1])
            if ocr_blocks:
                pages[pno - 1]["blocks"] = ocr_blocks
                ocr_pages.append(pno)
                total_chars += sum(len(b.get("text", "")) for b in ocr_blocks)
    remaining_scanned = [p for p in empty_pages if p not in ocr_pages]
    if ocr_pages:
        warnings.insert(
            0,
            "Se ha aplicado OCR (Tesseract) a las páginas "
            + ", ".join(map(str, ocr_pages))
            + ". Revisa el texto en el editor: el OCR puede contener errores.",
        )
    if remaining_scanned:
        warnings.insert(
            0,
            "Páginas sin capa de texto (probablemente escaneadas): "
            + ", ".join(map(str, remaining_scanned))
            + ". Se han conservado como imagen; para extraer texto editable hace falta OCR "
            "(instala Tesseract en el servidor para activarlo).",
        )

    n_tables = sum(1 for p in pages for b in p["blocks"] if b["type"] == "table")
    n_images = sum(1 for p in pages for b in p["blocks"] if b["type"] == "image")
    n_head = sum(1 for p in pages for b in p["blocks"] if b["type"] == "heading")
    if not total_chars and not n_images:
        warnings.append("No se ha podido extraer contenido de este PDF.")

    meta = {
        "filename": filename,
        "pages": len(doc),
        "page_width": pages[0]["w"] if pages else 595,
        "page_height": pages[0]["h"] if pages else 842,
        "landscape": bool(pages and pages[0]["w"] > pages[0]["h"]),
        "body_font": body_font,
        "body_size": body_size,
        "heading_sizes": {str(k): v for k, v in heading_sizes.items()},
        "header": hdr,
        "footer": ftr,
        "chars": total_chars,
        "tables": n_tables,
        "images": n_images,
        "headings": n_head,
        "encrypted": bool(doc.needs_pass),
        "outline": [list(t) for t in doc.get_toc()],
    }

    sections = []
    for p in pages:
        sections.append(
            f'<section class="pdfpage" data-page="{p["n"]}" '
            f'style="--pw:{p["w"]}px;--ph:{p["h"]}px">'
            + _blocks_to_html(p["blocks"])
            + "</section>"
        )
    doc.close()

    return {
        "html": "\n".join(sections),
        "meta": meta,
        "warnings": warnings,
        "stats": {
            "pages": meta["pages"],
            "chars": total_chars,
            "tables": n_tables,
            "images": n_images,
            "headings": n_head,
            "scanned_pages": remaining_scanned,
            "ocr_pages": ocr_pages,
        },
    }
