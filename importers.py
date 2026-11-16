"""Importadores MD / DOCX / ODT -> misma representación HTML editable que PDF."""
from __future__ import annotations

import base64
import html as _html
import io
import re
import zipfile
import xml.etree.ElementTree as ET

BASE_META = {
    "pages": 1, "page_width": 595.28, "page_height": 841.89, "landscape": False,
    "body_font": "Calibri", "body_size": 11.0, "heading_sizes": {},
    "header": None, "footer": None,
}


def _result(html_body: str, filename: str, fmt: str, warnings=None, extra_stats=None):
    meta = dict(BASE_META)
    meta.update({"filename": filename, "source_format": fmt, "encrypted": False})
    text_len = len(re.sub(r"<[^>]+>", "", html_body))
    stats = {
        "pages": 1, "chars": text_len,
        "tables": html_body.count("<table"),
        "images": html_body.count("<img"),
        "headings": len(re.findall(r"<h[1-4]", html_body)),
        "scanned_pages": [], "ocr_pages": [],
    }
    if extra_stats:
        stats.update(extra_stats)
    meta.update({"chars": stats["chars"], "tables": stats["tables"],
                 "images": stats["images"], "headings": stats["headings"]})
    html_doc = f'<section class="pdfpage" data-page="1">{html_body}</section>'
    return {"html": html_doc, "meta": meta, "warnings": warnings or [], "stats": stats}


def _img_tag(data: bytes, ext: str = "png") -> str:
    mime = "image/jpeg" if ext.lower() in ("jpg", "jpeg") else f"image/{ext.lower()}"
    return f'<figure><img src="data:{mime};base64,{base64.b64encode(data).decode()}" alt=""/></figure>'


# --------------------------------------------------------------------------- #
# Markdown
# --------------------------------------------------------------------------- #
def import_markdown(data_bytes: bytes, filename: str) -> dict:
    from markdown_it import MarkdownIt

    text = data_bytes.decode("utf-8", errors="replace")
    md = MarkdownIt("commonmark", {"html": True, "linkify": True}).enable("table").enable("strikethrough")
    body = md.render(text)
    warnings = []
    if re.search(r"!\[[^\]]*\]\((?!https?://|data:)[^)]+\)", text):
        warnings.append("El Markdown referencia imágenes externas o locales que no se han subido: "
                        "se conserva la referencia, pero la imagen no aparecerá hasta que la insertes en el documento final.")
    return _result(body, filename, "md", warnings)


# --------------------------------------------------------------------------- #
# DOCX
# --------------------------------------------------------------------------- #
def import_docx(data_bytes: bytes, filename: str) -> dict:
    from docx import Document
    from docx.oxml.ns import qn as DQ

    doc = Document(io.BytesIO(data_bytes))
    out: list[str] = []
    warnings: list[str] = []

    def run_html(r) -> str:
        t = _html.escape(r.text or "")
        if not t:
            return ""
        if r.font.superscript:
            t = f"<sup>{t}</sup>"
        if r.font.name and r.font.name not in ("Calibri", "Cambria"):
            pass  # el nombre se conserva vía estilo del párrafo en Word al exportar
        try:
            if r.font.color and r.font.color.rgb is not None:
                t = f'<span style="color:#{str(r.font.color.rgb).lower()}">{t}</span>'
        except Exception:
            pass
        if r.bold:
            t = f"<strong>{t}</strong>"
        if r.italic:
            t = f"<em>{t}</em>"
        if r.underline:
            t = f"<u>{t}</u>"
        return t

    def para_html(p) -> str:
        style = (p.style.name or "") if p.style else ""
        align = ""
        try:
            from docx.enum.text import WD_ALIGN_PARAGRAPH as WA
            align = {WA.CENTER: "center", WA.RIGHT: "right", WA.JUSTIFY: "justify"}.get(p.alignment, "")
        except Exception:
            align = ""
        st = f' style="text-align:{align}"' if align else ""
        m = re.match(r"(?:Heading|Título|Encabezado)\s*(\d)", style, re.I)
        # imágenes incrustadas en el párrafo
        imgs = []
        for blip in p._p.iter():
            if blip.tag.endswith("}blip"):
                rid = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                if rid and rid in doc.part.related_parts:
                    part = doc.part.related_parts[rid]
                    ext = part.partname.ext if hasattr(part, "partname") else "png"
                    imgs.append(_img_tag(part._blob, ext))
        text = "".join(run_html(r) for r in p.runs)
        # hipervínculos (conserva la URL)
        for hl in p._p.findall(DQ("w:hyperlink")):
            rid = hl.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            url = ""
            if rid and rid in doc.part.rels:
                url = doc.part.rels[rid].target_ref
            txt = "".join(node.text or "" for node in hl.iter() if node.tag.endswith("}t"))
            if txt:
                text += f'<a href="{_html.escape(url, quote=True)}">{_html.escape(txt)}</a>'
        if m:
            lvl = min(4, max(1, int(m.group(1))))
            return (f"<h{lvl}{st}>{text}</h{lvl}>" if text else "") + "".join(imgs)
        if style.lower().startswith(("list bullet", "viñeta", "lista con viñetas")):
            return f"<ul><li>{text}</li></ul>" + "".join(imgs)
        if style.lower().startswith(("list number", "lista numerada", "list paragraph")):
            return f"<ol><li>{text}</li></ol>" + "".join(imgs)
        if not text and imgs:
            return "".join(imgs)
        if not text:
            return ""
        return f"<p{st}>{text}</p>" + "".join(imgs)

    for child in doc.element.body.iterchildren():
        if child.tag.endswith("}p"):
            from docx.text.paragraph import Paragraph
            out.append(para_html(Paragraph(child, doc)))
        elif child.tag.endswith("}tbl"):
            from docx.table import Table
            tbl = Table(child, doc)
            rows_html = []
            for ri, row in enumerate(tbl.rows):
                tag = "th" if ri == 0 else "td"
                cells = "".join(f"<{tag}>{_html.escape(c.text.strip())}</{tag}>" for c in row.cells)
                rows_html.append(f"<tr>{cells}</tr>")
            out.append('<table class="doc-table">' + "".join(rows_html) + "</table>")

    try:
        sec = doc.sections[0]
        htxt = " ".join(p.text for p in sec.header.paragraphs).strip()
        ftxt = " ".join(p.text for p in sec.footer.paragraphs).strip()
    except Exception:
        htxt = ftxt = ""
    res = _result("".join(out), filename, "docx", warnings)
    if htxt:
        res["meta"]["header"] = {"text": htxt}
    if ftxt:
        res["meta"]["footer"] = {"text": ftxt}
    try:
        res["meta"]["body_font"] = doc.styles["Normal"].font.name or "Calibri"
        if doc.styles["Normal"].font.size:
            res["meta"]["body_size"] = doc.styles["Normal"].font.size.pt
    except Exception:
        pass
    return res


# --------------------------------------------------------------------------- #
# ODT
# --------------------------------------------------------------------------- #
NS = {
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    "xlink": "http://www.w3.org/1999/xlink",
}


def _q(ns: str, tag: str) -> str:
    return f"{{{NS[ns]}}}{tag}"


def import_odt(data_bytes: bytes, filename: str) -> dict:
    z = zipfile.ZipFile(io.BytesIO(data_bytes))
    content = ET.fromstring(z.read("content.xml"))
    styles_root = None
    if "styles.xml" in z.namelist():
        try:
            styles_root = ET.fromstring(z.read("styles.xml"))
        except Exception:
            styles_root = None

    # mapa de estilos -> propiedades de texto/párrafo
    props: dict[str, dict] = {}

    def collect(root):
        if root is None:
            return
        for st in root.iter(_q("style", "style")):
            name = st.get(_q("style", "name"))
            if not name:
                continue
            d: dict = {}
            tp = st.find(_q("style", "text-properties"))
            if tp is not None:
                if tp.get(_q("style", "font-weight")) == "bold" or tp.get(_q("fo", "font-weight")) == "bold":
                    d["b"] = True
                if tp.get(_q("style", "font-style")) == "italic" or tp.get(_q("fo", "font-style")) == "italic":
                    d["i"] = True
                color = tp.get(_q("fo", "color"))
                if color:
                    d["color"] = color
            pp = st.find(_q("style", "paragraph-properties"))
            if pp is not None:
                ta = pp.get(_q("fo", "text-align"))
                if ta:
                    d["align"] = {"center": "center", "end": "right", "justify": "justify",
                                  "start": "left"}.get(ta, "")
                bg = pp.get(_q("fo", "background-color"))
                if bg and bg != "transparent":
                    d["bg"] = bg
            parent = st.get(_q("style", "parent-style-name"))
            if parent and parent in props:
                merged = dict(props[parent])
                merged.update(d)
                d = merged
            props[name] = d

    collect(styles_root)
    auto = content.find(_q("office", "automatic-styles"))
    if auto is not None:
        collect(auto)

    def text_of(el) -> str:
        parts = []
        for node in el.iter():
            if node.tag in (_q("text", "p"), _q("text", "h")) and node is not el:
                continue
            if node.text:
                parts.append(node.text)
            if node.tag == _q("text", "s"):
                parts.append(" ")
            if node.tail:
                parts.append(node.tail)
        return "".join(parts)

    def inline_of(p_el) -> str:
        out: list[str] = []

        def walk(el, fmt: dict):
            for child in el:
                f = dict(fmt)
                if child.tag == _q("text", "span"):
                    f.update(props.get(child.get(_q("text", "style-name")), {}))
                    if child.text:
                        out.append(_fmt(child.text, f))
                    walk(child, f)
                elif child.tag == _q("text", "a"):
                    href = child.get(_q("xlink", "href")) or ""
                    txt = _html.escape(text_of(child))
                    out.append(f'<a href="{_html.escape(href, quote=True)}">{txt}</a>' if href else txt)
                elif child.tag == _q("text", "s"):
                    out.append(" ")
                elif child.tag == _q("text", "line-break"):
                    out.append("<br/>")
                else:
                    if child.text:
                        out.append(_fmt(child.text, f))
                    walk(child, f)
                if child.tail:
                    out.append(_fmt(child.tail, fmt))

        if p_el.text:
            out.append(_fmt(p_el.text, {}))
        walk(p_el, {})
        return "".join(out)

    def _fmt(txt: str, f: dict) -> str:
        t = _html.escape(txt)
        if f.get("color") and f["color"] != "#000000":
            t = f'<span style="color:{f["color"]}">{t}</span>'
        if f.get("b"):
            t = f"<strong>{t}</strong>"
        if f.get("i"):
            t = f"<em>{t}</em>"
        return t

    def para_block(el, level: int = 0) -> str:
        stname = el.get(_q("text", "style-name")) or ""
        pr = props.get(stname, {})
        body = inline_of(el)
        imgs = []
        for fr in el.iter(_q("draw", "image")):
            href = fr.get(_q("xlink", "href")) or ""
            if href in z.namelist():
                ext = href.rsplit(".", 1)[-1]
                imgs.append(_img_tag(z.read(href), ext))
        style_bits = []
        if pr.get("align"):
            style_bits.append(f'text-align:{pr["align"]}')
        if pr.get("bg"):
            style_bits.append(f'background:{pr["bg"]}')
        st = f' style="{";".join(style_bits)}"' if style_bits else ""
        cls = ' class="callout"' if pr.get("bg") else ""
        if level:
            return (f"<h{level}{st}>{body}</h{level}>" if body else "") + "".join(imgs)
        if not body and imgs:
            return "".join(imgs)
        if not body:
            return ""
        return f"<p{cls}{st}>{body}</p>" + "".join(imgs)

    out: list[str] = []
    body_el = content.find(f"{_q('office', 'body')}/{_q('office', 'text')}")

    def walk_body(el):
        for child in el:
            if child.tag == _q("text", "h"):
                lvl = int(child.get(_q("text", "outline-level") or "1"))
                out.append(para_block(child, min(4, max(1, lvl))))
            elif child.tag == _q("text", "p"):
                out.append(para_block(child))
            elif child.tag == _q("text", "list"):
                ordered = "number" in (child.get(_q("text", "style-name")) or "").lower()
                tag = "ol" if ordered else "ul"
                items = []
                for item in child.findall(_q("text", "list-item")):
                    txt = "".join(inline_of(p) for p in item.findall(_q("text", "p")))
                    if txt:
                        items.append(f"<li>{txt}</li>")
                out.append(f"<{tag}>" + "".join(items) + f"</{tag}>")
            elif child.tag == _q("table", "table"):
                rows = []
                for ri, row in enumerate(child.findall(_q("table", "table-row"))):
                    tag = "th" if ri == 0 else "td"
                    cells = []
                    for cell in row.findall(_q("table", "table-cell")):
                        txt = " ".join(text_of(p) for p in cell.findall(_q("text", "p"))).strip()
                        cells.append(f"<{tag}>{_html.escape(txt)}</{tag}>")
                    rows.append("<tr>" + "".join(cells) + "</tr>")
                out.append('<table class="doc-table">' + "".join(rows) + "</table>")
            else:
                walk_body(child)

    if body_el is not None:
        walk_body(body_el)
    return _result("".join(out), filename, "odt",
                   ["ODT: se ha reconstruido el contenido como en la vista previa; estilos muy específicos de LibreOffice pueden simplificarse."])
