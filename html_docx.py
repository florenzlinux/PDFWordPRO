"""
HTML (editado en la vista previa) -> DOCX real y editable con python-docx.

Se reconstruyen títulos, párrafos, listas, tablas, imágenes, hipervínculos,
estilos inline, alineación, saltos de página y encabezado/pie de página.
"""
from __future__ import annotations

import base64
import io
import re
from html.parser import HTMLParser

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu, Pt, RGBColor, Twips

VOID = {"img", "br", "hr", "meta", "link", "input"}


class Node:
    __slots__ = ("tag", "attrs", "children", "text")

    def __init__(self, tag=None, attrs=None, text=None):
        self.tag = tag
        self.attrs = dict(attrs or {})
        self.children: list[Node] = []
        self.text = text

    def iter(self):
        yield self
        for c in self.children:
            yield from c.iter()

    def find_all(self, tag):
        return [n for n in self.iter() if n.tag == tag]


class TreeBuilder(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root = Node("root")
        self.stack = [self.root]

    def handle_starttag(self, tag, attrs):
        node = Node(tag, attrs)
        self.stack[-1].children.append(node)
        if tag not in VOID:
            self.stack.append(node)

    def handle_startendtag(self, tag, attrs):
        self.stack[-1].children.append(Node(tag, attrs))

    def handle_endtag(self, tag):
        for i in range(len(self.stack) - 1, 0, -1):
            if self.stack[i].tag == tag:
                del self.stack[i:]
                return

    def handle_data(self, data):
        if data:
            self.stack[-1].children.append(Node(text=data))


def parse_html(html: str) -> Node:
    tb = TreeBuilder()
    tb.feed(html or "")
    tb.close()
    return tb.root


# --------------------------------------------------------------------------- #
# utilidades XML
# --------------------------------------------------------------------------- #
def _add_hyperlink(paragraph, url: str, runs_xml: list):
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    for el in runs_xml:
        hyperlink.append(el)
    paragraph._p.append(hyperlink)


def _field(paragraph, instr: str):
    run = paragraph.add_run()
    run.font.size = Pt(9)
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = instr
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(it)
    run._r.append(fld_end)


def _para_shade(paragraph, hexcolor: str):
    try:
        pPr = paragraph._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), hexcolor)
        pPr.append(shd)
    except Exception:
        pass


def _shade(cell, hexcolor: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)


def _set_cell_margins(table, twips=80):
    tblPr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(twips))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)


def _img_bytes(src: str):
    m = re.match(r"data:image/[^;]+;base64,(.*)$", src or "", re.S)
    if not m:
        return None
    try:
        return base64.b64decode(m.group(1))
    except Exception:
        return None


ALIGN = {
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "left": WD_ALIGN_PARAGRAPH.LEFT,
}


def _align_of(node: Node):
    st = node.attrs.get("style", "") or ""
    m = re.search(r"text-align\s*:\s*(\w+)", st)
    al = (node.attrs.get("align") or (m.group(1) if m else "")).lower()
    return ALIGN.get(al)


def _style_map(node: Node) -> dict:
    st = node.attrs.get("style", "") or ""
    out = {}
    m = re.search(r"color\s*:\s*#([0-9a-fA-F]{6})", st)
    if m:
        out["color"] = m.group(1)
    m = re.search(r"font-size\s*:\s*([\d.]+)pt", st)
    if m:
        out["size"] = float(m.group(1))
    m = re.search(r"font-family\s*:\s*([^;]+)", st)
    if m:
        out["font"] = m.group(1).split(",")[0].strip().strip("'\"")
    m = re.search(r"background(?:-color)?\s*:\s*#([0-9a-fA-F]{6})", st)
    if m:
        out["bg"] = m.group(1)
    return out


# --------------------------------------------------------------------------- #
# conversión principal
# --------------------------------------------------------------------------- #
def html_to_docx(html: str, meta: dict | None = None, title: str = "") -> bytes:
    meta = meta or {}
    root = parse_html(html)
    doc = Document()

    # ---- geometría de página ------------------------------------------------
    sec = doc.sections[0]
    pw = float(meta.get("page_width") or 595.28)
    ph = float(meta.get("page_height") or 841.89)
    sec.page_width = Emu(int(pw * 12700))
    sec.page_height = Emu(int(ph * 12700))
    if pw > ph:
        sec.orientation = WD_ORIENT.LANDSCAPE
    else:
        sec.orientation = WD_ORIENT.PORTRAIT
    margin = float(meta.get("margin_pt") or 56.7)  # ~2 cm
    sec.left_margin = sec.right_margin = Emu(int(margin * 12700))
    sec.top_margin = sec.bottom_margin = Emu(int(margin * 12700))

    body_font = meta.get("body_font") or "Calibri"
    body_size = float(meta.get("body_size") or 11)
    normal = doc.styles["Normal"]
    normal.font.name = body_font
    normal.font.size = Pt(max(6, min(28, body_size)))
    rpr = normal.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(a), body_font)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.08

    # estilos de título escalados al documento original
    hs = meta.get("heading_sizes") or {}
    for lvl in range(1, 5):
        size = hs.get(str(lvl))
        try:
            size = float(size) if size else [0, 22, 17, 14, 12][lvl]
        except Exception:
            size = [0, 22, 17, 14, 12][lvl]
        st = doc.styles[f"Heading {lvl}"]
        st.font.name = body_font
        st.font.size = Pt(max(9, min(40, size)))
        st.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)
        st.paragraph_format.space_before = Pt(12 if lvl <= 2 else 8)
        st.paragraph_format.space_after = Pt(6)
        st.paragraph_format.keep_with_next = True

    # ---- encabezado / pie --------------------------------------------------
    hdr = meta.get("header")
    ftr = meta.get("footer")
    if hdr and (hdr.get("text") or "").strip():
        p = sec.header.paragraphs[0]
        p.text = ""
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if hdr.get("fill"):
            _para_shade(p, hdr["fill"].lstrip("#"))
        r = p.add_run(_clean_hf(hdr["text"]))
        r.font.size = Pt(9)
        r.bold = bool(hdr.get("fill"))
        if hdr.get("color"):
            try:
                r.font.color.rgb = RGBColor.from_string(hdr["color"].lstrip("#").upper())
            except Exception:
                pass
    if ftr and (ftr.get("text") or "").strip():
        p = sec.footer.paragraphs[0]
        p.text = ""
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fcolor = (ftr.get("color") or "").lstrip("#")
        if _has_pagenum(ftr.get("text", "")):
            _footer_pagenum(p, ftr.get("text", ""), fcolor)
        else:
            r = p.add_run(_clean_hf(ftr["text"]))
            r.font.size = Pt(9)
            if fcolor:
                try:
                    r.font.color.rgb = RGBColor.from_string(fcolor.upper())
                except Exception:
                    pass

    # ---- contenido ---------------------------------------------------------
    state = {"first_section": True}
    _walk_children(root, doc, {}, state, meta)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _clean_hf(text: str) -> str:
    return re.sub(r"\s*·\s*", "  |  ", (text or "").strip())


def _has_pagenum(text: str) -> bool:
    return bool(re.search(r"(p[áa]gina|page|pág\.?|\d+\s*(de|of|/)\s*\d+)", text or "", re.I))


def _style_hf_run(run, color: str = ""):
    run.font.size = Pt(9)
    if color:
        try:
            run.font.color.rgb = RGBColor.from_string(color.upper())
        except Exception:
            pass
    return run


def _footer_pagenum(paragraph, text: str, color: str = ""):
    t = _clean_hf(text)
    # «…título… 4 de 22», «Página 4 de 22», «Página 4», «4 / 22»
    m = re.match(r"^(.*?)(\d+)\s*(de|of|/)\s*(\d+)\s*$", t)
    if m and (m.group(1).strip() or True):
        prefix = m.group(1).strip()
        if prefix:
            _style_hf_run(paragraph.add_run(prefix + " "), color)
        _field(paragraph, "PAGE")
        _style_hf_run(paragraph.add_run(f" {m.group(3)} "), color)
        _field(paragraph, "NUMPAGES")
        return
    m = re.match(r"^(.*?(?:p[áa]gina|page|pág\.?))\s*\d+\s*$", t, re.I)
    if m:
        _style_hf_run(paragraph.add_run(m.group(1).strip() + " "), color)
        _field(paragraph, "PAGE")
        return
    _style_hf_run(paragraph.add_run(t), color)


def _walk_children(node: Node, doc, fmt: dict, state: dict, meta: dict, prefix: str = ""):
    for ch in node.children:
        _walk(ch, doc, fmt, state, meta, prefix)


def _walk(node: Node, doc, fmt: dict, state: dict, meta: dict, prefix: str = ""):
    tag = node.tag
    if tag is None:
        return

    if tag == "section" and "pdfpage" in node.attrs.get("class", ""):
        if not state.get("first_section"):
            doc.add_page_break()
        state["first_section"] = False
        _walk_children(node, doc, fmt, state, meta)
        return

    if re.fullmatch(r"h[1-6]", tag):
        lvl = int(tag[1])
        p = doc.add_paragraph(style=f"Heading {min(lvl, 4)}")
        al = _align_of(node)
        if al is not None:
            p.alignment = al
        _add_inline(node, p, {**fmt, "b": False})
        if not p.runs:
            p.add_run("")
        return

    if tag == "p" or (tag == "div" and "block" in node.attrs.get("class", "")):
        p = doc.add_paragraph()
        al = _align_of(node)
        if al is not None:
            p.alignment = al
        bg = _style_map(node).get("bg")
        if bg:
            _para_shade(p, bg)
        _add_inline(node, p, fmt)
        return

    if tag in ("ul", "ol"):
        items = [c for c in node.children if c.tag == "li"]
        num = 1
        for li in items:
            style = "List Number" if tag == "ol" else "List Bullet"
            p = doc.add_paragraph(style=style)
            _add_inline(li, p, fmt)
            num += 1
        return

    if tag == "table":
        _add_table(node, doc, fmt)
        return

    if tag == "figure" or tag == "img":
        _add_image(node, doc)
        return

    if tag == "hr":
        cls = node.attrs.get("class", "")
        if "pagebreak" in cls:
            doc.add_page_break()
        return

    if tag == "blockquote":
        for child in node.children:
            if child.tag in ("p", None):
                p = doc.add_paragraph(style="Quote" if _has_style(doc, "Quote") else None)
                _add_inline(child if child.tag == "p" else node, p, fmt)
            else:
                _walk(child, doc, fmt, state, meta)
        return

    if tag in ("br",):
        if doc.paragraphs:
            doc.paragraphs[-1].add_run().add_break(WD_BREAK.LINE)
        return

    # contenedores genéricos
    _walk_children(node, doc, fmt, state, meta)


def _has_style(doc, name) -> bool:
    try:
        doc.styles[name]
        return True
    except KeyError:
        return False


def _add_inline(node: Node, paragraph, fmt: dict, hyperlink: str | None = None):
    for ch in node.children:
        if ch.tag is None:
            if ch.text:
                f = dict(fmt)
                run = paragraph.add_run(ch.text)
                _apply_run(run, f)
            continue
        tag = ch.tag
        f = dict(fmt)
        if tag in ("strong", "b"):
            f["b"] = True
        elif tag in ("em", "i"):
            f["i"] = True
        elif tag == "u":
            f["u"] = True
        elif tag in ("s", "del", "strike"):
            f["s"] = True
        elif tag == "code":
            f["mono"] = True
        elif tag == "sup":
            f["sup"] = True
        elif tag == "sub":
            f["sub"] = True
        elif tag == "a":
            href = ch.attrs.get("href", "")
            txt = "".join(c.text or "" for c in ch.iter() if c.tag is None).strip()
            if href and txt:
                r = paragraph.add_run(txt)
                _apply_run(r, f)
                r.font.color.rgb = RGBColor(0x0B, 0x57, 0xD0)
                r.font.underline = True
                _add_hyperlink(paragraph, href, [r._r])
                continue
        elif tag == "br":
            paragraph.add_run().add_break(WD_BREAK.LINE)
            continue
        elif tag == "img":
            # imágenes dentro de un párrafo -> se insertan como párrafo aparte
            _add_image(ch, paragraph.part.document if hasattr(paragraph.part, "document") else None)
            continue
        elif tag in ("span", "font", "mark", "small", "big"):
            f.update(_style_map(ch))
        _add_inline(ch, paragraph, f)


def _apply_run(run, f: dict):
    if f.get("b"):
        run.bold = True
    if f.get("i"):
        run.italic = True
    if f.get("u"):
        run.underline = True
    if f.get("s"):
        run.font.strike = True
    if f.get("mono"):
        run.font.name = "Consolas"
    if f.get("sup"):
        run.font.superscript = True
    if f.get("sub"):
        run.font.subscript = True
    if f.get("color"):
        try:
            run.font.color.rgb = RGBColor.from_string(f["color"].upper())
        except Exception:
            pass
    if f.get("size"):
        run.font.size = Pt(f["size"])
    if f.get("font"):
        run.font.name = f["font"]


def _add_image(node: Node, doc):
    if doc is None:
        return
    imgs = [node] if node.tag == "img" else node.find_all("img")
    for im in imgs:
        data = _img_bytes(im.attrs.get("src", ""))
        if not data:
            continue
        try:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run()
            sec = doc.sections[0]
            max_w = sec.page_width - sec.left_margin - sec.right_margin
            run.add_picture(io.BytesIO(data), width=max_w if _is_wide(data) else None)
        except Exception:
            continue


def _is_wide(data: bytes) -> bool:
    return True


def _add_table(node: Node, doc, fmt: dict):
    rows = []
    for tr in node.find_all("tr"):
        cells = [c for c in tr.children if c.tag in ("td", "th")]
        rows.append(cells)
    if not rows:
        return
    ncols = max(len(r) for r in rows)
    if ncols == 0:
        return
    table = doc.add_table(rows=len(rows), cols=ncols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_cell_margins(table)
    for ri, cells in enumerate(rows):
        for ci in range(ncols):
            cell = table.cell(ri, ci)
            cell.paragraphs[0].text = ""
            if ci < len(cells):
                src = cells[ci]
                is_th = src.tag == "th"
                p = cell.paragraphs[0]
                _add_inline(src, p, {**fmt, "b": is_th or fmt.get("b", False)})
                if is_th:
                    _shade(cell, _style_map(src).get("bg") or "EDEDED")
            if not cell.paragraphs[0].runs:
                cell.paragraphs[0].add_run("")
    # repite la fila de cabecera si la primera fila es de cabecera
    if node.find_all("th"):
        tr = table.rows[0]._tr
        trPr = tr.get_or_add_trPr()
        el = OxmlElement("w:tblHeader")
        el.set(qn("w:val"), "true")
        trPr.append(el)
