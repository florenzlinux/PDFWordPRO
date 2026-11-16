"""Fusión de varios documentos en uno solo (DOCX con marcadores / PDF con outline)."""
from __future__ import annotations

import io
import re

import pymupdf as fitz
from docx import Document
from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .html_docx import html_to_docx


def _slug(text: str, idx: int) -> str:
    s = re.sub(r"[^A-Za-z0-9_]+", "_", (text or "documento")).strip("_")[:40]
    return f"doc_{idx}_{s or 'documento'}"


def _insert_title_and_toc(data: bytes, doc_title: str, file_titles: list[str]) -> bytes:
    """Inserta título + campo Índice al principio y marcadores por documento."""
    doc = Document(io.BytesIO(data))
    body = doc.element.body

    def new_p(text: str, style: str | None = None):
        p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
        if text:
            p.add_run(text)
        return p

    # construir en memoria y luego insertar al principio en orden inverso
    elems = []
    toc_p = new_p("")
    run = toc_p.add_run()
    for instr, txt in (("begin", None), (None, ' TOC \\o "1-3" \\h \\z \\u '), ("separate", None)):
        if instr:
            fld = OxmlElement("w:fldChar")
            fld.set(qn("w:fldCharType"), instr)
            run._r.append(fld)
        else:
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = txt
            run._r.append(it)
    hint = new_p("Índice automático: en Word, clic derecho sobre el índice → «Actualizar campos» para paginarlo.", None)
    elems.append(hint._p)
    elems.append(toc_p._p)
    h = new_p("Índice", "Heading 1")
    elems.append(h._p)
    t = new_p(doc_title, "Title")
    elems.append(t._p)
    for el in elems:  # ya están al final: mover al principio en orden inverso
        body.insert(0, el)

    # marcadores (bookmarks) reales en el título de cada documento fusionado
    wanted = {t.strip(): i for i, t in enumerate(file_titles)}
    next_id = [1000]
    for p in doc.paragraphs:
        key = p.text.strip()
        if key in wanted and p.style.name.startswith("Heading"):
            name = _slug(key, wanted[key] + 1)
            start = OxmlElement("w:bookmarkStart")
            start.set(qn("w:id"), str(next_id[0]))
            start.set(qn("w:name"), name)
            end = OxmlElement("w:bookmarkEnd")
            end.set(qn("w:id"), str(next_id[0]))
            next_id[0] += 1
            p._p.insert(0, start)
            p._p.append(end)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def merged_docx(items: list[dict], doc_title: str = "Documentos refundidos") -> bytes:
    """
    items: [{title, html, meta}] en el orden deseado.
    Cada documento empieza en página nueva encabezado por su título (Heading 1),
    para que Word genere navegación/marcadores y el Índice funcione.
    """
    parts: list[str] = []
    titles: list[str] = []
    for i, it in enumerate(items):
        title = (it.get("title") or f"Documento {i + 1}").strip()
        titles.append(title)
        html = it.get("html") or ""
        # inserta el título de la unidad dentro de su primera sección (página)
        esc = title.replace("&", "&amp;").replace("<", "&lt;")
        h1 = f"<h1>{esc}</h1>"
        m = re.search(r'(<section class="pdfpage"[^>]*>)', html)
        if m:
            html = html[: m.end()] + h1 + html[m.end():]
        else:
            html = f'<section class="pdfpage">{h1}{html}</section>'
        parts.append(html)
    meta0 = dict(items[0].get("meta") or {}) if items else {}
    data = html_to_docx("\n".join(parts), meta0, doc_title)
    try:
        data = _insert_title_and_toc(data, doc_title, titles)
    except Exception:
        pass  # si algo falla, se entrega el refundido sin índice automático
    return data


# --------------------------------------------------------------------------- #
def _html_to_pdf_bytes(html: str) -> bytes:
    """HTML -> PDF con PyMuPDF Story (para fuentes MD/DOCX/ODT en el refundido PDF)."""
    css = """
    @page { size: A4; margin: 2cm; }
    body { font-family: sans-serif; font-size: 10pt; }
    h1 { font-size: 18pt; } h2 { font-size: 14pt; } h3 { font-size: 12pt; }
    table { border-collapse: collapse; } td, th { border: 0.5pt solid #999; padding: 3pt; }
    img { max-width: 100%; } section { }
    """
    clean = re.sub(r'<section class="pdfpage"[^>]*>', "", html or "")
    clean = clean.replace("</section>", "")
    out = fitz.open()
    story = fitz.Story(html=clean, user_css=css)
    page_rect = fitz.paper_rect("a4")
    content_rect = page_rect + (56, 56, -56, -56)
    buf = io.BytesIO()
    writer = fitz.DocumentWriter(buf)
    more = True
    while more:
        device = writer.begin_page(page_rect)
        more, _ = story.place(content_rect)
        story.draw(device)
        writer.end_page()
    writer.close()
    return buf.getvalue()


def merged_pdf(items: list[dict]) -> tuple[bytes, list[str]]:
    """
    items: [{title, kind: 'pdf'|'other', src: bytes|None, html, outline: [(lvl,title,page)]}]
    Conserva los marcadores (outline) de cada PDF y los anida bajo el título de cada archivo.
    """
    out = fitz.open()
    toc: list[list] = []
    notes: list[str] = []
    for it in items:
        title = (it.get("title") or "Documento").strip()
        start = out.page_count + 1  # 1-based
        if it.get("kind") == "pdf" and it.get("src"):
            src = fitz.open(stream=it["src"], filetype="pdf")
            out.insert_pdf(src)
            toc.append([1, title, start])
            for lvl, t, page in (it.get("outline") or []):
                toc.append([min(4, lvl + 1), t, start + page - 1])
            src.close()
        else:
            try:
                gen = _html_to_pdf_bytes(it.get("html") or "")
                src = fitz.open(stream=gen, filetype="pdf")
                out.insert_pdf(src)
                src.close()
                toc.append([1, title, start])
            except Exception as e:  # noqa: BLE001
                notes.append(f"{title}: no se ha podido incluir en el PDF refundido ({e}); sí está en el DOCX refundido.")
    if toc:
        try:
            out.set_toc(toc)
        except Exception:
            pass
    buf = io.BytesIO()
    out.save(buf, garbage=3, deflate=True)
    out.close()
    return buf.getvalue(), notes
