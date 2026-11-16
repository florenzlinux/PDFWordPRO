from __future__ import annotations
import asyncio, io, json, os, re, shutil, threading, traceback, uuid, zipfile
from pathlib import Path
from fastapi import FastAPI, File, UploadFile, HTTPException, Body
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from .extract import analyze
from .html_docx import html_to_docx, parse_html
from .html_md import html_to_markdown
from .importers import import_docx, import_markdown, import_odt
from .layout_docx import convert_layout
from .merge import merged_docx, merged_pdf

BASE = Path(__file__).resolve().parent.parent
# Los datos pueden vivir fuera del directorio de instalación (p. ej. ~/.local/share
# cuando se lanza desde el menú de aplicaciones sin permisos de escritura en /opt).
DATA_ROOT = Path(os.environ.get("PDFWORDPRO_DATA", BASE / "data"))
DATA = DATA_ROOT / "jobs"
DATA.mkdir(parents=True, exist_ok=True)
STATIC = Path(__file__).resolve().parent / "static"

app = FastAPI(title="PDF → Word PRO", version="1.0.0")

# --- modelos ---------------------------------------------------------------
class DocxRequest(BaseModel):
    html: str = ""
    mode: str = "editable"  # editable | fidelidad
    header: str | None = None
    footer: str | None = None
    title: str = ""

class MdRequest(BaseModel):
    html: str = ""
    title: str = ""

class OcrRequest(BaseModel):
    lang: str = "spa"

# --- utilidades ------------------------------------------------------------
def job_dir(jid: str) -> Path:
    d = DATA / jid
    if not d.exists():
        raise HTTPException(404, "Documento no encontrado (la sesión puede haber caducado).")
    return d

def safe_name(name: str) -> str:
    name = Path(name or "documento").stem
    name = re.sub(r"[^\w áéíóúñüÁÉÍÓÚÑÜ-]+", "_", name, flags=re.U).strip()
    return name or "documento"

def load_meta(d: Path) -> dict:
    return json.loads((d / "meta.json").read_text(encoding="utf-8"))

def build_docx(jid: str, html: str, mode: str, header=None, footer=None, title="") -> tuple[bytes, str]:
    d = job_dir(jid)
    html = _rehydrate(jid, html)
    meta = load_meta(d)
    if header is not None:
        old = meta.get("header") or {}
        meta["header"] = ({"text": header, **{k: old[k] for k in ("color", "fill") if old.get(k)}}
                          if header.strip() else None)
    if footer is not None:
        old = meta.get("footer") or {}
        meta["footer"] = ({"text": footer, **{k: old[k] for k in ("color", "fill") if old.get(k)}}
                          if footer.strip() else None)
    pdf_path = d / "input.pdf"
    if mode == "fidelidad" and not pdf_path.exists():
        return html_to_docx(html, meta, title), "editable"
    pdf_bytes = pdf_path.read_bytes() if pdf_path.exists() else b""
    if mode == "fidelidad":
        try:
            data, _ = convert_layout(pdf_bytes)
            return data, "fidelidad"
        except Exception:
            data = html_to_docx(html, meta, title)
            return data, "editable"
    return html_to_docx(html, meta, title), "editable"

# --- rutas -----------------------------------------------------------------
@app.get("/api/health")
def health():
    tess = shutil.which("tesseract")
    return {"ok": True, "ocr": bool(tess), "external_service": False}

_analyze_lock = threading.Semaphore(1)  # un análisis pesado cada vez (evita agotar memoria)
_tasks: set = set()

import base64 as _b64
import hashlib as _hashlib

_IMG_DATA_RE = re.compile(r'src="data:image/(\w+);base64,([^"]+)"')
_IMG_URL_RE = re.compile(r'src="(/api/jobs/[^/]+/assets/[^"]+)"')


def _extract_assets(jid: str, html: str) -> str:
    """Sustituye las imágenes base64 por URLs a archivos del servidor.

    Así la respuesta del análisis pesa muy poco aunque el PDF tenga muchas
    fotos, y el editor carga cada imagen bajo demanda.
    """
    d = DATA / jid
    assets = d / "assets"
    seen: dict[str, str] = {}

    def repl(m):
        ext, b64 = m.group(1).lower(), m.group(2)
        key = _hashlib.sha1(b64[:2000].encode() + str(len(b64)).encode()).hexdigest()[:16]
        if key not in seen:
            fname = f"img_{key}.{'jpg' if ext in ('jpeg', 'jpg') else ext}"
            assets.mkdir(parents=True, exist_ok=True)
            try:
                (assets / fname).write_bytes(_b64.b64decode(b64))
                seen[key] = fname
            except Exception:
                return m.group(0)
        return f'src="/api/jobs/{jid}/assets/{seen[key]}"'

    return _IMG_DATA_RE.sub(repl, html)


def _rehydrate(jid: str, html: str) -> str:
    """Inverso: vuelve a incrustar las imágenes (al exportar a DOCX/Markdown)."""
    if not html or "/assets/" not in html:
        return html
    d = DATA / jid / "assets"

    def repl(m):
        fname = Path(m.group(1)).name
        f = d / fname
        if not f.exists():
            return m.group(0)
        mime = "image/png" if fname.lower().endswith(".png") else "image/jpeg"
        return f'src="data:{mime};base64,' + _b64.b64encode(f.read_bytes()).decode() + '"'

    return _IMG_URL_RE.sub(repl, html)


def _write_status(d: Path, status: str, **kw):
    data = {"status": status}
    data.update(kw)
    (d / "status.json").write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")


def _process_job(jid: str, filename: str):
    """Análisis en segundo plano: la subida ya se ha aceptado, esto puede tardar."""
    d = DATA / jid
    try:
        with _analyze_lock:
            _write_status(d, "processing", step="Analizando texto, tablas e imágenes…")
            src = next(d.glob("input.*"))
            raw = src.read_bytes()
            ext = src.suffix.lower().lstrip(".")
            if ext == "pdf":
                result = analyze(raw, filename)
            elif ext in ("md", "markdown"):
                result = import_markdown(raw, filename)
            elif ext == "docx":
                result = import_docx(raw, filename)
            elif ext == "odt":
                result = import_odt(raw, filename)
            else:
                raise ValueError(f"Formato no admitido: .{ext}")
        result["html"] = _extract_assets(jid, result["html"])
        (d / "ir.html").write_text(result["html"], encoding="utf-8")
        meta = dict(result["meta"])
        meta["job_id"] = jid
        (d / "meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=1), encoding="utf-8")
        (d / "result.json").write_text(json.dumps({
            "job_id": jid, "filename": filename, "html": result["html"],
            "meta": meta, "warnings": result["warnings"], "stats": result["stats"],
        }, ensure_ascii=False), encoding="utf-8")
        _write_status(d, "done")
    except ValueError as e:
        _write_status(d, "error", detail=str(e))
    except Exception as e:  # noqa: BLE001
        traceback.print_exc()
        _write_status(d, "error", detail=f"No se pudo analizar el PDF: {e}")


ALLOWED_EXT = {"pdf", "docx", "odt", "md", "markdown"}


@app.post("/api/analyze")
async def analyze_pdf(file: UploadFile = File(...)):
    ext = Path(file.filename or "").suffix.lower().lstrip(".")
    if ext not in ALLOWED_EXT:
        raise HTTPException(400, "Formato no admitido. Sube un PDF, DOCX, ODT o Markdown (.md).")
    raw = await file.read()
    if not raw:
        raise HTTPException(400, "El archivo está vacío.")
    if len(raw) > 80 * 1024 * 1024:
        raise HTTPException(413, "Archivo demasiado grande (máx. 80 MB). Divídelo en partes e inténtalo de nuevo.")
    if ext == "pdf" and raw[:5] != b"%PDF-":
        raise HTTPException(400, "El archivo no parece un PDF válido (cabecera no reconocida).")
    jid = uuid.uuid4().hex[:12]
    d = DATA / jid
    d.mkdir(parents=True, exist_ok=True)
    (d / f"input.{ext}").write_bytes(raw)
    _write_status(d, "processing", step="Archivo recibido. Analizando…", filename=file.filename)
    task = asyncio.get_event_loop().create_task(asyncio.to_thread(_process_job, jid, file.filename))
    _tasks.add(task)
    task.add_done_callback(_tasks.discard)
    return {"job_id": jid, "filename": file.filename, "status": "processing", "size": len(raw)}


@app.get("/api/jobs/{jid}/status")
def job_status(jid: str):
    d = job_dir(jid)
    st_file = d / "status.json"
    if not st_file.exists():
        raise HTTPException(404, "Estado no disponible.")
    st = json.loads(st_file.read_text(encoding="utf-8"))
    if st.get("status") == "done" and (d / "result.json").exists():
        result = json.loads((d / "result.json").read_text(encoding="utf-8"))
        result["status"] = "done"
        return result
    return st

@app.get("/api/jobs/{jid}/pdf")
def get_pdf(jid: str):
    d = job_dir(jid)
    if not (d / "input.pdf").exists():
        raise HTTPException(400, "Este documento no es un PDF (es la fuente importada).")
    fname = "documento.pdf"
    if (d / "meta.json").exists():
        fname = load_meta(d).get("filename", fname)
    return FileResponse(d / "input.pdf", media_type="application/pdf",
                        filename=fname,
                        content_disposition_type="inline")


@app.get("/api/jobs/{jid}/page/{pno}")
def get_page_image(jid: str, pno: int, dpi: int = 130):
    """Página del PDF original renderizada como imagen (visor integrado).

    Así la vista del original no depende del visor PDF del navegador,
    que en muchos entornos bloquea los PDF incrustados.
    """
    d = job_dir(jid)
    if not (d / "input.pdf").exists():
        raise HTTPException(400, "Este documento no es un PDF.")
    import pymupdf as fitz

    doc = fitz.open(d / "input.pdf")
    try:
        if pno < 1 or pno > len(doc):
            raise HTTPException(404, "Página fuera de rango.")
        dpi = max(60, min(220, dpi))
        pix = doc[pno - 1].get_pixmap(dpi=dpi)
        data = pix.tobytes("png")
    finally:
        doc.close()
    return StreamingResponse(io.BytesIO(data), media_type="image/png",
                             headers={"Cache-Control": "private, max-age=3600"})


@app.get("/api/jobs/{jid}/original")
def get_original(jid: str):
    d = job_dir(jid)
    src = next(d.glob("input.*"), None)
    if src is None:
        raise HTTPException(404, "Archivo original no encontrado.")
    fname = src.name
    if (d / "meta.json").exists():
        fname = load_meta(d).get("filename", fname)
    return FileResponse(src, filename=fname)

@app.post("/api/jobs/{jid}/docx")
def make_docx(jid: str, req: DocxRequest):
    d = job_dir(jid)
    meta = load_meta(d)
    html = req.html or (d / "ir.html").read_text(encoding="utf-8")
    data, used = build_docx(jid, html, req.mode, req.header, req.footer, req.title)
    out = d / "output.docx"
    out.write_bytes(data)
    name = safe_name(meta.get("filename", "documento")) + ".docx"
    return {"ok": True, "mode_used": used, "download": f"/api/jobs/{jid}/download/docx", "filename": name, "size": len(data)}

@app.get("/api/jobs/{jid}/download/docx")
def download_docx(jid: str):
    d = job_dir(jid)
    meta = load_meta(d)
    if not (d / "output.docx").exists():
        raise HTTPException(404, "Primero genera el DOCX.")
    return FileResponse(d / "output.docx",
                        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        filename=safe_name(meta.get("filename", "documento")) + ".docx")

@app.get("/api/jobs/{jid}/assets/{name}")
def get_asset(jid: str, name: str):
    d = job_dir(jid)
    f = d / "assets" / Path(name).name
    if not f.exists():
        raise HTTPException(404, "Imagen no encontrada.")
    mime = "image/png" if f.suffix.lower() == ".png" else "image/jpeg"
    return FileResponse(f, media_type=mime)


@app.post("/api/jobs/{jid}/markdown")
def make_markdown(jid: str, req: MdRequest):
    d = job_dir(jid)
    meta = load_meta(d)
    html = _rehydrate(jid, req.html or (d / "ir.html").read_text(encoding="utf-8"))
    md, images = html_to_markdown(html, meta, req.title)
    name = safe_name(meta.get("filename", "documento")) + ".md"
    (d / "output.md").write_text(md, encoding="utf-8")
    zip_name = None
    if images:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr(name, md)
            for k, v in images.items():
                z.writestr(k, v)
        (d / "output_md.zip").write_bytes(buf.getvalue())
        zip_name = f"/api/jobs/{jid}/download/md.zip"
    return {"ok": True, "markdown": md, "images": list(images.keys()), "download": f"/api/jobs/{jid}/download/md", "zip": zip_name}

@app.get("/api/jobs/{jid}/download/md")
def download_md(jid: str):
    d = job_dir(jid)
    meta = load_meta(d)
    if not (d / "output.md").exists():
        raise HTTPException(404, "Primero genera el Markdown.")
    return FileResponse(d / "output.md", media_type="text/markdown; charset=utf-8",
                        filename=safe_name(meta.get("filename", "documento")) + ".md")

@app.get("/api/jobs/{jid}/download/md.zip")
def download_mdzip(jid: str):
    d = job_dir(jid)
    meta = load_meta(d)
    if not (d / "output_md.zip").exists():
        raise HTTPException(404, "No hay ZIP de Markdown.")
    return FileResponse(d / "output_md.zip", media_type="application/zip",
                        filename=safe_name(meta.get("filename", "documento")) + "_markdown.zip")

@app.post("/api/zip")
async def zip_all(payload: dict = Body(...)):
    """payload: {items:[{job_id, html, mode}]} -> ZIP con todos los DOCX"""
    items = payload.get("items") or []
    if not items:
        raise HTTPException(400, "No hay documentos.")
    buf = io.BytesIO()
    used_names: dict[str, int] = {}
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for it in items:
            jid = it.get("job_id")
            try:
                d = job_dir(jid)
                meta = load_meta(d)
                html = it.get("html") or (d / "ir.html").read_text(encoding="utf-8")
                data, _ = build_docx(jid, html, it.get("mode", "editable"), it.get("header"), it.get("footer"), it.get("title", ""))
                base = safe_name(meta.get("filename", "documento")) + ".docx"
                used_names[base] = used_names.get(base, 0) + 1
                name = base if used_names[base] == 1 else f"{base[:-5]}_{used_names[base]}.docx"
                z.writestr(name, data)
            except HTTPException:
                continue
    return _save_package(buf.getvalue(), "documentos_docx.zip", "application/zip")

def _merge_items(payload: dict) -> list[dict]:
    """Construye la lista ordenada de documentos para refundir."""
    items = []
    for it in payload.get("items") or []:
        jid = it.get("job_id")
        if not jid:
            continue
        try:
            d = job_dir(jid)
        except HTTPException:
            continue
        meta = load_meta(d) if (d / "meta.json").exists() else {}
        stored = (d / "ir.html").read_text(encoding="utf-8") if (d / "ir.html").exists() else ""
        html = _rehydrate(jid, it.get("html") or stored)
        src = next(d.glob("input.*"), None)
        items.append({
            "job_id": jid,
            "title": it.get("title") or Path(meta.get("filename", "documento")).stem,
            "html": html,
            "meta": meta,
            "kind": "pdf" if (src and src.suffix.lower() == ".pdf") else (src.suffix.lower().lstrip(".") if src else ""),
            "src": src.read_bytes() if (src and src.suffix.lower() == ".pdf") else None,
            "outline": meta.get("outline") or [],
        })
    return items


PACKAGES = DATA_ROOT / "packages"
PACKAGES.mkdir(parents=True, exist_ok=True)
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def _save_package(data: bytes, filename: str, mime: str) -> dict:
    """Guarda un archivo generado y devuelve su URL de descarga permanente."""
    token = uuid.uuid4().hex[:16]
    (PACKAGES / f"{token}.bin").write_bytes(data)
    (PACKAGES / f"{token}.json").write_text(json.dumps(
        {"filename": filename, "mime": mime, "size": len(data)}, ensure_ascii=False),
        encoding="utf-8")
    return {"download": f"/api/files/{token}", "filename": filename, "size": len(data)}


@app.get("/api/files/{token}")
def get_package(token: str):
    meta_file = PACKAGES / f"{token}.json"
    bin_file = PACKAGES / f"{token}.bin"
    if not meta_file.exists() or not bin_file.exists():
        raise HTTPException(404, "Archivo no encontrado (puede haber caducado).")
    info = json.loads(meta_file.read_text(encoding="utf-8"))
    return FileResponse(bin_file, media_type=info["mime"], filename=info["filename"])


@app.post("/api/merge/docx")
async def merge_docx_ep(payload: dict = Body(...)):
    items = _merge_items(payload)
    if not items:
        raise HTTPException(400, "No hay documentos para refundir.")
    title = payload.get("title") or "Documentos refundidos"
    data = merged_docx(items, title)
    return _save_package(data, "documento_refundido.docx", DOCX_MIME)


@app.post("/api/merge/pdf")
async def merge_pdf_ep(payload: dict = Body(...)):
    items = _merge_items(payload)
    if not items:
        raise HTTPException(400, "No hay documentos para refundir.")
    data, notes = merged_pdf(items)
    out = _save_package(data, "documento_refundido.pdf", "application/pdf")
    out["notes"] = notes
    return out


@app.get("/api/download/app")
def download_app():
    """Descarga la aplicación web completa (código fuente listo para ejecutar)."""
    root = BASE
    buf = io.BytesIO()
    skip_dirs = {"__pycache__", "jobs", "packages", "descargas", "ejemplos", "entregas", "_msi_stage"}
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for p in root.rglob("*"):
            if any(part in skip_dirs for part in p.parts):
                continue
            if p.is_file() and p.suffix.lower() not in (".deb", ".msi"):
                z.write(p, f"pdfwordpro/{p.relative_to(root)}")
    buf.seek(0)
    return StreamingResponse(buf, media_type="application/zip",
                             headers={"Content-Disposition": "attachment; filename=pdfwordpro_app.zip"})


@app.get("/api/download/{kind}")
def download_pkg(kind: str):
    files = {"deb": "pdfwordpro.deb", "msi": "pdfwordpro.msi", "wix": "instalador_windows.zip"}
    if kind not in files:
        raise HTTPException(404, "Paquete no disponible.")
    f = BASE / "descargas" / files[kind]
    if not f.exists():
        raise HTTPException(404, "Este paquete aún no se ha generado en el servidor.")
    return FileResponse(f, filename=files[kind])


@app.get("/")
def index():
    p = STATIC / "index.html"
    if p.exists():
        return HTMLResponse(p.read_text(encoding="utf-8"))
    return HTMLResponse("<h1>PDF → Word PRO</h1><p>Interfaz no encontrada.</p>")

if STATIC.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC)), name="static")
