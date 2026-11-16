"""Modo máxima fidelidad visual: pdf2docx (conserva geometría). Fallback al modo editable."""
from __future__ import annotations
import logging, threading, tempfile, os
logger=logging.getLogger("layout")

def convert_layout(pdf_bytes: bytes) -> tuple[bytes, list[str]]:
    warnings=[]
    try:
        from pdf2docx import Converter
    except Exception as e:
        raise RuntimeError(f"pdf2docx no disponible: {e}")
    with tempfile.TemporaryDirectory() as td:
        src=os.path.join(td,"in.pdf"); dst=os.path.join(td,"out.docx")
        open(src,"wb").write(pdf_bytes)
        cv=Converter(src)
        try:
            cv.convert(dst, start=0, end=None)
        finally:
            cv.close()
        if not os.path.exists(dst) or os.path.getsize(dst)==0:
            raise RuntimeError("pdf2docx no generó salida")
        data=open(dst,"rb").read()
    warnings.append("Modo fidelidad visual: la geometría se conserva mejor, pero algunos párrafos pueden estar dentro de tablas o cuadros de texto de Word.")
    return data, warnings
