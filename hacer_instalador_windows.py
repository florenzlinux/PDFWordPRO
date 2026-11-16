#!/usr/bin/env python3
"""Monta descargas/instalador_windows.zip (app + proyecto WiX para compilar el .msi)."""
import shutil, zipfile
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "descargas" / "instalador_windows.zip"
SKIP = {"__pycache__", "jobs", "packages", "descargas", "ejemplos", "entregas", "_msi_stage"}
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for p in ROOT.rglob("*"):
        if any(part in SKIP for part in p.parts) or not p.is_file():
            continue
        z.write(p, f"app/{p.relative_to(ROOT)}")
    win = ROOT / "packaging" / "windows"
    for name in ("build_msi.bat", "pdfwordpro.wxs", "LEEME_INSTALADOR.txt"):
        z.write(win / name, name)
print("Generado:", OUT)
