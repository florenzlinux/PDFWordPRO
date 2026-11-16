#!/usr/bin/env python3
"""
Genera pdfwordpro.msi en Linux con wixl (WiX v3), a partir del código de la app.

Uso:    python3 packaging/hacer_msi_linux.py
Salida: descargas/pdfwordpro.msi   (y el .wxs generado en packaging/windows/)

El MSI instala en Archivos de programa\\PDFWordPRO y crea accesos directos
en el menú Inicio y en el escritorio hacia iniciar_windows.bat, que crea su
entorno y sus datos en %LOCALAPPDATA%\\PDFWordPRO (sin escribir en la
carpeta de instalación, así funciona con usuarios sin permisos de admin).
"""
import re
import shutil
import subprocess
import sys
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "descargas" / "pdfwordpro.msi"
VERSION = "1.2.0"
SKIP_DIRS = {"__pycache__", "jobs", "packages", "descargas", "ejemplos", "entregas", "packaging", "data"}
NS = uuid.uuid5(uuid.NAMESPACE_DNS, "pdfwordpro")
LAUNCHER = "iniciar_windows.bat"


def gid(name: str) -> str:
    return "{" + str(uuid.uuid5(NS, name)).upper() + "}"


def fid(name: str) -> str:
    return ("f_" + re.sub(r"[^A-Za-z0-9_]", "_", name))[:60]


def main() -> int:
    if not shutil.which("wixl"):
        print("Falta wixl:  sudo apt install wixl wixl-data", file=sys.stderr)
        return 1

    # ---- recopilar archivos de la app (sin el lanzador: va en su componente) ----
    files: list[Path] = []
    for p in sorted(ROOT.rglob("*")):
        if not p.is_file():
            continue
        rel = p.relative_to(ROOT)
        if any(part in SKIP_DIRS for part in rel.parts):
            continue
        if rel.name.endswith((".deb", ".msi")) or str(rel) == LAUNCHER:
            continue
        files.append(rel)

    stage = ROOT / "packaging" / "_msi_stage"
    if stage.exists():
        shutil.rmtree(stage)
    stage.mkdir(parents=True)
    ico_src = ROOT / "packaging" / "windows" / "pdfwordpro.ico"
    for rel in files + [Path(LAUNCHER)]:
        target = stage / rel
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(ROOT / rel, target)

    # ---- componentes por archivo, agrupados por carpeta ----
    by_dir: dict[str, list[str]] = {}
    comp_ids: list[str] = []
    for rel in files:
        rel_win = str(rel)  # wixl en Linux usa rutas con /
        cid = "c_" + fid(rel_win)
        comp_ids.append(cid)
        comp = (
            f'<Component Id="{cid}" Guid="{gid(rel_win)}">'
            f'<File Id="{fid(rel_win)}" Name="{rel.name}" Source="{rel_win}" KeyPath="yes" />'
            f"</Component>"
        )
        by_dir.setdefault(str(rel.parent), []).append(comp)

    # ---- árbol de directorios (recursivo) ----
    all_dirs = set(by_dir.keys())

    def children_of(parent: str):
        return sorted(
            d for d in all_dirs
            if d not in ("", ".") and str(Path(d).parent) == parent
        )

    def dir_block(parent: str, dir_id: str, name: str, indent: int) -> str:
        pad = " " * indent
        lines = [f'{pad}<Directory Id="{dir_id}" Name="{name}">']
        for comp in by_dir.get(parent, []):
            lines.append(pad + "  " + comp)
        for child in children_of(parent):
            lines.append(dir_block(child, "d_" + fid(child), Path(child).name, indent + 2))
        lines.append(f"{pad}</Directory>")
        return "\n".join(lines)

    install_dir = dir_block(".", "INSTALLFOLDER", "PDFWordPRO", 8)

    feature_refs = "\n".join(f'      <ComponentRef Id="{c}" />' for c in comp_ids)
    shortcuts_guid = gid("shortcuts")

    wxs = f'''<?xml version="1.0" encoding="UTF-8"?>
<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">
  <Product Id="{gid("product" + VERSION)}" Name="PDF Word PRO" Language="3082"
           Version="{VERSION}" Manufacturer="PDF Word PRO"
           UpgradeCode="{gid("upgrade")}">
    <Package InstallerVersion="200" Compressed="yes" InstallScope="perMachine"
             Description="Conversor web PDF a DOCX editable y Markdown" />
    <Media Id="1" Cabinet="product.cab" EmbedCab="yes" />
    <Icon Id="appico" SourceFile="pdfwordpro.ico" />
    <Directory Id="TARGETDIR" Name="SourceDir">
      <Directory Id="ProgramFilesFolder">
{install_dir}
      </Directory>
      <Directory Id="ProgramMenuFolder">
        <Directory Id="AppShortcutFolder" Name="PDF Word PRO" />
      </Directory>
      <Directory Id="DesktopFolder" />
    </Directory>
    <DirectoryRef Id="INSTALLFOLDER">
      <Component Id="cmp_lanzador" Guid="{shortcuts_guid}">
        <File Id="f_lanzador" Name="{LAUNCHER}" Source="{LAUNCHER}" KeyPath="yes">
          <Shortcut Id="sc_menu" Directory="AppShortcutFolder" Name="PDF Word PRO"
                    Description="Arranca PDF Word PRO en http://localhost:8000"
                    WorkingDirectory="INSTALLFOLDER" Icon="appico" />
          <Shortcut Id="sc_desktop" Directory="DesktopFolder" Name="PDF Word PRO"
                    Description="Arranca PDF Word PRO en http://localhost:8000"
                    WorkingDirectory="INSTALLFOLDER" Icon="appico" />
        </File>
      </Component>
    </DirectoryRef>
    <Feature Id="MainFeature" Title="PDF Word PRO" Level="1">
{feature_refs}
      <ComponentRef Id="cmp_lanzador" />
    </Feature>
  </Product>
</Wix>
'''
    if ico_src.exists():
        shutil.copy2(ico_src, stage / "pdfwordpro.ico")
    wxs_generated = ROOT / "packaging" / "windows" / "pdfwordpro_v3.wxs"
    wxs_generated.write_text(wxs, encoding="utf-8")
    (stage / "pdfwordpro.wxs").write_text(wxs, encoding="utf-8")

    r = subprocess.run(["wixl", "-o", str(OUT), "pdfwordpro.wxs"], cwd=stage,
                       capture_output=True, text=True)
    shutil.rmtree(stage, ignore_errors=True)  # no dejar la copia intermedia
    if r.returncode != 0:
        print(r.stdout[-4000:])
        print(r.stderr[-4000:], file=sys.stderr)
        return r.returncode
    print(f"Generado: {OUT} ({OUT.stat().st_size/1024:.0f} KB) con {len(files)} archivos")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
