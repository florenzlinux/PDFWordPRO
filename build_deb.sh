#!/usr/bin/env bash
# Construye pdfwordpro_<version>_all.deb con la aplicación en /opt/pdfwordpro
set -euo pipefail
VERSION="${1:-1.3.0}"
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
BUILD="$(mktemp -d)"
PKG="$BUILD/pdfwordpro"
trap 'rm -rf "$BUILD"' EXIT

mkdir -p "$PKG/opt/pdfwordpro" "$PKG/usr/bin" "$PKG/lib/systemd/system" "$PKG/DEBIAN"

# código de la aplicación (sin datos de sesiones ni paquetes generados)
cp -a "$ROOT/." "$PKG/opt/pdfwordpro/"
find "$PKG/opt/pdfwordpro" -name '__pycache__' -type d -prune -exec rm -rf {} + 2>/dev/null || true
rm -rf "$PKG/opt/pdfwordpro/descargas" "$PKG/opt/pdfwordpro/data/jobs" "$PKG/opt/pdfwordpro/data/packages" "$PKG/opt/pdfwordpro/ejemplos" "$PKG/opt/pdfwordpro/entregas"
rm -rf "$PKG/opt/pdfwordpro/packaging/_msi_stage"
rm -f "$PKG/opt/pdfwordpro"/*.deb "$PKG/opt/pdfwordpro"/*.msi
mkdir -p "$PKG/opt/pdfwordpro/data/jobs"

cat > "$PKG/usr/bin/pdfwordpro" <<'LAUNCH'
#!/bin/sh
# PDF -> Word PRO (servidor). Valida que el entorno Python este completo;
# si no lo esta, lo crea/instala en el directorio del usuario (sin root).
cd /opt/pdfwordpro || exit 1
echo "[pdfwordpro] arranque $(date '+%F %T')"

py_ok() { [ -x "$1" ] && "$1" -c "import uvicorn, fastapi, pymupdf, docx, pdf2docx, markdown_it" 2>/dev/null; }

PY=""
if py_ok /opt/pdfwordpro/.venv/bin/python; then
  PY=/opt/pdfwordpro/.venv/bin/python
else
  USERBASE="${HOME:-/tmp}/.local/share/pdfwordpro"
  mkdir -p "$USERBASE"
  if ! py_ok "$USERBASE/venv/bin/python"; then
    echo "[pdfwordpro] preparando el entorno en $USERBASE/venv (solo la primera vez, necesita internet)..."
    if [ ! -x "$USERBASE/venv/bin/python" ]; then
      python3 -m venv "$USERBASE/venv" || {
        echo "[pdfwordpro] ERROR: no se pudo crear el entorno virtual."
        echo "Comprueba que estan instalados los paquetes: python3-venv y python3-pip"
        echo "  sudo apt install python3-venv python3-pip"
        exit 1
      }
    fi
    "$USERBASE/venv/bin/pip" install -r /opt/pdfwordpro/requirements.txt || {
      echo "[pdfwordpro] ERROR: no se pudieron instalar las dependencias (revisa tu conexion)."
      exit 1
    }
  fi
  PY="$USERBASE/venv/bin/python"
fi

: "${PDFWORDPRO_DATA:=${HOME:-/tmp}/.local/share/pdfwordpro/data}"
export PDFWORDPRO_DATA
mkdir -p "$PDFWORDPRO_DATA"
echo "[pdfwordpro] servidor en http://localhost:${PDFWORDPRO_PORT:-8000} (datos en $PDFWORDPRO_DATA)"
exec "$PY" -m uvicorn app.main:app --host 0.0.0.0 --port "${PDFWORDPRO_PORT:-8000}"
LAUNCH
chmod 755 "$PKG/usr/bin/pdfwordpro"

# Lanzador de escritorio: arranca el servidor si hace falta y abre el navegador
mkdir -p "$PKG/usr/share/applications" "$PKG/usr/share/icons/hicolor/scalable/apps"
cat > "$PKG/usr/bin/pdfwordpro-app" <<'APP'
#!/bin/sh
# Abre PDF -> Word PRO: levanta el servidor en segundo plano y abre el navegador.
# Si algo falla, lo avisa y abre el registro (nada de fallos silenciosos).
URL="http://localhost:${PDFWORDPRO_PORT:-8000}"
LOGDIR="${HOME:-/tmp}/.local/share/pdfwordpro"
LOG="$LOGDIR/server.log"
mkdir -p "$LOGDIR"
sano() { python3 -c "import urllib.request; urllib.request.urlopen('$URL/api/health', timeout=2)" 2>/dev/null; }

if ! sano; then
  echo "[pdfwordpro-app] $(date '+%F %T') arrancando servidor..." >>"$LOG"
  nohup pdfwordpro >>"$LOG" 2>&1 &
  # la primera vez instala dependencias: puede tardar varios minutos
  i=0
  while [ $i -lt 300 ]; do
    sano && break
    i=$((i + 1)); sleep 1
  done
fi

if sano; then
  xdg-open "$URL" >/dev/null 2>&1 &
  exit 0
fi

# No ha arrancado: avisar con el motivo a la vista
if command -v notify-send >/dev/null 2>&1; then
  notify-send "PDF Word PRO no pudo arrancar" "Se abre el registro con el motivo. Tambien puedes ejecutar 'pdfwordpro' en una terminal." 2>/dev/null || true
fi
xdg-open "$LOG" >/dev/null 2>&1 &
exit 1
APP
chmod 755 "$PKG/usr/bin/pdfwordpro-app"

cat > "$PKG/usr/share/applications/pdfwordpro.desktop" <<'DESKTOP'
[Desktop Entry]
Type=Application
Version=1.0
Name=PDF Word PRO
GenericName=Conversor PDF a Word
Comment=Convierte PDF, DOCX, ODT y Markdown a DOCX editable y Markdown
Exec=/usr/bin/pdfwordpro-app
Icon=pdfwordpro
Terminal=false
Categories=Office;Utility;FileTools;
Keywords=pdf;docx;word;markdown;convertir;conversor;
StartupNotify=true
DESKTOP
cp "$ROOT/packaging/pdfwordpro.svg" "$PKG/usr/share/icons/hicolor/scalable/apps/pdfwordpro.svg"
mkdir -p "$PKG/usr/share/metainfo"
cat > "$PKG/usr/share/metainfo/pdfwordpro.metainfo.xml" <<'META'
<?xml version="1.0" encoding="UTF-8"?>
<component type="desktop-application">
  <id>pdfwordpro</id>
  <metadata_license>CC0-1.0</metadata_license>
  <project_license>MIT</project_license>
  <name>PDF Word PRO</name>
  <summary>Convierte PDF a DOCX editable y Markdown</summary>
  <description>
    <p>Aplicación web local para convertir PDF, DOCX, ODT y Markdown a documentos
    Word editables y Markdown, con editor previo, fusión de documentos en uno solo
    conservando marcadores, conversión por lotes y OCR opcional. Funciona sin enviar
    tus documentos a ningún servicio externo.</p>
  </description>
  <categories>
    <category>Office</category>
    <category>Utility</category>
  </categories>
  <launchable type="desktop-id">pdfwordpro.desktop</launchable>
  <provides>
    <binary>pdfwordpro</binary>
  </provides>
</component>
META

cat > "$PKG/lib/systemd/system/pdfwordpro.service" <<'UNIT'
[Unit]
Description=PDF Word PRO - conversor PDF a DOCX/Markdown
After=network.target

[Service]
Type=simple
ExecStart=/usr/bin/pdfwordpro
Restart=on-failure
Environment=PDFWORDPRO_PORT=8000

[Install]
WantedBy=multi-user.target
UNIT

cat > "$PKG/DEBIAN/control" <<CONTROL
Package: pdfwordpro
Version: $VERSION
Section: web
Priority: optional
Architecture: all
Depends: python3, python3-venv, python3-pip, xdg-utils
Recommends: tesseract-ocr, tesseract-ocr-spa
Maintainer: PDF Word PRO
Description: Conversor web PDF a DOCX editable y Markdown
 Aplicación web local para convertir PDF, DOCX, ODT y Markdown a DOCX
 editable y Markdown, con editor previo, fusión de documentos con
 marcadores y OCR opcional. Se arranca con: pdfwordpro
 (o el servicio systemd pdfwordpro). Interfaz en http://localhost:8000
CONTROL

cat > "$PKG/DEBIAN/postinst" <<'POST'
#!/bin/sh
# Instalación ligera y segura: nada de red ni entornos en este paso.
# (El entorno Python se crea en la carpeta del usuario al abrir la app.)
if [ "$1" = "configure" ]; then
  if command -v update-desktop-database >/dev/null 2>&1; then
    update-desktop-database -q /usr/share/applications 2>/dev/null || true
  fi
  if command -v gtk-update-icon-cache >/dev/null 2>&1; then
    gtk-update-icon-cache -q -t -f /usr/share/icons/hicolor 2>/dev/null || true
  fi
  echo "PDF Word PRO instalado: buscalo en el menu de aplicaciones."
  echo "Si no aparece al instante, cierra sesion y vuelve a entrar."
fi
exit 0
POST
chmod 755 "$PKG/DEBIAN/postinst"

cat > "$PKG/DEBIAN/postrm" <<'RM'
#!/bin/sh
case "$1" in
  remove)
    # el entorno generado ya no sirve sin el paquete
    rm -rf /opt/pdfwordpro/.venv 2>/dev/null || true
    ;;
  purge)
    rm -rf /opt/pdfwordpro 2>/dev/null || true
    ;;
esac
exit 0
RM
chmod 755 "$PKG/DEBIAN/postrm"

mkdir -p "$ROOT/descargas"
dpkg-deb --root-owner-group --build "$PKG" "$ROOT/descargas/pdfwordpro.deb"
echo "Generado: $ROOT/descargas/pdfwordpro.deb"
