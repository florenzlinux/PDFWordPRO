# PDF → Word PRO

Aplicación web para convertir **PDF → DOCX editable** y **PDF → Markdown**, con vista previa editable antes de exportar, conversión por lotes e informe de calidad por documento.

## Qué hace

- **Ver el original siempre**: la pestaña «Documento original» muestra el PDF página a página renderizado en el servidor como imagen (anterior/siguiente y zoom), sin depender del visor PDF del navegador; también permite abrirlo en otra pestaña o descargarlo.
- **Diseño**: interfaz clara tipo panel (fondo gris muy suave, tarjetas blancas, azul principal, tipografía del sistema con Inter si está disponible).
- **Reordenar antes de exportar**: flechas ↑ ↓ o arrastrar las tarjetas del panel lateral; el ZIP y el refundido respetan ese orden (las tarjetas van numeradas).
- **Refundir en uno solo**: un único **DOCX** (con campo de Índice y un marcador de Word por cada documento) o un único **PDF** (une los PDF originales y construye/conserva sus marcadores: entrada por cada archivo y, si el PDF traía índice interno, anidado bajo él). Las fuentes MD/DOCX/ODT entran en el PDF refundido renderizadas.
- **Importar también MD, DOCX y ODT**: se importan al mismo editor y se pueden editar, exportar y refundir como un PDF más.
- **Estilos y colores**: se conservan el color real de cada texto (p. ej. títulos de color), fondos de recuadros, cabeceras de tabla con su color, texto justificado, enlaces, y encabezado/pie con su color y banda de fondo y numeración automática «N de M».
- **Subida múltiple**: arrastra o selecciona varios PDF a la vez (hasta 80 MB por archivo).
- **Análisis en segundo plano**: la subida se acepta al instante y la interfaz muestra el progreso con contador, por lo que los documentos largos (decenas de páginas) ya no provocan cortes del tipo *Failed to fetch*. Las imágenes se sirven como archivos bajo demanda en vez de incrustadas, así la vista previa se mueve con ligereza aunque el PDF tenga una foto por página.
- **Editor previo real**: modifica títulos, párrafos, negrita/cursiva/subrayado, listas, alineación, tablas, imágenes, inserta saltos de página y deshaz/rehaz. **Lo que ves en el editor es lo que se exporta.**
- **Dos motores a Word**:
  - *Editable* (defecto): reconstruye con PyMuPDF + python-docx → títulos jerárquicos, párrafos con estilos inline, listas, tablas, imágenes, encabezado/pie detectados, tamaño de página/orientación y saltos entre páginas.
  - *Fidelidad visual*: `pdf2docx`; se parece más al original cuando la maquetación manda, aunque el documento puede ser menos cómodo de editar.
- **Markdown**: generar, editar, **copiar** con un clic, descargar `.md` o `.md` + imágenes en ZIP (`images/imagen-N.png`).
- **Lotes**: descarga individual o **ZIP con todos los DOCX**.
- **Encabezados y pies**: detecta texto repetido en márgenes y lo pasa al encabezado/pie de Word (escribe "Página" en el pie para numeración automática).
- **Informe de calidad**: avisos por tablas dudosas, columnas, páginas escaneadas, OCR aplicado, etc.
- **OCR**: Tesseract opcional para páginas sin capa texto.

## Descargar la aplicación

En la propia app (pestaña Exportación) o en `descargas/`:

- **Aplicación web (.zip)**: código completo; descomprimir, `pip install -r requirements.txt`, `./run.sh` (en Windows, `iniciar_windows.bat`).
- **`.deb` (Linux)**: `sudo dpkg -i pdfwordpro.deb` -> aparece como **«PDF Word PRO» en el menú de aplicaciones** (Oficina/Utilidades, con su icono); al abrirlo levanta el servidor y abre el navegador en http://localhost:8000. También se puede arrancar con `pdfwordpro` en terminal o como servicio: `sudo systemctl enable --now pdfwordpro`. La instalación prepara el entorno Python automáticamente (necesita red en ese momento; si no la hay, el entorno se crea en tu usuario al abrir la app). Tus documentos de trabajo quedan en `~/.local/share/pdfwordpro/`.
- **Windows (.msi)**: un .msi solo puede compilarse en Windows con WiX Toolset. `descargas/instalador_windows.zip` trae la app + `pdfwordpro.wxs` + `build_msi.bat` (doble clic y genera `pdfwordpro.msi`). Sin compilar nada, la carpeta `app\` con `iniciar_windows.bat` funciona igual.

## Arranque rápido

```bash
cd pdfwordpro
pip install -r requirements.txt
# opcional pero recomendado para OCR:
sudo apt install tesseract-ocr tesseract-ocr-spa
./run.sh
# abre http://localhost:8000
```

En Docker / Render / Railway / VPS: sirve `uvicorn app.main:app --host 0.0.0.0 --port 8000` tras instalar requirements (+ Tesseract). En España, un VPS europeo o tu Raspberry Pi con Tailscale son buenas opciones.

## Privacidad

- **Modo privado**: todo se procesa en este servidor. Ningún PDF sale fuera.
- **Modo externo (Adobe PDF Services, etc.)**: **no configurado por defecto y deshabilitado en la interfaz**. La app avisa explícitamente de ello para que nunca haya envíos silenciosos. Si lo quieres, hay que integrar la API con credenciales y una autorización por documento.

## Honestidad técnica

Un PDF fija la apariencia; Word es texto reflujo. **Ningún conversor garantiza copia 1:1 universal.** Casos límite (columnas complejas, cuadros de texto, gráficos vectoriales, formularios) pueden requerir retoques en el editor. Los PDF escaneados dependen del OCR.

## Estructura

- `app/main.py` – API FastAPI
- `app/extract.py` – PDF → representación intermedia (texto enriquecido, tablas, imágenes, columnas, encabezados)
- `app/html_docx.py` – HTML editado → DOCX
- `app/layout_docx.py` – modo fidelidad visual (pdf2docx)
- `app/html_md.py` – HTML → Markdown + extracción de imágenes
- `app/static/index.html` – interfaz

## Pruebas realizadas

Ejemplos generados en `ejemplos/`:
- `ejemplo-editable.docx` y `ejemplo.md` – desde un PDF de 2 páginas con encabezado/pie, listas y tabla.
- Conversiones comprobadas: PDF simple (títulos/estilos/tabla/imagen), PDF multipágina con listas, PDF solo-imagen (aviso), y OCR Tesseract sobre página escaneada.
