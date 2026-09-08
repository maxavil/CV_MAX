"""
Genera el examen en PDF con el formato institucional de Qualitas,
listo para imprimir (tamano Carta).

Uso:
    python3 generar_pdf.py            # genera EXAMEN_Qualitas.pdf
    python3 generar_pdf.py otro.md    # genera otro.pdf

Requiere:  pip install markdown playwright
"""

import base64
import pathlib
import re
import sys

import markdown

BASE = pathlib.Path(__file__).parent.resolve()
LOGO = BASE / "assets" / "logo_qualitas.png"

MORADO = "#8E1C7A"     # morado institucional (tomado del logo)
TEAL = "#1192A5"       # turquesa institucional
GRIS = "#4A4A4A"
GRIS_CLARO = "#E7E7E7"


def logo_data_uri():
    return "data:image/png;base64," + base64.b64encode(LOGO.read_bytes()).decode()


CAMPOS_CANDIDATO = """
<div class="datos">
  <div class="campo ancho"><span class="etiqueta">Nombre del candidato</span><span class="linea"></span></div>
  <div class="campo"><span class="etiqueta">Fecha</span><span class="linea"></span></div>
  <div class="campo"><span class="etiqueta">Hora de inicio</span><span class="linea"></span></div>
  <div class="campo"><span class="etiqueta">Hora de t&eacute;rmino</span><span class="linea"></span></div>
</div>
"""


def md_a_html(ruta_md):
    texto = pathlib.Path(ruta_md).read_text(encoding="utf-8")
    html = markdown.markdown(
        texto,
        extensions=["tables", "fenced_code", "sane_lists", "attr_list"],
    )

    # Casilla para marcar en las opciones de la Parte A (solo presentacion)
    html = re.sub(r"<li>(<strong>[A-D]\)</strong>)", r'<li class="opcion"><span class="casilla"></span>\1', html)
    html = re.sub(r"<p>(<strong>[A-D]\)</strong>)</p>", r'<p class="opcion-sql"><span class="casilla"></span>\1</p>', html)

    # Las listas que contienen opciones se marcan para no partirse entre paginas
    html = html.replace("<ul>\n<li class=\"opcion\"", '<ul class="opciones">\n<li class="opcion"')

    # Portada: primer h1 + subtitulo
    html = html.replace("<h1>", '<h1 class="titulo">', 1)

    # Campos para llenar a mano (nombre, fecha, horas)
    html = re.sub(
        r"<p><strong>Nombre del candidato:</strong>.*?</p>",
        CAMPOS_CANDIDATO,
        html,
        flags=re.S,
    )
    return html


def construir_html(cuerpo):
    return f"""<meta charset="utf-8">
<style>
  @page {{ size: Letter; }}

  html {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}

  body {{
    font-family: "Liberation Sans", Arial, Helvetica, sans-serif;
    font-size: 10.2pt;
    line-height: 1.42;
    color: #1A1A1A;
    margin: 0;
  }}

  /* ---------- Portada ---------- */
  h1.titulo {{
    font-size: 21pt;
    color: {MORADO};
    margin: 0 0 2mm 0;
    letter-spacing: -0.2px;
    border: 0;
    padding: 0;
  }}
  h1.titulo + h3 {{
    font-size: 10.5pt;
    font-weight: 600;
    color: {TEAL};
    text-transform: uppercase;
    letter-spacing: 1.1px;
    margin: 0 0 4mm 0;
    padding: 0 0 3mm 0;
    border-bottom: 2.5px solid {MORADO};
    border-left: 0;
  }}
  .datos {{
    display: flex;
    flex-wrap: wrap;
    gap: 3mm 6mm;
    margin: 0 0 5mm 0;
  }}
  .datos .campo {{ flex: 1 1 26%; min-width: 34mm; }}
  .datos .campo.ancho {{ flex: 1 1 100%; }}
  .datos .etiqueta {{
    display: block;
    font-size: 7.4pt;
    font-weight: bold;
    color: {TEAL};
    text-transform: uppercase;
    letter-spacing: 0.6px;
    margin-bottom: 4.5mm;
  }}
  .datos .linea {{ display: block; border-bottom: 0.9px solid #9A9A9A; }}

  /* ---------- Encabezados de parte ---------- */
  h1:not(.titulo) {{
    font-size: 13pt;
    color: #FFFFFF;
    background: {MORADO};
    padding: 2.6mm 4mm;
    margin: 6mm 0 3.5mm 0;
    border-radius: 2px;
    letter-spacing: 0.3px;
    break-after: avoid;
  }}

  h2 {{
    font-size: 12pt;
    color: {MORADO};
    margin: 6mm 0 2.5mm 0;
    padding-bottom: 1.2mm;
    border-bottom: 1.5px solid {TEAL};
    break-after: avoid;
  }}

  h3 {{
    font-size: 11pt;
    color: {MORADO};
    margin: 5.5mm 0 2.5mm 0;
    padding-left: 3mm;
    border-left: 3.5px solid {TEAL};
    break-after: avoid;
  }}

  p {{ margin: 0 0 2.6mm 0; text-align: justify; }}

  strong {{ color: #000000; }}
  em {{ color: {GRIS}; }}

  /* ---------- Opciones de respuesta ---------- */
  ul, ol {{ margin: 0 0 3mm 0; padding-left: 6mm; }}
  li {{ margin-bottom: 1.6mm; break-inside: avoid; }}

  ul.opciones {{ list-style: none; padding-left: 1mm; }}
  ul.opciones li.opcion {{ margin-bottom: 2.2mm; text-align: justify; }}

  .casilla {{
    display: inline-block;
    width: 3.4mm;
    height: 3.4mm;
    border: 1.1px solid {MORADO};
    border-radius: 1px;
    margin-right: 2.2mm;
    vertical-align: -0.4mm;
  }}
  p.opcion-sql {{ margin: 2.4mm 0 1.2mm 0; break-after: avoid; }}

  /* ---------- Tablas ---------- */
  table {{
    border-collapse: collapse;
    width: 100%;
    margin: 3mm 0 4mm 0;
    font-size: 9.4pt;
    break-inside: avoid;
  }}
  th {{
    background: {MORADO};
    color: #FFFFFF;
    text-align: left;
    padding: 2mm 2.6mm;
    font-weight: 600;
  }}
  td {{ padding: 1.8mm 2.6mm; border-bottom: 0.8px solid {GRIS_CLARO}; }}
  tbody tr:nth-child(even) td {{ background: #F7F3F7; }}

  /* ---------- Codigo ---------- */
  pre {{
    background: #F5F7F8;
    border-left: 3px solid {TEAL};
    border-radius: 2px;
    padding: 2.2mm 3mm;
    margin: 1.4mm 0 2.6mm 0;
    font-size: 8.4pt;
    line-height: 1.3;
    white-space: pre-wrap;
    break-inside: avoid;
  }}
  pre, code {{ font-family: "Liberation Mono", "Courier New", monospace; }}
  p code, li code, td code {{
    background: #F1EDF1;
    color: {MORADO};
    padding: 0.3mm 1.1mm;
    border-radius: 2px;
    font-size: 9.1pt;
  }}

  hr {{
    border: 0;
    border-top: 0.8px solid #EFEFEF;
    margin: 3.4mm 0;
  }}

  blockquote {{
    margin: 0 0 3mm 0;
    padding-left: 3mm;
    border-left: 2px solid {GRIS_CLARO};
    color: {GRIS};
  }}
</style>
{cuerpo}
"""


def encabezado_html(logo_uri):
    return f"""
<div style="width:100%; font-family:'Liberation Sans',Arial,sans-serif; font-size:7pt;
            color:{GRIS}; padding:0 16mm; box-sizing:border-box; -webkit-print-color-adjust:exact;">
  <div style="display:flex; align-items:flex-end; justify-content:space-between;
              border-bottom:1.5px solid {MORADO}; padding-bottom:1.5mm;">
    <div style="line-height:1.3;">
      <div style="font-size:7.6pt; font-weight:bold; color:{MORADO}; letter-spacing:0.4px;">
        PRUEBA T&Eacute;CNICA &mdash; CIENCIA DE DATOS
      </div>
      <div style="font-size:6.6pt; color:{TEAL}; letter-spacing:0.6px;">
        PROCESO DE SELECCI&Oacute;N &middot; 30 MINUTOS &middot; 100 PUNTOS
      </div>
    </div>
    <img src="{logo_uri}" style="height:9.2mm;">
  </div>
</div>
"""


def pie_html():
    return f"""
<div style="width:100%; font-family:'Liberation Sans',Arial,sans-serif; font-size:6.8pt;
            color:{GRIS}; padding:0 16mm; box-sizing:border-box; -webkit-print-color-adjust:exact;">
  <div style="display:flex; justify-content:space-between; align-items:center;
              border-top:0.8px solid {GRIS_CLARO}; padding-top:1.5mm;">
    <span>Qu&aacute;litas Compa&ntilde;&iacute;a de Seguros &middot; Documento de uso interno</span>
    <span>P&aacute;gina <span class="pageNumber"></span> de <span class="totalPages"></span></span>
  </div>
</div>
"""


def buscar_chromium():
    """Chromium a usar. Prioriza CHROMIUM_PATH; si no, deja que Playwright decida."""
    import os
    import glob

    ruta = os.environ.get("CHROMIUM_PATH")
    if ruta and pathlib.Path(ruta).exists():
        return ruta
    patrones = [
        "/opt/pw-browsers/chromium-*/chrome-linux/chrome",
        "/usr/bin/chromium",
        "/usr/bin/chromium-browser",
        "/usr/bin/google-chrome",
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    ]
    for patron in patrones:
        encontrados = sorted(glob.glob(patron))
        if encontrados:
            return encontrados[-1]
    return None


def generar(ruta_md, ruta_pdf):
    from playwright.sync_api import sync_playwright

    html = construir_html(md_a_html(ruta_md))
    tmp = BASE / "_examen_tmp.html"
    tmp.write_text(html, encoding="utf-8")
    logo_uri = logo_data_uri()

    with sync_playwright() as p:
        ejecutable = buscar_chromium()
        navegador = p.chromium.launch(executable_path=ejecutable) if ejecutable \
            else p.chromium.launch()
        pagina = navegador.new_page()
        pagina.goto(tmp.as_uri())
        pagina.emulate_media(media="print")
        pagina.pdf(
            path=str(ruta_pdf),
            format="Letter",
            print_background=True,
            display_header_footer=True,
            header_template=encabezado_html(logo_uri),
            footer_template=pie_html(),
            margin={"top": "23mm", "bottom": "15mm", "left": "16mm", "right": "16mm"},
        )
        navegador.close()

    tmp.unlink(missing_ok=True)
    print(f"PDF generado: {ruta_pdf}")


if __name__ == "__main__":
    origen = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else BASE / "EXAMEN.md"
    destino = (BASE / "EXAMEN_Qualitas.pdf") if len(sys.argv) < 2 else origen.with_suffix(".pdf")
    generar(origen, destino)
