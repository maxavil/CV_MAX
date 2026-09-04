#!/usr/bin/env python3
"""Comprueba el bloque de celda_jupyter.py contra las cifras de control.

Ejecuta la lectura de los dos Excel del cierre de junio 2026 sin abrir ninguna
ventana, y contrasta contra los importes verificados uno por uno.

    python verificar.py Balanza_062026.xlsx ResultadosQES.xlsb
"""

import sys
import tempfile
from pathlib import Path

AQUI = Path(__file__).resolve().parent

# ---------------------------------------------------------------------------
# El bloque se ejecuta tal cual, salvo la última línea que abre la ventana.
# ---------------------------------------------------------------------------
FUENTE = (AQUI / "celda_jupyter.py").read_text(encoding="utf-8")
assert FUENTE.count("\nabrir_ventana()\n") == 1, "cambió el pie del bloque"
BLOQUE = {}
exec(compile(FUENTE.replace("\nabrir_ventana()\n", "\n"), "celda_jupyter.py", "exec"), BLOQUE)

Historico = BLOQUE["Historico"]
construir_html = BLOQUE["construir_html"]
escribir_vista = BLOQUE["escribir_vista"]
etiqueta_corta = BLOQUE["etiqueta_corta"]

# ---------------------------------------------------------------------------
# Cifras de control (sección 05 del encargo), en USD
# ---------------------------------------------------------------------------
ESPERADO = {
    #  periodo        rrc local      rrc cnsf       rsr (los dos)  rsnr local  rsnr cnsf
    "2025-07-31": (8_331_317.86, 11_142_522.97, 2_818_658.21, 12_701.88, 325_011.19),
    "2025-09-30": (8_723_257.33, 11_817_782.40, 3_722_728.06, 12_701.88, 423_226.93),
    "2025-11-30": (9_132_774.22, 12_231_493.45, 4_246_854.22, 12_701.88, 599_138.64),
    "2025-12-31": (9_278_886.43, 12_604_139.42, 3_890_396.27,  7_461.06, 419_827.70),
    "2026-03-31": (9_171_744.38, 12_718_231.37, 4_187_990.32,  7_461.06, 555_998.70),
    "2026-06-30": (9_084_601.70, 12_799_633.26, 2_773_742.22,  7_461.06, 809_405.38),
}

# La vista, en millones de USD: (local, cnsf, dif) por corte y el incremento
VISTA = {
    "2025-12-31": {"rrc": (9.28, 12.60, 3.33), "rsr": (3.89, 3.89, None),
                   "rsnr": (0.01, 0.42, 0.41), "total": (13.18, 16.91, 3.74)},
    "2026-03-31": {"rrc": (9.17, 12.72, 3.55), "rsr": (4.19, 4.19, None),
                   "rsnr": (0.01, 0.56, 0.55), "total": (13.37, 17.46, 4.10)},
    "2026-06-30": {"rrc": (9.08, 12.80, 3.72), "rsr": (2.77, 2.77, None),
                   "rsnr": (0.01, 0.81, 0.80), "total": (11.87, 16.38, 4.52)},
}
INCREMENTO = {"rrc": 0.17, "rsr": None, "rsnr": 0.25, "total": 0.42}

fallos: "list[str]" = []
pruebas = 0


def igual(etiqueta, obtenido, esperado, tol):
    global pruebas
    pruebas += 1
    if obtenido is None and esperado is None:
        return
    if obtenido is None or esperado is None or abs(obtenido - esperado) > tol:
        fallos.append(f"{etiqueta}: se obtuvo {obtenido!r}, se esperaba {esperado!r}")


def main(argv):
    global pruebas
    balanza = Path(argv[0]) if argv else AQUI / "Balanza_062026.xlsx"
    actuarios = Path(argv[1]) if len(argv) > 1 else AQUI / "ResultadosQES.xlsb"
    for f in (balanza, actuarios):
        if not f.exists():
            print(f"No encuentro {f}. Uso: python verificar.py <balanza.xlsx> <actuarios.xlsb>")
            return 2

    tmp = Path(tempfile.mkdtemp(prefix="reservas-qes-"))

    # ---- 1. lectura de los dos archivos ---------------------------------
    h = Historico(tmp / "historico.json")
    for aviso in h.procesar(actuarios):
        print("·", aviso)
    for aviso in h.procesar(balanza):
        print("·", aviso)
    h.guardar()
    print()

    igual("número de cortes", len(h.periodos()), 6, 0)
    for p, (rrc_l, rrc_c, rsr, rsnr_l, rsnr_c) in ESPERADO.items():
        if p not in h.datos:
            fallos.append(f"falta el corte {p}")
            continue
        e = h.datos[p]
        igual(f"{p} rrc local",  e["local"].get("rrc"),  rrc_l,  0.005)
        igual(f"{p} rrc cnsf",   e["cnsf"].get("rrc"),   rrc_c,  0.005)
        igual(f"{p} rsr local",  e["local"].get("rsr"),  rsr,    0.005)
        igual(f"{p} rsr cnsf",   e["cnsf"].get("rsr"),   rsr,    0.005)
        igual(f"{p} rsnr local", e["local"].get("rsnr"), rsnr_l, 0.005)
        igual(f"{p} rsnr cnsf",  e["cnsf"].get("rsnr"),  rsnr_c, 0.005)

    # la balanza de junio debe cuadrar al centavo con la columna local de actuarios
    ref = h.datos["2026-06-30"].get("local_actuarios", {})
    for cid in ("rrc", "rsr", "rsnr"):
        igual(f"2026-06-30 balanza vs actuarios ({cid})",
              h.datos["2026-06-30"]["local"][cid], ref.get(cid), 0.005)
    if h.datos["2026-06-30"].get("aviso"):
        fallos.append("junio 2026 quedó marcado: " + h.datos["2026-06-30"]["aviso"])

    # ---- 2. la vista en millones ----------------------------------------
    ps = h.periodos()[-3:]
    igual("cortes en la vista", len(ps), 3, 0)
    for p, filas in VISTA.items():
        for cid, (loc, cnsf, dif) in filas.items():
            if cid == "total":
                igual(f"{p} total local", h.total(p, "local") / 1e6, loc, 0.005)
                igual(f"{p} total cnsf", h.total(p, "cnsf") / 1e6, cnsf, 0.005)
                igual(f"{p} total dif", h.diferencia(p) / 1e6, dif, 0.005)
            else:
                igual(f"{p} {cid} local", h.datos[p]["local"][cid] / 1e6, loc, 0.005)
                igual(f"{p} {cid} cnsf", h.datos[p]["cnsf"][cid] / 1e6, cnsf, 0.005)
                d = h.diferencia(p, cid) / 1e6
                igual(f"{p} {cid} dif", None if abs(d) < 0.005 else d, dif, 0.005)

    for cid, inc in INCREMENTO.items():
        v = ((h.diferencia(ps[-1]) - h.diferencia(ps[-2])) / 1e6 if cid == "total"
             else h.incremento(ps[-1], ps[-2], cid) / 1e6)
        igual(f"incremento {cid}", None if abs(v) < 0.005 else v, inc, 0.005)

    # ---- 3. indicadores --------------------------------------------------
    fx = BLOQUE["FX_DEFAULT"]
    d1 = h.diferencia("2026-06-30")
    d0 = h.diferencia("2026-03-31")
    igual("KPI diferencia total (MM USD)", d1 / 1e6, 4.52, 0.005)
    igual("KPI variación del trimestre (MM USD)", (d1 - d0) / 1e6, 0.42, 0.005)
    igual("KPI variación %", (d1 - d0) / d0 * 100, 10.3, 0.05)
    igual("KPI equivalente en MXN (MM)", d1 * fx / 1e6, 79.04, 0.01)

    # ---- 4. el HTML ------------------------------------------------------
    ruta = escribir_vista(h, periodos=ps, abrir=False)
    doc = ruta.read_text(encoding="utf-8")
    igual("nombre del archivo", 1 if ruta.name == "vista_reservas_2026-06-30.html" else 0, 1, 0)
    for cadena in ("Resultados reservas técnicas QES",   # el CSS lo pone en versalitas
                   "text-transform:uppercase",
                   "Comparativo de metodologías y evolución del diferencial",
                   "17.4986 MXN / USD", "Incremento", "Total reservas",
                   "Esta reserva se calcula una vez al año",
                   "Mensajes clave", "Nota relevante", "<svg", "4.52", "10.3%", "0.42"):
        pruebas += 1
        if cadena not in doc:
            fallos.append(f"el HTML no trae «{cadena}»")
    for prohibido in ("http://", "https://", "<script"):
        pruebas += 1
        if prohibido in doc:
            fallos.append(f"el HTML no es autocontenido: contiene «{prohibido}»")
    pruebas += 1
    if doc.count("—") < 3:      # el guion donde el valor es cero
        fallos.append("faltan los guiones donde el valor es cero")

    # ---- 5. el mes que entra: solo la balanza, el histórico se conserva ---
    h2 = Historico(tmp / "historico.json")
    igual("el histórico se releyó completo", len(h2.periodos()), 6, 0)
    antes = dict(h2.datos["2025-12-31"]["cnsf"])
    h2.procesar(balanza)          # como si fuera la balanza del mes siguiente
    igual("cortes tras recargar la balanza", len(h2.periodos()), 6, 0)
    igual("un corte anterior quedó intacto", h2.datos["2025-12-31"]["cnsf"]["rrc"],
          antes["rrc"], 0.005)

    print(f"{pruebas} comprobaciones · {len(fallos)} fallo(s)")
    if fallos:
        for f in fallos:
            print("  ✗", f)
        return 1
    print("Todo cuadra. HTML de muestra:", ruta)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
