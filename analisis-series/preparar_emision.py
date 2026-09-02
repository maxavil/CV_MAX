# -*- coding: utf-8 -*-
"""
preparar_emision.py
-------------------
Saca de la base de emision (el parquet que arma consolidar_aseg.py, o los
spoolers ASEG_MM_AAAA.txt en crudo) SOLO los registros cuya poliza o cuya
clave de asegurado aparece en el Excel de Analisis de Series, y los deja en
un CSV chico listo para el HTML.

Sirve sobre todo para el parquet: el navegador no lo lee. Si vas a usar los
.txt tal cual, no hace falta este paso, el HTML los abre directo.

El CSV que genera tiene esta forma:
    POLIZA12,CLAVE_ASEGURADO,<las columnas que pediste...>

USO:
    python preparar_emision.py
    python preparar_emision.py --parquet "C:\\Spoolers\\consolidado\\aseg_consolidado.parquet"
    python preparar_emision.py --spoolers "C:\\Spoolers"
    python preparar_emision.py --columnas NOMBRE DIRECCION TELEFONO APODERADO_LEGAL
    python preparar_emision.py --listar-columnas
"""

import os
import re
import csv
import sys
import glob
import argparse
import unicodedata

from construir_resumen import leer_excel, buscar_excel, ORDEN, AQUI

# Un registro nuevo abre con la poliza. En el spooler son 12 digitos y un
# "|", pero segun como venga el relleno puede traer espacios de por medio,
# asi que se prueban varios y gana el que mas lineas explique. Si ninguno
# pasa del 50% se asume una linea = un registro.
PATRONES_INICIO = [
    re.compile(rb"^\d{12}\|"),
    re.compile(rb"^\s*\d{12}\s*\|"),
    re.compile(rb"^\s*\d{9,18}\s*\|"),
]

# Que columnas de la emision interesan por omision.
PATRONES = {
    "nombre":    re.compile(r"(NOMBRE|RAZON.?SOC)"),
    "direccion": re.compile(r"(DIRECC|DOMICIL|CALLE|COLONIA|MUNICIP|ESTADO|C_?P\b|CODIGO_?POSTAL)"),
    "telefono":  re.compile(r"(TELEF|CELUL|^TEL|_TEL|LADA)"),
    "apoderado": re.compile(r"(APODER|REPRESENT|LEGAL)"),
    "llaves":    re.compile(r"(POLIZA|CLAVE|ASEGURAD|INCISO|AGENTE|OFICINA|FECHA)"),
}


def llave_poliza(v):
    d = re.sub(r"\D+", "", str(v or ""))
    if not d:
        return ""
    return d[:12] if len(d) >= 12 else d.zfill(12)


def llave_clave(v):
    d = re.sub(r"\D+", "", str(v or ""))
    if not d:
        return ""
    return d[-10:] if len(d) >= 10 else d.zfill(10)


def sin_llave(v):
    return not v or not re.search(r"[1-9]", v)


def limpiar(s):
    s = str(s or "").strip()
    return "".join(c for c in s if unicodedata.category(c)[0] != "C")


def normalizar_nombre(c, i, vistos):
    base = re.sub(r"[^\w]+", "_", limpiar(c)).strip("_").upper() or ("COL_%d" % (i + 1))
    if base in vistos:
        vistos[base] += 1
        base = "%s_%d" % (base, vistos[base])
    else:
        vistos[base] = 1
    return base


def adivinar_llaves(columnas, muestra, pol_ok, cla_ok):
    """Elige la columna de poliza y la de clave por cuantos valores caen
    dentro de las llaves del Excel; el nombre solo desempata."""
    def puntua(j, fn, llaves):
        hit = 0
        for fila in muestra:
            v = fila[j] if j < len(fila) else ""
            if v and fn(v) in llaves:
                hit += 1
        return hit

    mejor_p = max(range(len(columnas)), key=lambda j: (puntua(j, llave_poliza, pol_ok),
                                                       bool(re.search(r"POLIZ", columnas[j]))))
    mejor_c = max(range(len(columnas)), key=lambda j: (puntua(j, llave_clave, cla_ok),
                                                       bool(re.search(r"CLAVE|ASEGURAD", columnas[j]))))
    return columnas[mejor_p], columnas[mejor_c]


def elegir_columnas(columnas, pedidas, llaves=()):
    """Las dos llaves ya van como POLIZA12 y CLAVE_ASEGURADO, no se repiten."""
    columnas = [c for c in columnas if c not in llaves]
    if pedidas:
        faltan = [c for c in pedidas if c not in columnas]
        if faltan:
            print("  ! no existen en la emision y se ignoran: %s" % ", ".join(faltan))
        return [c for c in pedidas if c in columnas]
    sel = [c for c in columnas if any(p.search(c) for p in PATRONES.values())]
    return sel or list(columnas)


# --------------------------- fuentes ---------------------------

def desde_parquet(ruta, pol_ok, cla_ok, pedidas):
    try:
        import pyarrow.parquet as pq
    except ImportError:
        sys.exit("\nFalta pyarrow para leer el parquet:\n    pip install pyarrow\n")

    pf = pq.ParquetFile(ruta)
    columnas = list(pf.schema_arrow.names)
    muestra = []
    for lote in pf.iter_batches(batch_size=2000):
        d = lote.to_pydict()
        muestra = [[str(d[c][i] or "") for c in columnas] for i in range(min(500, lote.num_rows))]
        break
    col_p, col_c = adivinar_llaves(columnas, muestra, pol_ok, cla_ok)
    salida_cols = elegir_columnas(columnas, pedidas, (col_p, col_c))
    print("  llave poliza: %s | llave clave: %s" % (col_p, col_c))
    print("  %d columnas exportadas" % len(salida_cols))

    filas, leidas = [], 0
    for lote in pf.iter_batches(batch_size=100_000,
                                columns=list({col_p, col_c, *salida_cols})):
        d = lote.to_pydict()
        for i in range(lote.num_rows):
            leidas += 1
            p = llave_poliza(d[col_p][i])
            c = llave_clave(d[col_c][i])
            if (sin_llave(p) or p not in pol_ok) and (sin_llave(c) or c not in cla_ok):
                continue
            filas.append([p, c] + [limpiar(d[col][i]) for col in salida_cols])
        print("\r  %s filas leidas, %s conservadas" % (f"{leidas:,}", f"{len(filas):,}"),
              end="", flush=True)
    print()
    return salida_cols, filas, leidas, 0


def detectar_inicio(ruta, muestra=600):
    """Que patron abre un registro en ESTE archivo (o None si cada linea
    es un registro completo)."""
    crudas = []
    with open(ruta, "rb") as fh:
        fh.readline()                      # cabecera
        for cruda in fh:
            crudas.append(cruda.rstrip(b"\r\n"))
            if len(crudas) >= muestra:
                break
    if not crudas:
        return None
    mejor, mejor_tasa = None, 0.5
    for patron in PATRONES_INICIO:
        tasa = sum(1 for l in crudas if patron.match(l)) / len(crudas)
        if tasa > mejor_tasa:
            mejor, mejor_tasa = patron, tasa
    return mejor


def registros_txt(ruta, patron=None):
    """Reconstruye los registros partidos igual que consolidar_aseg.py."""
    buffer = None
    with open(ruta, "rb") as fh:
        cabecera = fh.readline().decode("latin-1").rstrip("\r\n")
        yield cabecera, True
        for cruda in fh:
            cruda = cruda.rstrip(b"\r\n")
            if patron is None:
                if cruda.strip():
                    yield cruda.decode("latin-1"), False
                continue
            if patron.match(cruda):
                if buffer is not None:
                    yield buffer.decode("latin-1"), False
                buffer = cruda
            elif buffer is None:
                buffer = cruda
            else:
                buffer += cruda
        if buffer is not None:
            yield buffer.decode("latin-1"), False


def desde_spoolers(carpeta, pol_ok, cla_ok, pedidas):
    archivos = sorted(glob.glob(os.path.join(carpeta, "ASEG_*.txt")))
    if not archivos:
        sys.exit("No encontre ningun ASEG_*.txt en %s" % carpeta)

    columnas = salida_cols = None
    jp = jc = None
    jsel = []
    filas, leidas, rechazadas = [], 0, 0

    def guarda(campos):
        p, c = llave_poliza(campos[jp]), llave_clave(campos[jc])
        if (sin_llave(p) or p not in pol_ok) and (sin_llave(c) or c not in cla_ok):
            return
        filas.append([p, c] + [campos[j] for j in jsel])

    for ruta in archivos:
        patron = detectar_inicio(ruta)
        print("  %s  (%s)" % (os.path.basename(ruta),
              "reconstruyendo registros partidos" if patron else "una linea = un registro"),
              flush=True)
        buffer = []                      # registros previos a fijar las llaves
        for texto, es_cabecera in registros_txt(ruta, patron):
            if es_cabecera:
                vistos = {}
                cols = [normalizar_nombre(c, i, vistos) for i, c in enumerate(texto.split("|"))]
                if columnas is None:
                    columnas = cols
                elif cols != columnas:
                    print("    ! encabezado distinto; uso el del primer archivo")
                continue

            campos = [limpiar(c) for c in texto.split("|")]
            leidas += 1
            if len(campos) != len(columnas):
                rechazadas += 1          # un "|" dentro del dato: mismo criterio
                continue                 # que consolidar_aseg.py

            if jp is None:
                buffer.append(campos)
                if len(buffer) < 300:
                    continue
                col_p, col_c = adivinar_llaves(columnas, buffer, pol_ok, cla_ok)
                salida_cols = elegir_columnas(columnas, pedidas, (col_p, col_c))
                jp, jc = columnas.index(col_p), columnas.index(col_c)
                jsel = [columnas.index(c) for c in salida_cols]
                print("    llave poliza: %s | llave clave: %s" % (col_p, col_c))
                print("    %d columnas exportadas" % len(salida_cols))
                for previo in buffer:
                    guarda(previo)
                buffer = []
                continue
            guarda(campos)

        if jp is None and buffer:        # archivo con menos de 300 registros
            col_p, col_c = adivinar_llaves(columnas, buffer, pol_ok, cla_ok)
            salida_cols = elegir_columnas(columnas, pedidas, (col_p, col_c))
            jp, jc = columnas.index(col_p), columnas.index(col_c)
            jsel = [columnas.index(c) for c in salida_cols]
            print("    llave poliza: %s | llave clave: %s" % (col_p, col_c))
            for previo in buffer:
                guarda(previo)
        print("    %s leidas | %s conservadas" % (f"{leidas:,}", f"{len(filas):,}"))

    if jp is None:
        sys.exit("Los archivos no traen registros utilizables.")
    return salida_cols, filas, leidas, rechazadas


# --------------------------- principal ---------------------------

def main():
    descargas = os.path.join(os.path.expanduser("~"), "Downloads")
    ap = argparse.ArgumentParser(description="Extrae de la emision solo lo que cruza con el analisis")
    ap.add_argument("--excel", default=None)
    ap.add_argument("--parquet", default=r"C:\Spoolers\consolidado\aseg_consolidado.parquet")
    ap.add_argument("--spoolers", default=None, help="carpeta con los ASEG_*.txt en crudo")
    ap.add_argument("--salida", default=None)
    ap.add_argument("--columnas", nargs="*", default=None,
                    help="columnas de la emision a exportar (por omision, las utiles)")
    ap.add_argument("--listar-columnas", action="store_true")
    args = ap.parse_args()

    excel = args.excel or buscar_excel([
        os.path.join(descargas, "*nalisis*eries*.xls*"),
        os.path.join(AQUI, "*.xlsx"),
    ])
    if not excel or not os.path.exists(excel):
        sys.exit("No encontre el Excel. Pasalo con --excel \"ruta\\Analisis_de_Series.xlsx\"")

    print("Excel: %s" % excel)
    datos = leer_excel(excel)
    ip, ic = ORDEN.index("POLIZA"), ORDEN.index("CLAVE")
    pol_ok = {llave_poliza(f[ip]) for f in datos["filas"]}
    cla_ok = {llave_clave(f[ic]) for f in datos["filas"]}
    pol_ok = {v for v in pol_ok if not sin_llave(v)}
    cla_ok = {v for v in cla_ok if not sin_llave(v)}
    print("  %d polizas y %d claves a buscar\n" % (len(pol_ok), len(cla_ok)))

    if args.spoolers:
        print("Emision: spoolers en %s" % args.spoolers)
        cols, filas, leidas, rech = desde_spoolers(args.spoolers, pol_ok, cla_ok, args.columnas)
    else:
        if not os.path.exists(args.parquet):
            sys.exit("No existe %s. Usa --parquet o --spoolers." % args.parquet)
        print("Emision: %s" % args.parquet)
        cols, filas, leidas, rech = desde_parquet(args.parquet, pol_ok, cla_ok, args.columnas)

    if args.listar_columnas:
        print("\nColumnas exportadas:")
        for c in cols:
            print("  %s" % c)
        return

    salida = args.salida or os.path.join(
        descargas if os.path.isdir(descargas) else AQUI, "emision_para_html.csv")
    with open(salida, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["POLIZA12", "CLAVE_ASEGURADO"] + cols)
        w.writerows(filas)

    print("\n%s filas leidas | %s conservadas | %s rechazadas" %
          (f"{leidas:,}", f"{len(filas):,}", f"{rech:,}"))
    print("Listo -> %s  (%.1f KB)" % (salida, os.path.getsize(salida) / 1024))
    print("\nAhora puedes:")
    print("  a) arrastrar ese CSV al HTML, o")
    print("  b) incrustarlo:  python construir_resumen.py --emision \"%s\"" % salida)


if __name__ == "__main__":
    main()
