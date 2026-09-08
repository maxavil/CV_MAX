"""Solucion de referencia - Ejercicio 2 (solo libreria estandar)."""

import csv
from collections import OrderedDict
from datetime import datetime

FORMATOS = ("%Y-%m-%d", "%d/%m/%Y")


def _a_float(texto):
    if texto is None:
        return None
    limpio = texto.replace("$", "").replace(",", "").strip()
    if not limpio:
        return None
    try:
        return float(limpio)
    except ValueError:
        return None


def _a_fecha(texto):
    texto = (texto or "").strip()
    for fmt in FORMATOS:
        try:
            return datetime.strptime(texto, fmt)
        except ValueError:
            continue
    return None


def analizar_siniestros(ruta_csv):
    vistos = set()
    filas = []

    with open(ruta_csv, newline="", encoding="utf-8") as f:
        for fila in csv.DictReader(f):
            sid = (fila["siniestro_id"] or "").strip()
            if sid in vistos:                      # regla 1: duplicados
                continue
            vistos.add(sid)

            monto = _a_float(fila["monto_pagado"])  # regla 2: monto
            if monto is None or monto <= 0:
                continue

            fecha = _a_fecha(fila["fecha_siniestro"])  # regla 4: fecha
            if fecha is None or fecha.year != 2024:    # regla 5: solo 2024
                continue

            filas.append({
                "estado": (fila["estado"] or "").strip().upper(),  # regla 3
                "monto": monto,
                "mes": fecha.strftime("%Y-%m"),
            })

    monto_total = sum(f["monto"] for f in filas)
    n = len(filas)

    por_estado = {}
    por_mes = {}
    for f in filas:
        por_estado[f["estado"]] = por_estado.get(f["estado"], 0.0) + f["monto"]
        por_mes[f["mes"]] = por_mes.get(f["mes"], 0.0) + f["monto"]

    monto_por_estado = OrderedDict(
        (e, round(m, 2)) for e, m in sorted(por_estado.items(), key=lambda kv: -kv[1])
    )
    mes_mayor = max(por_mes.items(), key=lambda kv: kv[1])[0] if por_mes else ""

    return {
        "n_siniestros": n,
        "monto_total": round(monto_total, 2),
        "ticket_promedio": round(monto_total / n, 2) if n else 0.0,
        "monto_por_estado": monto_por_estado,
        "mes_mayor_monto": mes_mayor,
    }


if __name__ == "__main__":
    import json
    import os
    ruta = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                        "datos", "siniestros_2024.csv")
    print(json.dumps(analizar_siniestros(ruta), indent=2, ensure_ascii=False))
