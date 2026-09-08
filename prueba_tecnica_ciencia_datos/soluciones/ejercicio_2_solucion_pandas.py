"""Solucion de referencia alterna - Ejercicio 2 (con pandas)."""

import pandas as pd


def analizar_siniestros(ruta_csv):
    df = pd.read_csv(ruta_csv, dtype=str)

    # regla 1: duplicados por siniestro_id, conservando la primera aparicion
    df = df.drop_duplicates(subset="siniestro_id", keep="first")

    # regla 2: monto limpio y positivo
    df["monto"] = pd.to_numeric(
        df["monto_pagado"].str.replace(r"[$,\s]", "", regex=True), errors="coerce"
    )
    df = df[df["monto"].notna() & (df["monto"] > 0)]

    # regla 3: estado normalizado
    df["estado"] = df["estado"].str.strip().str.upper()

    # regla 4: fecha en cualquiera de los dos formatos
    f1 = pd.to_datetime(df["fecha_siniestro"], format="%Y-%m-%d", errors="coerce")
    f2 = pd.to_datetime(df["fecha_siniestro"], format="%d/%m/%Y", errors="coerce")
    df["fecha"] = f1.fillna(f2)
    df = df[df["fecha"].notna()]

    # regla 5: solo 2024
    df = df[df["fecha"].dt.year == 2024]

    monto_total = float(df["monto"].sum())
    n = int(len(df))

    por_estado = df.groupby("estado")["monto"].sum().sort_values(ascending=False)
    por_mes = df.groupby(df["fecha"].dt.strftime("%Y-%m"))["monto"].sum()

    return {
        "n_siniestros": n,
        "monto_total": round(monto_total, 2),
        "ticket_promedio": round(monto_total / n, 2) if n else 0.0,
        "monto_por_estado": {e: round(float(m), 2) for e, m in por_estado.items()},
        "mes_mayor_monto": str(por_mes.idxmax()) if n else "",
    }


if __name__ == "__main__":
    import json
    import os
    ruta = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                        "datos", "siniestros_2024.csv")
    print(json.dumps(analizar_siniestros(ruta), indent=2, ensure_ascii=False))
