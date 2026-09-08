"""Solucion de referencia - Ejercicio 1 (solo libreria estandar)."""


def resumen_por_estado(polizas, siniestros):
    estado_de = {p["poliza_id"]: p["estado"] for p in polizas}

    acum = {}
    for p in polizas:
        d = acum.setdefault(p["estado"], {"n_polizas": 0, "n_siniestros": 0,
                                          "prima_total": 0.0, "monto_total": 0.0})
        d["n_polizas"] += 1
        d["prima_total"] += p["prima"]

    for s in siniestros:
        estado = estado_de.get(s["poliza_id"])
        if estado is None:          # regla 1: poliza inexistente -> se ignora
            continue
        d = acum[estado]
        d["n_siniestros"] += 1
        d["monto_total"] += s["monto"]

    resultado = []
    for estado, d in acum.items():
        n_pol, n_sin = d["n_polizas"], d["n_siniestros"]
        prima, monto = d["prima_total"], d["monto_total"]
        resultado.append({
            "estado": estado,
            "n_polizas": n_pol,
            "n_siniestros": n_sin,
            "prima_total": round(prima, 2),
            "monto_total": round(monto, 2),
            "frecuencia": round(n_sin / n_pol, 4) if n_pol else 0.0,
            "severidad": round(monto / n_sin, 4) if n_sin else 0.0,
            "loss_ratio": round(monto / prima, 4) if prima else 0.0,
        })

    resultado.sort(key=lambda r: (-r["loss_ratio"], r["estado"]))
    return resultado
