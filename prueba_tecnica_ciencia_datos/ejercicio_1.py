"""
EJERCICIO 1 - Siniestralidad por estado  (25 puntos | tiempo sugerido: 7 min)
=============================================================================

Contexto de negocio
-------------------
El area de suscripcion quiere saber en que estados de la Republica la cartera
de autos esta perdiendo dinero. Para eso te dan dos listas ya cargadas en
memoria: las polizas vigentes y los siniestros pagados.

Tarea
-----
Implementa `resumen_por_estado(polizas, siniestros)` usando SOLO la libreria
estandar de Python (no pandas, no numpy).

Entradas
--------
polizas: lista de diccionarios
    {"poliza_id": "P-001", "estado": "CDMX", "prima": 12000.0}

siniestros: lista de diccionarios
    {"siniestro_id": "S-001", "poliza_id": "P-001", "monto": 4500.0}

Salida
------
Lista de diccionarios, UNO POR ESTADO presente en `polizas`, con las llaves:

    {
        "estado":       str,
        "n_polizas":    int,    # polizas del estado
        "n_siniestros": int,    # siniestros de esas polizas
        "prima_total":  float,  # suma de primas del estado
        "monto_total":  float,  # suma de siniestros del estado
        "frecuencia":   float,  # n_siniestros / n_polizas
        "severidad":    float,  # monto_total / n_siniestros
        "loss_ratio":   float,  # monto_total / prima_total
    }

Reglas
------
1. Un siniestro cuyo `poliza_id` NO exista en `polizas` se ignora
   (son siniestros de polizas ya canceladas y no deben contarse).
2. Un estado sin siniestros debe aparecer igual, con n_siniestros = 0,
   monto_total = 0.0, frecuencia = 0.0, severidad = 0.0, loss_ratio = 0.0.
3. Si `prima_total` es 0, `loss_ratio` debe ser 0.0 (no truena el programa).
4. Redondea `frecuencia`, `severidad` y `loss_ratio` a 4 decimales
   (usa round(x, 4)). `prima_total` y `monto_total` a 2 decimales.
5. Ordena el resultado por `loss_ratio` DESCENDENTE y, en caso de empate,
   por `estado` ascendente (A-Z).

Ejecuta este archivo para probar tu solucion:  python3 ejercicio_1.py
"""


def resumen_por_estado(polizas, siniestros):
    # TODO: implementa aqui tu solucion
    raise NotImplementedError


# ---------------------------------------------------------------------------
# NO MODIFIQUES NADA DEBAJO DE ESTA LINEA
# ---------------------------------------------------------------------------

POLIZAS = [
    {"poliza_id": "P-001", "estado": "CDMX", "prima": 12000.0},
    {"poliza_id": "P-002", "estado": "CDMX", "prima": 9500.0},
    {"poliza_id": "P-003", "estado": "JALISCO", "prima": 8000.0},
    {"poliza_id": "P-004", "estado": "JALISCO", "prima": 11000.0},
    {"poliza_id": "P-005", "estado": "NUEVO LEON", "prima": 15000.0},
    {"poliza_id": "P-006", "estado": "PUEBLA", "prima": 7000.0},
]

SINIESTROS = [
    {"siniestro_id": "S-01", "poliza_id": "P-001", "monto": 4000.0},
    {"siniestro_id": "S-02", "poliza_id": "P-001", "monto": 6500.0},
    {"siniestro_id": "S-03", "poliza_id": "P-002", "monto": 12000.0},
    {"siniestro_id": "S-04", "poliza_id": "P-003", "monto": 3000.0},
    {"siniestro_id": "S-05", "poliza_id": "P-005", "monto": 21000.0},
    {"siniestro_id": "S-06", "poliza_id": "P-999", "monto": 50000.0},  # poliza inexistente
]

ESPERADO = [
    {
        "estado": "NUEVO LEON", "n_polizas": 1, "n_siniestros": 1,
        "prima_total": 15000.0, "monto_total": 21000.0,
        "frecuencia": 1.0, "severidad": 21000.0, "loss_ratio": 1.4,
    },
    {
        "estado": "CDMX", "n_polizas": 2, "n_siniestros": 3,
        "prima_total": 21500.0, "monto_total": 22500.0,
        "frecuencia": 1.5, "severidad": 7500.0, "loss_ratio": 1.0465,
    },
    {
        "estado": "JALISCO", "n_polizas": 2, "n_siniestros": 1,
        "prima_total": 19000.0, "monto_total": 3000.0,
        "frecuencia": 0.5, "severidad": 3000.0, "loss_ratio": 0.1579,
    },
    {
        "estado": "PUEBLA", "n_polizas": 1, "n_siniestros": 0,
        "prima_total": 7000.0, "monto_total": 0.0,
        "frecuencia": 0.0, "severidad": 0.0, "loss_ratio": 0.0,
    },
]


def _correr_pruebas():
    obtenido = resumen_por_estado(POLIZAS, SINIESTROS)
    fallos = []

    if not isinstance(obtenido, list):
        fallos.append(f"Se esperaba una lista, se obtuvo {type(obtenido).__name__}")
    elif len(obtenido) != len(ESPERADO):
        fallos.append(f"Se esperaban {len(ESPERADO)} estados, se obtuvieron {len(obtenido)}")
    else:
        for i, (esp, obt) in enumerate(zip(ESPERADO, obtenido)):
            for llave, valor in esp.items():
                if llave not in obt:
                    fallos.append(f"[fila {i}] falta la llave '{llave}'")
                elif obt[llave] != valor:
                    fallos.append(
                        f"[fila {i}][{llave}] esperado {valor!r}, obtenido {obt[llave]!r}"
                    )

    if fallos:
        print("PRUEBAS FALLIDAS:")
        for f in fallos:
            print("  -", f)
    else:
        print("Todas las pruebas pasaron.")
    return not fallos


if __name__ == "__main__":
    import sys
    sys.exit(0 if _correr_pruebas() else 1)
