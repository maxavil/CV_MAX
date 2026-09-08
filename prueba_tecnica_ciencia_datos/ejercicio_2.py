"""
EJERCICIO 2 - Limpieza y analisis de un archivo de siniestros
(25 puntos | tiempo sugerido: 10 min)
=============================================================================

Contexto de negocio
-------------------
Te llega un extracto operativo de siniestros de auto (`datos/siniestros_2024.csv`).
Viene sucio: montos con simbolos, fechas en dos formatos, estados escritos de
mil maneras, registros duplicados y filas invalidas. Antes de que nadie tome
una decision con estos numeros, hay que limpiarlo.

Tarea
-----
Implementa `analizar_siniestros(ruta_csv)`. Puedes usar pandas o solo la
libreria estandar; ambas son validas y se evaluan igual.

Reglas de limpieza (aplicalas en este orden)
--------------------------------------------
1. DUPLICADOS: si un `siniestro_id` aparece mas de una vez, conserva
   unicamente la PRIMERA aparicion (orden del archivo).
2. MONTO: `monto_pagado` puede traer "$", comas y espacios (ej. " $8,300.50 ").
   Conviertelo a float. Descarta la fila si el monto esta vacio, no es
   numerico, o es menor o igual a 0.
3. ESTADO: normaliza quitando espacios al inicio/final y pasando a MAYUSCULAS
   (ej. " puebla " y "PUEBLA" son el mismo estado).
4. FECHA: `fecha_siniestro` viene en dos formatos, "YYYY-MM-DD" y "DD/MM/YYYY".
   Descarta la fila si la fecha no es valida en ninguno de los dos
   (ojo: hay fechas que no existen en el calendario).
5. PERIODO: conserva unicamente siniestros del anio 2024.

Salida
------
Un diccionario con exactamente estas llaves:

    {
        "n_siniestros":      int,    # filas que sobrevivieron la limpieza
        "monto_total":       float,  # suma de montos, redondeado a 2 decimales
        "ticket_promedio":   float,  # monto_total / n_siniestros, a 2 decimales
        "monto_por_estado":  dict,   # {"CDMX": 120500.0, ...} a 2 decimales,
                                     # ordenado de mayor a menor monto
        "mes_mayor_monto":   str,    # "YYYY-MM" con mayor monto acumulado
    }

Ejecuta este archivo para probar tu solucion:  python3 ejercicio_2.py
"""

import os

RUTA_CSV = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "datos", "siniestros_2024.csv")


def analizar_siniestros(ruta_csv):
    # TODO: implementa aqui tu solucion
    raise NotImplementedError


# ---------------------------------------------------------------------------
# NO MODIFIQUES NADA DEBAJO DE ESTA LINEA
# ---------------------------------------------------------------------------

ESPERADO = {
    "n_siniestros": 19,
    "monto_total": 281503.40,
    "ticket_promedio": 14815.97,
    "monto_por_estado": {
        "CDMX": 120450.00,
        "JALISCO": 67376.50,
        "PUEBLA": 55700.60,
        "NUEVO LEON": 37976.30,
    },
    "mes_mayor_monto": "2024-04",
}


def _correr_pruebas():
    obtenido = analizar_siniestros(RUTA_CSV)
    fallos = []

    if not isinstance(obtenido, dict):
        fallos.append(f"Se esperaba un dict, se obtuvo {type(obtenido).__name__}")
    else:
        for llave, valor in ESPERADO.items():
            if llave not in obtenido:
                fallos.append(f"falta la llave '{llave}'")
                continue
            obt = obtenido[llave]
            if llave == "monto_por_estado":
                if dict(obt) != valor:
                    fallos.append(f"[{llave}] esperado {valor}, obtenido {dict(obt)}")
                elif list(obt.keys()) != list(valor.keys()):
                    fallos.append(
                        f"[{llave}] el orden debe ser de mayor a menor monto: "
                        f"esperado {list(valor.keys())}, obtenido {list(obt.keys())}"
                    )
            elif obt != valor:
                fallos.append(f"[{llave}] esperado {valor!r}, obtenido {obt!r}")

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
