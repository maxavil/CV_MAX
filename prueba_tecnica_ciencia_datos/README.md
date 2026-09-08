# Prueba tecnica de Ciencia de Datos — Qualitas

Examen de **60 minutos** y **100 puntos** para evaluar candidatos a Ciencia de Datos.
Todo el contenido esta ambientado en seguros de auto: siniestralidad, fraude, tarificacion.

## Estructura

| Parte | Contenido | Preguntas | Puntos | Tiempo |
|---|---|---|---|---|
| A | Opcion multiple | 4 | 20 | 8 min |
| B | Preguntas abiertas | 2 | 30 | 15 min |
| C | Programacion en Python | 2 | 50 | 35 min |
| | | **8** | **100** | **58 min** |

## Archivos

```
prueba_tecnica_ciencia_datos/
├── EXAMEN.md                        <- se le entrega al candidato
├── plantilla_respuestas.md          <- ahi escribe las partes A y B
├── ejercicio_1.py                   <- Pregunta 7 (plantilla + pruebas)
├── ejercicio_2.py                   <- Pregunta 8 (plantilla + pruebas)
├── datos/
│   └── siniestros_2024.csv          <- datos "sucios" del ejercicio 2
├── EVALUADOR_clave_y_rubrica.md     <- NO ENTREGAR AL CANDIDATO
└── soluciones/                      <- NO ENTREGAR AL CANDIDATO
    ├── ejercicio_1_solucion.py
    ├── ejercicio_2_solucion.py            (libreria estandar)
    └── ejercicio_2_solucion_pandas.py     (pandas)
```

## Como aplicarla

**Antes:**

1. Copia a la maquina del candidato **solo**: `EXAMEN.md`, `plantilla_respuestas.md`, `ejercicio_1.py`, `ejercicio_2.py` y `datos/`.
   **Retira `EVALUADOR_clave_y_rubrica.md` y la carpeta `soluciones/`.**
2. Verifica que tenga Python 3.8+ (`python3 --version`). pandas es opcional: el ejercicio 2 se puede resolver con la libreria estandar.

**Durante:** el candidato puede consultar documentacion oficial; no asistentes de IA. Cada ejercicio de codigo trae sus pruebas: se ejecuta el archivo y se ve si pasan.

**Despues:**

```bash
python3 ejercicio_1.py     # "Todas las pruebas pasaron."
python3 ejercicio_2.py
```

Califica con `EVALUADOR_clave_y_rubrica.md`, que incluye la clave, la rubrica punto por punto, los errores tipicos de cada ejercicio y preguntas de seguimiento para la sesion en vivo.

## Que evalua cada pieza

| Pregunta | Habilidad |
|---|---|
| 1 | No dejarse enganar por el accuracy con clases desbalanceadas |
| 2 | Detectar fuga de informacion y exigir validacion temporal |
| 3 | SQL: joins, agregacion y el error clasico del `WHERE` sobre un `LEFT JOIN` |
| 4 | Leer la brecha entrenamiento/validacion y actuar en consecuencia |
| 5 | Disenar un proyecto de punta a punta con criterio de negocio y de riesgo |
| 6 | Diagnosticar un modelo degradado en produccion y comunicarlo |
| 7 | Python puro: agregacion, casos borde, divisiones entre cero, ordenamiento |
| 8 | Datos sucios del mundo real: parseo, fechas, duplicados, normalizacion |

Las tres soluciones de referencia pasan las pruebas incluidas en las plantillas.
