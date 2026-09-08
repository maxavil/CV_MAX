# Prueba tecnica de Ciencia de Datos — Qualitas

Examen de **30 minutos** y **100 puntos** para evaluar candidatos a Ciencia de Datos.
Todo el contenido esta ambientado en seguros de auto: siniestralidad, fraude, tarificacion.

## Estructura

| Parte | Contenido | Preguntas | Puntos | Tiempo |
|---|---|---|---|---|
| A | Opcion multiple | 4 | 20 | 5 min |
| B | Preguntas abiertas | 2 | 30 | 8 min |
| C | Programacion en Python | 2 | 50 | 17 min |
| | | **8** | **100** | **30 min** |

## Archivos

```
prueba_tecnica_ciencia_datos/
├── EXAMEN_Qualitas.pdf              <- version imprimible con formato institucional
├── EXAMEN.md                        <- fuente del examen (se le entrega al candidato)
├── plantilla_respuestas.md          <- ahi escribe las partes A y B
├── ejercicio_1.py                   <- Pregunta 7 (plantilla + pruebas)
├── ejercicio_2.py                   <- Pregunta 8 (plantilla + pruebas)
├── datos/
│   └── siniestros_2024.csv          <- datos "sucios" del ejercicio 2
├── generar_pdf.py                   <- regenera el PDF a partir de EXAMEN.md
├── assets/
│   └── logo_qualitas.png            <- logo institucional usado en el encabezado
├── EVALUADOR_clave_y_rubrica.md     <- NO ENTREGAR AL CANDIDATO
└── soluciones/                      <- NO ENTREGAR AL CANDIDATO
    ├── ejercicio_1_solucion.py
    ├── ejercicio_2_solucion.py            (libreria estandar)
    └── ejercicio_2_solucion_pandas.py     (pandas)
```

## Version imprimible

`EXAMEN_Qualitas.pdf` es el examen con el formato institucional de Qualitas (logo,
morado `#8E1C7A` y turquesa `#1192A5`, encabezado y pie en cada pagina), en tamano
Carta y 4 paginas. Incluye casillas para marcar las respuestas de opcion multiple
y campos de nombre, fecha y hora de inicio/termino.

Si editas `EXAMEN.md`, regenera el PDF con:

```bash
pip install markdown playwright
python3 generar_pdf.py
```

El script busca Chromium automaticamente; si hace falta, indicale la ruta:
`CHROMIUM_PATH=/ruta/a/chrome python3 generar_pdf.py`.

## Como aplicarla

**Antes:**

1. Imprime `EXAMEN_Qualitas.pdf` (4 paginas, tamano Carta) o entrega `EXAMEN.md`.
2. Copia a la maquina del candidato **solo**: `EXAMEN.md`, `plantilla_respuestas.md`, `ejercicio_1.py`, `ejercicio_2.py` y `datos/`.
   **Retira `EVALUADOR_clave_y_rubrica.md`, `soluciones/` y `generar_pdf.py`.**
3. Verifica que tenga Python 3.8+ (`python3 --version`). pandas es opcional: el ejercicio 2 se puede resolver con la libreria estandar.

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
