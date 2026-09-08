# CLAVE Y RUBRICA — Uso exclusivo del entrevistador
### NO compartir con el candidato

---

## Resumen de puntuacion

| Parte | Puntos |
|---|---|
| A — Opcion multiple (4 × 5) | 20 |
| B — Preguntas abiertas (2 × 15) | 30 |
| C — Programacion (2 × 25) | 50 |
| **Total** | **100** |

### Interpretacion sugerida

| Puntaje | Lectura |
|---|---|
| **85 – 100** | Fuerte. Codigo limpio y correcto, criterio de negocio y de validacion solido. Contratar / pasar a siguiente etapa con confianza. |
| **70 – 84** | Solido. Resuelve lo tecnico, quiza con huecos en negocio o en casos borde. Buen candidato; profundizar en entrevista. |
| **55 – 69** | Intermedio. Programa, pero se le escapan detalles importantes (fuga de informacion, casos borde, metricas). Considerar para perfil junior. |
| **< 55** | No cumple el perfil. |

> **Sobre el tiempo:** la prueba dura **30 minutos** y el contenido esta calibrado para exigir priorizacion. Es normal que un buen candidato no termine todo. Pondera el **criterio y el enfoque** por encima de la cobertura: codigo parcial pero bien encaminado vale mas que codigo completo sin manejo de casos borde.

> **Senal de alerta independiente del puntaje:** que en la Pregunta 2 no detecte la fuga de informacion, o que en la Parte C ignore por completo los casos borde (divisiones entre cero, registros invalidos). Es lo que mas cuesta en produccion.

---

## PARTE A — Clave (5 pts cada una, sin puntos parciales)

| Pregunta | Respuesta correcta | Por que |
|---|---|---|
| **1** | **B** | Con 0.9% de positivos, predecir siempre la clase mayoritaria da 99.1% de accuracy. La metrica no informa nada. Hay que ir a precision/recall, PR-AUC y al costo del error. **A** y **C** confunden accuracy con calidad; **D** confunde el sintoma (metrica mal elegida) con la solucion, y ademas rebalancear a 50/50 distorsiona las probabilidades predichas. |
| **2** | **B** | Doble error clasico: *target leakage* (`monto_pagado_siniestro` solo existe si ya hubo siniestro) + particion aleatoria en un problema con estructura temporal. El AUC 0.97 es artificial. |
| **3** | **B** | `LEFT JOIN` conserva estados sin siniestros, `COALESCE` los deja en 0 y `COUNT(DISTINCT p.poliza_id)` evita inflar el conteo de polizas por el fan-out del join. **A** excluye esos estados y duplica polizas; **C** anula el LEFT JOIN al filtrar en `WHERE` sobre la tabla derecha; **D** invierte el sentido del join. |
| **4** | **B** | Brecha de ~3x entre entrenamiento y validacion = sobreajuste. La accion es reducir capacidad y aumentar regularizacion con early stopping. **A** empeora el problema; **C** es irrelevante; **D** normaliza un mal resultado. |

---

## PARTE B — Rubrica (15 pts cada una)

Puntua cada inciso de 0 a 4 (Pregunta 5) o segun lo indicado (Pregunta 6). No se busca la respuesta "perfecta", sino **criterio**.

### Pregunta 5 — Diseno de la solucion (15 pts)

| Inciso | Pts | Que debe aparecer para el puntaje completo |
|---|---|---|
| **1. Datos y variables** | 4 | Variables del vehiculo (marca, modelo, anio, valor, uso), del conductor (edad, antiguedad de licencia, historial de siniestros previos), geograficas (estado, CP, zona de riesgo), de la poliza (cobertura, deducible, prima, antiguedad, canal de venta). **Clave:** debe decir que **evitaria variables posteriores al hecho** (montos pagados, fecha de reporte, ajustador asignado) y ser consciente de **variables sensibles o proxies** (genero, ingreso, CP como proxy socioeconomico) por riesgo regulatorio y reputacional. |
| **2. Dataset y validacion** | 4 | Definir la **fecha de corte** y la ventana: features con informacion disponible hasta t, objetivo = siniestro en (t, t+12m]. **Particion temporal** (out-of-time), no aleatoria. Menciona evitar polizas con exposicion incompleta o ajustar por exposicion. Suma si menciona validacion cruzada temporal o backtesting en varios cortes. |
| **3. Metricas** | 4 | Tecnicas: AUC / PR-AUC, **calibracion** (curva de calibracion, Brier) — critico si el output alimenta tarifas —, lift por decil. Negocio: loss ratio por decil de riesgo, prima esperada vs. siniestralidad observada, impacto en suscripcion o retencion, % de cartera afectada. |
| **4. Riesgos** | 3 | Dos riesgos concretos con mitigacion. Validos: calidad/disponibilidad del dato, deriva del modelo, sesgo y cumplimiento regulatorio (CNSF), adopcion por el area usuaria, cambio de mezcla de cartera, seleccion adversa. |

**Restar/no otorgar:** respuestas que saltan directo a "uso XGBoost y listo" sin hablar de definicion del objetivo, ventana temporal ni metricas de negocio.

### Pregunta 6 — Modelo degradado (15 pts)

| Inciso | Pts | Que debe aparecer |
|---|---|---|
| **1. Diagnostico** | 5 | Ordena antes de opinar: (a) confirmar que la caida es real y no un problema de medicion o de etiquetas incompletas — el fraude se confirma con rezago; (b) revisar el **pipeline de datos** (campos nulos, cambios de esquema, unidades, un catalogo que cambio); (c) comparar **distribuciones** de features y del score entre el periodo de entrenamiento y el actual (data drift) y la relacion feature-objetivo (concept drift); (d) desglosar el desempeno por segmento, estado, canal y tipo de cobertura para ver si la caida es global o localizada. Pediria: logs de scoring con features, etiquetas confirmadas con su fecha, historial de despliegues y de cambios en sistemas fuente. |
| **2. Hipotesis** | 4 | Al menos tres, por ejemplo: cambio en el pipeline o en la fuente de datos; deriva de poblacion (nueva mezcla de cartera, nuevo canal, nueva region); deriva de concepto (los defraudadores cambiaron de tactica, incluso adaptandose al propio modelo); umbral fijo con score descalibrado; **retroalimentacion sesgada** — solo se investiga lo que el modelo alerta, asi que las etiquetas nuevas estan sesgadas; estacionalidad o un evento puntual. |
| **3. Decision** | 4 | Debe **ligar la accion a la causa**, no elegir al azar: si el score sigue ordenando bien pero cambio la prevalencia o la capacidad operativa → **recalibrar umbral** (rapido, reversible, se decide con la capacidad de investigacion del area y el costo por caso); si hay deriva de datos con la misma estructura del problema → **reentrenar** con datos recientes; si cambio el fenomeno o hay fuga/variables muertas → **redisenar** (features nuevas, replantear el objetivo). Suma si menciona corregir primero el pipeline si ahi esta el error, o comparar contra un baseline. |
| **4. Comunicacion** | 2 | Honesto y accionable: explicar en lenguaje de negocio que la precision cayo, dar una expectativa realista (cuantas alertas por semana y que % vale la pena revisar), ajustar el umbral para no quemar al equipo, priorizar por monto expuesto y fijar fecha de revision. Reconoce el costo de perder la confianza del usuario. |

**Excelente respuesta:** menciona el **ciclo de retroalimentacion** (solo se etiqueta lo que se alerta) y que el desempeno se debe monitorear de forma continua, no descubrir por queja.

---

## PARTE C — Rubrica de codigo (25 pts cada uno)

### Como corregir

```bash
cd prueba_tecnica_ciencia_datos
python3 ejercicio_1.py     # debe imprimir "Todas las pruebas pasaron."
python3 ejercicio_2.py
```

Las soluciones de referencia estan en `soluciones/`:

- `ejercicio_1_solucion.py`
- `ejercicio_2_solucion.py` (libreria estandar)
- `ejercicio_2_solucion_pandas.py` (pandas)

Las tres pasan las pruebas incluidas en las plantillas.

### Pregunta 7 — `ejercicio_1.py` (25 pts)

| Criterio | Pts |
|---|---|
| Las pruebas pasan | 10 |
| Ignora correctamente siniestros de polizas inexistentes (regla 1) | 3 |
| Maneja estados sin siniestros y evita divisiones entre cero (reglas 2 y 3) | 4 |
| Ordenamiento correcto, incluido el desempate por estado (regla 5) | 3 |
| Calidad: agregacion en una sola pasada (no O(n·m) con busquedas anidadas), nombres claros, sin repeticion | 5 |

**Resultado esperado**

| estado | n_polizas | n_siniestros | prima_total | monto_total | frecuencia | severidad | loss_ratio |
|---|---|---|---|---|---|---|---|
| NUEVO LEON | 1 | 1 | 15000.0 | 21000.0 | 1.0 | 21000.0 | 1.4 |
| CDMX | 2 | 3 | 21500.0 | 22500.0 | 1.5 | 7500.0 | 1.0465 |
| JALISCO | 2 | 1 | 19000.0 | 3000.0 | 0.5 | 3000.0 | 0.1579 |
| PUEBLA | 1 | 0 | 7000.0 | 0.0 | 0.0 | 0.0 | 0.0 |

**Errores tipicos:** contar el siniestro de la poliza `P-999` (inflaria CDMX o tronaria con `KeyError`); omitir PUEBLA por no tener siniestros; `ZeroDivisionError` al calcular la severidad de PUEBLA; ordenar solo por loss ratio sin el desempate.

### Pregunta 8 — `ejercicio_2.py` (25 pts)

| Criterio | Pts |
|---|---|
| Las pruebas pasan | 10 |
| Parseo correcto del monto: "$", comas, espacios, vacios y no numericos (regla 2) | 4 |
| Maneja los **dos** formatos de fecha y descarta la fecha inexistente `31/02/2024` (regla 4) | 4 |
| Duplicados por `siniestro_id` conservando la primera aparicion (regla 1) | 3 |
| Calidad: usa el modulo `csv` o `pandas` en vez de partir por comas a mano (el archivo tiene comas dentro de comillas), codigo legible y sin numeros magicos | 4 |

**Resultado esperado**

```python
{
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
```

**De 27 filas de datos, 8 se descartan:**

| Fila | Motivo |
|---|---|
| `S-001` (2a aparicion) | Duplicado |
| `S-002` (2a aparicion) | Duplicado |
| `S-006` | Monto vacio |
| `S-007` | Monto negativo (-5000) |
| `S-011` | Fecha de 2023, fuera del periodo |
| `S-013` | Fecha inexistente: `31/02/2024` |
| `S-015` | Monto igual a 0 |
| `S-018` | Monto no numerico: `n/a` |

**Errores tipicos:** partir las lineas con `line.split(",")` (los montos entre comillas traen comas y se rompe todo); dejar pasar `31/02/2024` con un parseo manual; no distinguir `" puebla "` de `"PUEBLA"`; eliminar duplicados con la fila completa en vez de por `siniestro_id`; devolver `monto_por_estado` sin ordenar.

**Pregunta de seguimiento recomendada (oral, no puntua):**
*"El mes de mayor monto es abril, pero un solo siniestro de $45,000 explica casi la mitad de ese mes. ¿Reportarias abril como el peor mes? ¿Que le dirias al negocio?"*
Busca que note el efecto de los valores extremos en severidad, que proponga mirar la mediana, el conteo de siniestros o un monto capado, y que separe "mes con mas siniestros" de "mes con un siniestro muy caro". **Es la mejor senal de madurez de toda la prueba.**

---

## Guia para la sesion en vivo (opcional, 10 min al final)

Si haces la prueba con el candidato presente, estas preguntas separan a quien memorizo de quien entiende:

1. "Explicame tu solucion del ejercicio 2 en 2 minutos." — claridad al comunicar.
2. "Si el archivo tuviera 50 millones de filas, ¿que cambiarias?" — procesamiento por lotes, tipos de dato, base de datos, `chunksize`.
3. "¿Como probarias que tu limpieza es correcta si no te hubiera dado las pruebas?" — pruebas propias, controles de calidad, conteos de control.
4. "En el ejercicio 1, ¿que significa un loss ratio de 1.4 para el negocio?" — que por cada peso de prima se pagan 1.40 en siniestros: esa cartera pierde dinero antes de gastos.
