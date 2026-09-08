# Prueba técnica — Ciencia de Datos
### Quálitas | Duración: 30 minutos | Total: 100 puntos

**Nombre del candidato:** __________  **Fecha:** __________  **Hora de inicio:** __________  **Hora de término:** __________

---

## Instrucciones

- Tienes **30 minutos**. Administra tu tiempo: los ejercicios de programación valen la mitad del examen.
- Puedes consultar documentación oficial (Python, pandas, SQL). **No** uses asistentes de IA.
- Escribe tus respuestas de opción múltiple y abiertas en `plantilla_respuestas.md`.
- El código va directo en `ejercicio_1.py` y `ejercicio_2.py`.
- Si no terminas un ejercicio, **deja comentado tu razonamiento**: se evalúa también el enfoque.
- El tiempo es ajustado a propósito. Si vas contra reloj, **prioriza la Parte C**: vale la mitad del examen.

| Parte | Contenido | Preguntas | Puntos | Tiempo sugerido |
|---|---|---|---|---|
| A | Opción múltiple | 1 – 4 | 20 | 5 min |
| B | Preguntas abiertas | 5 – 6 | 30 | 8 min |
| C | Programación | 7 – 8 | 50 | 17 min |

---

# PARTE A — Opción múltiple

**4 preguntas × 5 puntos = 20 puntos.** Marca **una sola** opción por pregunta.

---

### Pregunta 1 (5 pts) — Evaluación con clases desbalanceadas

Entrenas un modelo para detectar siniestros con posible fraude. De 200,000 siniestros históricos, 1,800 (0.9%) fueron confirmados como fraude. Tu modelo alcanza **99.1% de accuracy** en el conjunto de prueba.

**¿Cuál es la interpretación correcta?**

- **A)** El modelo es excelente y puede pasar a producción tal cual.
- **B)** El accuracy es engañoso: un modelo que siempre prediga "no fraude" lograría ~99.1%. Hay que evaluar con precision, recall, PR-AUC y el costo económico de falsos positivos vs. falsos negativos.
- **C)** El accuracy tan alto indica sobreajuste; hay que simplificar el modelo.
- **D)** Basta con balancear las clases a 50/50 con oversampling; después de eso el accuracy sí será confiable.

---

### Pregunta 2 (5 pts) — Validación y fuga de información

Construyes un modelo para predecir si una póliza de auto tendrá **al menos un siniestro en los próximos 12 meses**. Tienes datos de 2019 a 2024. Partes el dataset **aleatoriamente** 80/20 e incluyes entre las variables predictoras el campo `monto_pagado_siniestro`. Obtienes **AUC 0.97** en validación, pero en producción el desempeño cae a **0.61**.

**¿Cuál es la causa MÁS probable?**

- **A)** Faltó estandarizar las variables numéricas antes de entrenar.
- **B)** Hay dos problemas: (i) fuga de información, porque `monto_pagado_siniestro` solo se conoce *después* de que ocurrió el siniestro y es prácticamente la variable objetivo disfrazada; y (ii) la partición debió ser **temporal** (entrenar con el pasado, validar con el futuro), no aleatoria.
- **C)** El modelo está subajustado y necesita más árboles / más iteraciones.
- **D)** El AUC no es una métrica válida para problemas de clasificación binaria.

---

### Pregunta 3 (5 pts) — SQL

Tienes dos tablas:

```
polizas(poliza_id, estado, prima, fecha_inicio)
siniestros(siniestro_id, poliza_id, monto)
```

Necesitas, **por estado**, el número de pólizas y el monto total de siniestros, **incluyendo los estados que no tuvieron ningún siniestro** (deben aparecer con monto 0).

**¿Cuál consulta es correcta?**

**A)**

```sql
SELECT p.estado, COUNT(*) AS n_polizas, SUM(s.monto) AS monto
FROM polizas p
JOIN siniestros s ON s.poliza_id = p.poliza_id
GROUP BY p.estado;
```

**B)**

```sql
SELECT p.estado, COUNT(DISTINCT p.poliza_id) AS n_polizas,
       COALESCE(SUM(s.monto), 0) AS monto
FROM polizas p
LEFT JOIN siniestros s ON s.poliza_id = p.poliza_id
GROUP BY p.estado;
```

**C)**

```sql
SELECT p.estado, COUNT(*) AS n_polizas, SUM(s.monto) AS monto
FROM polizas p
LEFT JOIN siniestros s ON s.poliza_id = p.poliza_id
WHERE s.monto > 0
GROUP BY p.estado;
```

**D)**

```sql
SELECT p.estado, COUNT(DISTINCT p.poliza_id) AS n_polizas, SUM(s.monto) AS monto
FROM polizas p
RIGHT JOIN siniestros s ON s.poliza_id = p.poliza_id
GROUP BY p.estado;
```

---

### Pregunta 4 (5 pts) — Diagnóstico de un modelo

Entrenas un Gradient Boosting para predecir la **severidad** (monto pagado) de un siniestro. Obtienes:

| Conjunto | RMSE |
|---|---|
| Entrenamiento | 8,200 |
| Validación | 24,500 |

**¿Cuál es el diagnóstico y la acción MÁS razonable?**

- **A)** Subajuste: aumentar la profundidad de los árboles y el número de estimadores.
- **B)** Sobreajuste: reducir la profundidad, aumentar `min_samples_leaf`, subir la regularización (L1/L2), usar *early stopping* contra el conjunto de validación y, de ser posible, más datos o menos variables ruidosas.
- **C)** Es un problema de escala: basta con estandarizar la variable objetivo y el RMSE se corrige.
- **D)** Es normal en severidad porque la distribución tiene cola larga; se puede publicar el modelo así.

---

# PARTE B — Preguntas abiertas

**2 preguntas × 15 puntos = 30 puntos.** Responde en `plantilla_respuestas.md`. **Se valora la estructura y el criterio, no la extensión**: de 5 a 8 líneas por respuesta son suficientes; puedes contestar con viñetas.

---

### Pregunta 5 (15 pts) — Diseño de la solución

Quálitas quiere un modelo que estime la **probabilidad de que una póliza de auto tenga al menos un siniestro en los próximos 12 meses**, como insumo para tarificación y para priorizar acciones de retención.

Describe cómo lo abordarías, cubriendo los cuatro puntos:

1. **Datos y variables:** qué información pedirías, qué variables construirías y **cuáles evitarías** (y por qué).
2. **Construcción y validación:** cómo armarías el conjunto de entrenamiento (definición del objetivo y de la ventana de observación) y cómo validarías el modelo.
3. **Métricas:** qué reportarías al equipo técnico y qué le reportarías al negocio.
4. **Riesgos:** menciona **dos** riesgos concretos del proyecto y cómo los mitigarías.

---

### Pregunta 6 (15 pts) — Caso: el modelo se degradó en producción

Un modelo de detección de fraude en siniestros lleva 8 meses en producción. En el último trimestre su **AUC bajó de 0.84 a 0.70** y el área de siniestros se queja de que **recibe demasiadas alertas que no son fraude** y ya casi no les hace caso.

Explica:

1. Cómo **diagnosticarías** el problema: pasos concretos, en orden, y qué datos pedirías.
2. Qué hipótesis manejarías sobre la causa (menciona al menos tres).
3. Cómo decidirías entre **reentrenar**, **recalibrar el umbral** o **rediseñar** el modelo.
4. Qué le comunicarías al área de siniestros mientras se resuelve.

---

# PARTE C — Programación

**2 ejercicios × 25 puntos = 50 puntos.** Trabaja directamente en los archivos. Cada uno trae sus propias pruebas: ejecútalo para verificar tu solución.

---

### Pregunta 7 (25 pts) — `ejercicio_1.py` · Siniestralidad por estado

Calcular frecuencia, severidad y *loss ratio* por estado a partir de dos listas de diccionarios. **Solo librería estándar de Python.**

Instrucciones completas dentro del archivo.

```bash
python3 ejercicio_1.py
```

---

### Pregunta 8 (25 pts) — `ejercicio_2.py` · Limpieza y análisis de siniestros

Limpiar `datos/siniestros_2024.csv` (montos con símbolos, dos formatos de fecha, estados mal escritos, duplicados, filas inválidas) y devolver un resumen. **Puedes usar pandas o solo librería estándar; ambas opciones valen igual.**

Instrucciones completas dentro del archivo.

```bash
python3 ejercicio_2.py
```

---

## Al terminar

1. Verifica que ambos ejercicios corran sin errores.
2. Guarda `plantilla_respuestas.md` con tus respuestas de las partes A y B.
3. Avisa al entrevistador.

**Éxito.**
