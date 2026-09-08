# Prueba tecnica — Ciencia de Datos
### Qualitas | Duracion: 60 minutos | Total: 100 puntos

**Nombre del candidato:** ______________________  **Fecha:** ____________

---

## Instrucciones

- Tienes **60 minutos**. Administra tu tiempo: los ejercicios de programacion valen la mitad del examen.
- Puedes consultar documentacion oficial (Python, pandas, SQL). **No** uses asistentes de IA.
- Escribe tus respuestas de opcion multiple y abiertas en `plantilla_respuestas.md`.
- El codigo va directo en `ejercicio_1.py` y `ejercicio_2.py`.
- Si no terminas un ejercicio, **deja comentado tu razonamiento**: se evalua tambien el enfoque.

| Parte | Contenido | Preguntas | Puntos | Tiempo sugerido |
|---|---|---|---|---|
| A | Opcion multiple | 1 – 4 | 20 | 8 min |
| B | Preguntas abiertas | 5 – 6 | 30 | 15 min |
| C | Programacion | 7 – 8 | 50 | 35 min |

---

# PARTE A — Opcion multiple (4 preguntas × 5 pts = 20 pts)

Marca **una sola** opcion por pregunta.

---

### Pregunta 1 (5 pts) — Evaluacion con clases desbalanceadas

Entrenas un modelo para detectar siniestros con posible fraude. De 200,000 siniestros historicos, 1,800 (0.9%) fueron confirmados como fraude. Tu modelo alcanza **99.1% de accuracy** en el conjunto de prueba.

**¿Cual es la interpretacion correcta?**

- **A)** El modelo es excelente y puede pasar a produccion tal cual.
- **B)** El accuracy es enganoso: un modelo que siempre prediga "no fraude" lograria ~99.1%. Hay que evaluar con precision, recall, PR-AUC y el costo economico de falsos positivos vs. falsos negativos.
- **C)** El accuracy tan alto indica sobreajuste; hay que simplificar el modelo.
- **D)** Basta con balancear las clases a 50/50 con oversampling; despues de eso el accuracy si sera confiable.

---

### Pregunta 2 (5 pts) — Validacion y fuga de informacion

Construyes un modelo para predecir si una poliza de auto tendra **al menos un siniestro en los proximos 12 meses**. Tienes datos de 2019 a 2024. Partes el dataset **aleatoriamente** 80/20 e incluyes entre las variables predictoras el campo `monto_pagado_siniestro`. Obtienes **AUC 0.97** en validacion, pero en produccion el desempeno cae a **0.61**.

**¿Cual es la causa MAS probable?**

- **A)** Falto estandarizar las variables numericas antes de entrenar.
- **B)** Hay dos problemas: (i) fuga de informacion, porque `monto_pagado_siniestro` solo se conoce *despues* de que ocurrio el siniestro y es practicamente la variable objetivo disfrazada; y (ii) la particion debio ser **temporal** (entrenar con el pasado, validar con el futuro), no aleatoria.
- **C)** El modelo esta subajustado y necesita mas arboles / mas iteraciones.
- **D)** El AUC no es una metrica valida para problemas de clasificacion binaria.

---

### Pregunta 3 (5 pts) — SQL

Tienes dos tablas:

```
polizas(poliza_id, estado, prima, fecha_inicio)
siniestros(siniestro_id, poliza_id, monto)
```

Necesitas, **por estado**, el numero de polizas y el monto total de siniestros, **incluyendo los estados que no tuvieron ningun siniestro** (deben aparecer con monto 0).

**¿Cual consulta es correcta?**

- **A)**
```sql
SELECT p.estado, COUNT(*) AS n_polizas, SUM(s.monto) AS monto
FROM polizas p
JOIN siniestros s ON s.poliza_id = p.poliza_id
GROUP BY p.estado;
```
- **B)**
```sql
SELECT p.estado, COUNT(DISTINCT p.poliza_id) AS n_polizas,
       COALESCE(SUM(s.monto), 0) AS monto
FROM polizas p
LEFT JOIN siniestros s ON s.poliza_id = p.poliza_id
GROUP BY p.estado;
```
- **C)**
```sql
SELECT p.estado, COUNT(*) AS n_polizas, SUM(s.monto) AS monto
FROM polizas p
LEFT JOIN siniestros s ON s.poliza_id = p.poliza_id
WHERE s.monto > 0
GROUP BY p.estado;
```
- **D)**
```sql
SELECT p.estado, COUNT(DISTINCT p.poliza_id) AS n_polizas, SUM(s.monto) AS monto
FROM polizas p
RIGHT JOIN siniestros s ON s.poliza_id = p.poliza_id
GROUP BY p.estado;
```

---

### Pregunta 4 (5 pts) — Diagnostico de un modelo

Entrenas un Gradient Boosting para predecir la **severidad** (monto pagado) de un siniestro. Obtienes:

| Conjunto | RMSE |
|---|---|
| Entrenamiento | 8,200 |
| Validacion | 24,500 |

**¿Cual es el diagnostico y la accion MAS razonable?**

- **A)** Subajuste: aumentar la profundidad de los arboles y el numero de estimadores.
- **B)** Sobreajuste: reducir la profundidad, aumentar `min_samples_leaf`, subir la regularizacion (L1/L2), usar *early stopping* contra el conjunto de validacion y, de ser posible, mas datos o menos variables ruidosas.
- **C)** Es un problema de escala: basta con estandarizar la variable objetivo y el RMSE se corrige.
- **D)** Es normal en severidad porque la distribucion tiene cola larga; se puede publicar el modelo asi.

---

# PARTE B — Preguntas abiertas (2 preguntas × 15 pts = 30 pts)

Responde en `plantilla_respuestas.md`. **Se valora la estructura y el criterio, no la extension**: ~15 lineas por respuesta son suficientes.

---

### Pregunta 5 (15 pts) — Diseno de la solucion

Qualitas quiere un modelo que estime la **probabilidad de que una poliza de auto tenga al menos un siniestro en los proximos 12 meses**, como insumo para tarificacion y para priorizar acciones de retencion.

Describe como lo abordarias, cubriendo los cuatro puntos:

1. **Datos y variables:** que informacion pedirias, que variables construirias y **cuales evitarias** (y por que).
2. **Construccion y validacion:** como armarias el conjunto de entrenamiento (definicion del objetivo y de la ventana de observacion) y como validarias el modelo.
3. **Metricas:** que reportarias al equipo tecnico y que le reportarias al negocio.
4. **Riesgos:** menciona **dos** riesgos concretos del proyecto y como los mitigarias.

---

### Pregunta 6 (15 pts) — Caso: el modelo se degrado en produccion

Un modelo de deteccion de fraude en siniestros lleva 8 meses en produccion. En el ultimo trimestre su **AUC bajo de 0.84 a 0.70** y el area de siniestros se queja de que **recibe demasiadas alertas que no son fraude** y ya casi no les hace caso.

Explica:

1. Como **diagnosticarias** el problema: pasos concretos, en orden, y que datos pedirias.
2. Que hipotesis manejarias sobre la causa (menciona al menos tres).
3. Como decidirias entre **reentrenar**, **recalibrar el umbral** o **rediseniar** el modelo.
4. Que le comunicarias al area de siniestros mientras se resuelve.

---

# PARTE C — Programacion (2 ejercicios × 25 pts = 50 pts)

Trabaja directamente en los archivos. Cada archivo trae sus propias pruebas: ejecutalo para verificar tu solucion.

---

### Pregunta 7 (25 pts) — `ejercicio_1.py` · Siniestralidad por estado

Calcular frecuencia, severidad y *loss ratio* por estado a partir de dos listas de diccionarios. **Solo libreria estandar de Python.**

Instrucciones completas dentro del archivo.

```bash
python3 ejercicio_1.py
```

---

### Pregunta 8 (25 pts) — `ejercicio_2.py` · Limpieza y analisis de siniestros

Limpiar `datos/siniestros_2024.csv` (montos con simbolos, dos formatos de fecha, estados mal escritos, duplicados, filas invalidas) y devolver un resumen. **Puedes usar pandas o solo libreria estandar; ambas opciones valen igual.**

Instrucciones completas dentro del archivo.

```bash
python3 ejercicio_2.py
```

---

## Al terminar

1. Verifica que ambos ejercicios corran sin errores.
2. Guarda `plantilla_respuestas.md` con tus respuestas de las partes A y B.
3. Avisa al entrevistador.

**Exito.**
