# Resumen de Series — movimiento anual de pólizas

Herramienta para ver, en una sola tabla, cómo se movió cada serie (VIN) año con
año: a la izquierda el pivote, a la derecha una columna por año de emisión, y
botones para cambiar qué dato se pinta en cada celda — incluyendo el nombre, la
dirección, el teléfono o el apoderado legal que vienen en el spooler de emisión.

Son dos bases:

| Base | De dónde sale | Para qué sirve aquí |
|---|---|---|
| **Análisis de Series** (`.xlsx`, SISE) | tu carpeta de Descargas | arma el pivote: serie, póliza, clave de asegurado, fecha de emisión |
| **Emisión** (`ASEG_MM_AAAA.txt` o el parquet de `consolidar_aseg.py`) | `C:\Spoolers` | aporta las variables que se pintan en las celdas |

---

## 1. Generar el HTML

```
pip install openpyxl
python construir_resumen.py
```

Sin argumentos busca `*Analisis*Series*.xlsx` en tu carpeta de Descargas y deja
ahí mismo `resumen_series.html`. Si prefieres decírselo:

```
python construir_resumen.py --excel "C:\Users\max\Downloads\Analisis_de_Series.xlsx" ^
                            --salida "C:\Users\max\Downloads\resumen_series.html"
```

El HTML queda autocontenido: los datos del Excel van incrustados, no necesita
servidor ni internet. Ábrelo con doble clic.

## 2. Cargar la emisión

Dentro del HTML, botón **Cargar emisión**. Acepta dos cosas:

- **El spooler tal cual** (`ASEG_07_2024.txt`, separado por `|`). Se lee por
  streaming en el navegador y **solo se guardan en memoria los registros cuya
  póliza o clave aparece en tu análisis**, así que un archivo de varios GB no
  tumba la pestaña. Reconstruye los registros partidos y descarta los que traen
  un `|` dentro del dato, con el mismo criterio que `consolidar_aseg.py`; los
  conteos salen en *Ver matches*.
- **Un CSV** — por ejemplo el que produce `preparar_emision.py`.

El archivo no sale de tu equipo: todo ocurre en el navegador.

Al soltarlo aparece un paso de configuración donde eliges:

- **la columna de póliza y la de clave** (se adivinan solas: se puntúa cada
  columna por cuántos de sus valores caen dentro de las llaves del análisis, y
  te dice cuántas coincidieron en la muestra antes de procesar el archivo);
- **qué variables traer** — vienen preseleccionadas nombre, dirección, teléfono
  y apoderado; con *Todas* traes las 67.

### El parquet

El navegador no lee parquet. Para esa ruta:

```
pip install pyarrow
python preparar_emision.py --parquet "C:\Spoolers\consolidado\aseg_consolidado.parquet"
```

Filtra el parquet a lo que cruza con el Excel y escribe
`emision_para_html.csv` (unos cientos de filas) en Descargas. Ese CSV lo
arrastras al HTML, o lo dejas ya incrustado:

```
python construir_resumen.py --emision "C:\Users\max\Downloads\emision_para_html.csv"
```

También trabaja directo sobre los `.txt` sin consolidar:

```
python preparar_emision.py --spoolers "C:\Spoolers"
python preparar_emision.py --columnas NOMBRE_ASEGURADO DIRECCION TELEFONO APODERADO_LEGAL
```

---

## 3. Cómo se cruzan las bases

El análisis trae la póliza con **18 dígitos** (`049510003960000000`) y el spooler
abre cada registro con **12** (el `^\d{12}\|` de `consolidar_aseg.py`), así que
la llave son **los primeros 12**. La clave de asegurado se normaliza a 10
dígitos con ceros a la izquierda. Los cuatro modos del botón **Cruce**:

| Modo | Qué hace |
|---|---|
| **Póliza** | solo por los 12 dígitos |
| **Clave** | solo por la clave de asegurado |
| **Ambas** | el registro de emisión debe coincidir en póliza **y** en clave |
| **Cualquiera** | intenta ambas, luego póliza, luego clave |

**Ver matches** abre el diagnóstico: cuántas pólizas cruzaron por cada llave,
cuántas por las dos, cuántas por ninguna, y el detalle póliza por póliza con lo
que se encontró del otro lado. Se exporta a CSV.

Un punto de color en cada celda indica por dónde entró el match
(azul = póliza, verde = clave, morado = ambas, rojo = sin match).

### Dos cosas que conviene tener presentes

- **Colisión de los 12 dígitos.** En tu propio archivo hay casos como
  `040800115143326775` y `040800115143000000`: distintos en 18 dígitos, iguales
  en 12. Ambos apuntan al mismo registro de emisión. Es inevitable si el spooler
  solo guarda 12; el panel de matches lo deja ver.
- **Varias pólizas por serie y año.** Hay 60 combinaciones serie+año con más de
  un registro (la serie `1FUJGLBG7DSFA0881` tiene 4 en 2022). La celda muestra
  el primer valor y una insignia `+n`; haz clic para ver todos.

---

## 4. Lo que se puede hacer en la tabla

- **Pivote**: serie, póliza (18 o 12), clave de asegurado, o serie + clave.
- **Dato en celda**: los botones rápidos, o el desplegable con las 14 variables
  del análisis más todas las que hayas traído de la emisión.
- **Años**: el rango arranca en el año más antiguo que aún cubre el 90 % de las
  pólizas (2018 con tu archivo actual), para no abrir con siete columnas vacías;
  se amplía con los selectores.
- **Resaltar cambios**: pinta la celda cuando el valor cambió respecto al año
  anterior de esa misma fila — que es justo cómo se ve el movimiento.
- **Solo con match**, buscador, y **Exportar CSV** de lo que estés viendo.
- Clic en cualquier celda: el detalle de esa póliza, con el registro completo de
  emisión al lado.

---

## Archivos

```
construir_resumen.py   genera el HTML a partir del Excel
preparar_emision.py    filtra el parquet o los spoolers a un CSV chico
plantilla.html         la interfaz; construir_resumen.py le inyecta los datos
```

Los datos no se versionan: el `.gitignore` excluye los `.xlsx`, los CSV de
emisión y los HTML generados, porque llevan nombres, direcciones y teléfonos de
asegurados.
