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

## 0. Todo en un solo paso (lo normal)

```
pip install openpyxl
python proceso_completo.py
```

Eso junta los `ASEG_MM_AAAA.txt` de `C:\Spoolers` de 2020 a 2026, los cruza con
el Excel de Descargas y deja ahí mismo `resumen_series.html` y
`emision_cruzada.csv`. Sin parquet: de cada spooler se queda solo con los
registros cuya póliza o clave está en el Excel — unos cuantos cientos — así que
corre en minutos y no en horas.

Con 84 spoolers la misma póliza reaparece cada mes, así que se colapsa a **un
registro por póliza + clave + año + inciso**, el del mes más reciente. En el
HTML la celda de cada año usa el registro de **ese** año (y si ese año no está
en la emisión, el del año más cercano); la casilla *Emisión del mismo año* lo
apaga, y el detalle de la celda siempre muestra los registros de todos los años.

### Las variables que trae de la emisión

Se toman por nombre del encabezado del spooler, no adivinando. El layout real
trae `A. PATERNO`, `C.P.`, `NOMBRE (S)`, `GENTE` (el agente) y el apoderado
escrito `APODERALO LEG`, así que los patrones contemplan esas formas:

| Canónica | Sale de |
|---|---|
| `NOMBRE_PERSONA` | **armada**: `A. PATERNO` + `A. MATERNO` + `NOMBRE (S)` |
| `NOMBRECOMPLETO` | `NOMBRECOMPLETO` |
| `DIRECCION` | **armada**: `CALLE NUMERO`, `COLONIA`, `DESC ESTADO`, `C.P.` |
| `TELEFONO`, `CELULAR`, `CORREO` | `TELEFONO`, `CELULAR`, `CORREO` |
| `APODERADO_LEGAL`, `ADMINISTRADOR`, `DIRECTOR` | `APODERALO LEG`, `ADMINISTRADOR`, `DIRECTOR` |
| `RFC`, `CURP`, `TP`, `USUARIO_EMI`, `FEC_EMI`, `EMI`, `AGENTE` | las homónimas |

`NOMBRE_PERSONA` y `DIRECCION` no existen en el spooler: las arma el script,
porque son las que quieres detrás de un botón. Al correr te imprime qué
variables encontró y cuáles no.

Ajustes:

```
python proceso_completo.py --spoolers "D:\Spoolers" --anios 2020 2026
python proceso_completo.py --columnas TELEFONO CORREO DIRECCION APODERADO_LEGAL
python proceso_completo.py --todos-los-meses      # guarda cada mes por separado
```

Necesita `plantilla.html` junto al script; si no la encuentra la baja del
repositorio.

Los pasos 1 a 3 de abajo son la versión desarmada, por si quieres correr sólo
una parte.

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

**Cobertura** responde la otra pregunta: de las 296 pólizas, ¿a cuántas les
llega teléfono? ¿y correo? Una fila por variable con el conteo, el porcentaje y
una barra, ordenadas de la más completa a la más hueca. Clic en cualquier fila
y esa variable se pinta en la tabla. Se exporta a CSV, y como se calcula con el
cruce activo, cambiando de **Póliza** a **Clave** ves cuánto gana cada llave.

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

## Lista de personas bloqueadas

La columna **Clasificación** del análisis marca al asegurado: `LPB`,
`Lista Interna` o `No`. Cuenta como bloqueo cualquier valor que no sea `No`,
por descarte, para que una etiqueta nueva no pase de largo.

En la tabla, una serie cuyo asegurado está en lista lleva **franja roja al pie**
de cada celda afectada, un **borde rojo y el chip "desde aquí"** en el primer año,
y bajo la llave dice *en lista desde 2024 (LPB)*. La casilla **Solo bloqueados**
deja únicamente esas filas. El botón **Bloqueos** abre el resumen por asegurado.

### Lo que ese año significa (y lo que no)

En tu archivo, **la Clasificación no tiene historia**: los 20 asegurados traen la
misma etiqueta en todas sus filas, de 2005 a 2026. Es el estatus de *hoy*
estampado hacia atrás, no la fecha en que se dio de alta a alguien en la lista.

Por eso el panel no dice "bloqueado desde": muestra estatus, primera y última
póliza. Lo que **sí** es real es el cambio de manos: *en lista desde 2024* en una
serie significa que esa serie pasó ese año a un asegurado que hoy está en lista,
estando antes con otro que no lo está. En tu archivo son **19 series** y es lo que
vale la pena mirar. Si consigues la fecha real de alta en lista como columna del
SISE, se usa esa en lugar del año inferido.

---

## 4. Lo que se puede hacer en la tabla

- **Pivote**: serie, póliza (18 o 12), clave de asegurado, o serie + clave.
- **Dato en celda**: los botones rápidos (Asegurado, Clasificación, Póliza,
  Nombre, Dirección, Teléfono, Celular, Correo, Apoderado) o el desplegable con
  las 14 variables del análisis más las 23 de la emisión.
- **Años**: el rango arranca en el año más antiguo que aún cubre el 90 % de las
  pólizas (2018 con tu archivo actual), para no abrir con siete columnas vacías;
  se amplía con los selectores.
- **Ancho**: *Ajustar* reparte el ancho de la pantalla entre los años que haya —
  los 22 años de 2005 a 2026 caben sin scroll horizontal, con el valor completo
  en el tooltip y en el detalle. *Compacto* y *Completo* fijan el ancho si
  prefieres leer y desplazarte.
- **Resaltar cambios**: pinta la celda cuando el valor cambió respecto al año
  anterior de esa misma fila — que es justo cómo se ve el movimiento.
- **Solo con match**, buscador, y **Exportar CSV** de lo que estés viendo.
- Clic en cualquier celda: el detalle de esa póliza, con el registro completo de
  emisión al lado.

---

## Archivos

```
proceso_completo.py    spoolers + Excel -> HTML, todo de una
construir_resumen.py   genera el HTML solo a partir del Excel
preparar_emision.py    filtra el parquet o los spoolers a un CSV chico
plantilla.html         la interfaz; los scripts le inyectan los datos
```

Los datos no se versionan: el `.gitignore` excluye los `.xlsx`, los CSV de
emisión y los HTML generados, porque llevan nombres, direcciones y teléfonos de
asegurados.
