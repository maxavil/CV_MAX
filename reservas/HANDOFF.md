# Encargo: vista de reservas técnicas QES desde Jupyter

Este documento es el encargo completo. Pásalo tal cual a una sesión local de
Claude Code, junto con los tres archivos que se listan abajo.

---

## 1. Qué se necesita

Un **solo bloque de código Python, copiable a una celda de Jupyter Notebook**, que al
ejecutarse abra una ventana donde el usuario carga dos archivos de Excel, pique un botón
**Procesar**, y obtenga un **archivo HTML** con la vista comparativa de reservas técnicas
armada igual a la de la imagen de referencia, abierto automáticamente en el navegador.

Requisitos del entregable:

- Un bloque, una celda. Pegar y ejecutar debe ser suficiente: **la última línea del bloque
  llama a la función que abre la ventana**, no hay que teclear nada más.
- La ventana debe permitir cargar los dos Excel (botón para elegir archivos, y arrastrar y
  soltar si `tkinterdnd2` está instalado), mostrar qué leyó de cada uno, y tener un botón
  **Procesar** que genera el HTML.
- El HTML se guarda junto al notebook (por ejemplo `vista_reservas_2026-06-30.html`) y se
  abre solo en el navegador.
- El HTML es autocontenido: sin dependencias externas obligatorias, se puede mandar por
  correo y se ve igual.
- Nada sale del equipo del usuario: todo se procesa en local.

## 2. Los archivos que se entregan

| Archivo | Qué es |
|---|---|
| `Balanza_062026.xlsx` | Balanza de comprobación de Quálitas El Salvador, expresada en dólares. Es de donde sale la **metodología local**. Llega un archivo nuevo cada mes. |
| `ResultadosQES.xlsb` | Archivo de los actuarios. Es de donde sale el **método estatutario CNSF**. Trae varios cortes acumulados y se actualiza cada mes. |
| `vista_objetivo.png` | La imagen a replicar. El HTML debe verse así. |

Los dos Excel son los reales del cierre de junio 2026, para que puedas verificar contra las
cifras de la sección 5 sin adivinar.

## 3. Cómo se leen los archivos

### 3.1 Balanza → columna «Metodología local»

Estructura real del archivo:

- Una sola hoja, `Hoja1`.
- Fila 2: `Balanza de Comprobación al 30 de Junio 2026` — de ahí sale el periodo.
- Fila 6: los ordinales `7mo.` `6to.` `5to.` `4to.` `3er.` `2do.` `1er.` en las columnas C a I.
- Fila 7: `CUENTA` en A, `NOMBRE DE LA CUENTA` en B, y `GRADO` en C a I.
- De la fila 9 en adelante, las cuentas. Cada cuenta pone su saldo en la columna del grado
  que le toca por su nivel en el catálogo; las demás van vacías.

Los tres conceptos salen de la **columna de 3er grado** (la G):

| Cuenta | Concepto |
|---|---|
| `2205` | Reserva de Riesgos en Curso |
| `2301` | Reserva de Siniestros Reportados |
| `2302` | Reserva de Siniestros No Reportados |

Los saldos vienen **en negativo** porque son pasivos: hay que invertirles el signo.

No amarres las columnas a letras fijas. Localiza la fila cuya primera celda dice `CUENTA`,
lee de esa fila las columnas que dicen `GRADO` y toma el ordinal de la fila de arriba para
saber cuál es el 3er grado. Si una cuenta no tiene importe en 3er grado, usa el primer grado
con importe de esa misma fila.

Si el encabezado no trae la fecha, dedúcela del nombre del archivo: `Balanza_062026.xlsx`
es el cierre del 30 de junio de 2026.

### 3.2 Archivo de actuarios → columna «CNSF Método Estatutario»

Estructura real:

- Dos hojas, `Marzo 2026` y `Marzo 2026 (2)`. **Los nombres de hoja no son confiables**
  (dicen marzo pero traen hasta junio): recorre todas las hojas.
- El archivo está hecho de bloques, uno por corte. Cada bloque trae encima un título tipo
  `31 de diciembre de 2025` o `\n30 de junio de 2026\n`, luego una fila de encabezados, y
  luego tres filas, una por reserva.
- En cada fila de reserva: columna A `QES`, columna B el **serial de fecha de Excel**
  (46203 = 2026-06-30), columna C el nombre de la reserva, columna D la metodología local,
  columna E el método estatutario CNSF. La segunda hoja agrega en F la diferencia.

Tampoco amarres posiciones: busca en cada fila el texto que nombre una de las tres reservas
y toma **los dos primeros números a la derecha** (local y estatutario). El periodo sale del
serial de fecha de la propia fila; si no hay, del título del bloque que venga arriba.

Ojo con distinguir «Reserva de Siniestros Reportados» de «Reserva de Siniestros No
reportados»: la segunda contiene a la primera como subcadena.

El archivo trae **seis cortes**, y todos deben cargarse de una pasada.

### 3.3 El cálculo

- **Diferencia** = método estatutario CNSF − metodología local. Es el exceso de constitución
  del método estatutario, y sale positiva.
- **Total reservas** = suma de los tres conceptos, de cada lado.
- **Incremento** = diferencia del corte actual − diferencia del corte anterior mostrado.
- **Variación %** = incremento ÷ diferencia del corte anterior.
- Conversión a pesos con un tipo de cambio editable; el de la imagen es **17.4986 MXN/USD**.

### 3.4 El histórico mes con mes

Cada archivo cargado actualiza **su** periodo y deja intactos los demás. El histórico se
guarda en un JSON junto al notebook, de modo que el mes que entra solo requiere cargar la
balanza nueva y el archivo de actuarios actualizado. La vista muestra por omisión los tres
últimos cortes.

Cuando un periodo tiene balanza y archivo de actuarios, **la balanza manda** para la columna
local; la columna local del archivo de actuarios se usa como contraste y, si no coincide, hay
que avisarlo en la bitácora con el detalle de la diferencia.

## 4. La vista a replicar

Mira `vista_objetivo.png`. De arriba abajo:

1. **Título** «RESULTADOS RESERVAS TÉCNICAS QES», subtítulo «Comparativo de metodologías y
   evolución del diferencial», y el tipo de cambio usado.
2. **Tres indicadores** arriba a la derecha: Diferencia total del último corte (en USD MM y
   su equivalente en MXN MM), Variación del trimestre, y Variación % de la diferencia.
3. **La matriz**: columna de reservas con fondo vino; un grupo de tres columnas por cada
   corte (Metodología local · CNSF Método Estatutario · Diferencia) con encabezado azul; y
   al final la columna «Incremento respecto de <corte anterior>». Tres filas de reservas más
   la fila **TOTAL RESERVAS** en azul. Cifras en millones de USD, con dos decimales, y un
   guion donde el valor es cero.
4. **Nota al pie**: «Esta reserva se calcula una vez al año, al cierre del ejercicio, en
   atención a la normativa de El Salvador», ligada con asterisco a Siniestros No Reportados.
5. **Gráfico de cascada** de la evolución de la diferencia: barra del corte anterior, los
   incrementos por reserva, y barra del corte actual.
6. **Mensajes clave**: tres viñetas redactadas con las cifras del propio corte (cuánto se
   movieron las reservas locales y por qué, cuánto las estatutarias, y cómo pasó la
   diferencia total de un corte a otro con su porcentaje).
7. **Diferencias por reserva** del último corte, y una **nota relevante** al final.

Paleta de la imagen: vino `#5B1A44`, azul `#134E63`, azul claro `#1B6C86`, hielo `#E9F1F6`,
línea `#C9D8E1`, tinta `#16252D`.

## 5. Cifras de verificación

Estos son los valores correctos con los archivos que se entregan. Úsalos para comprobar
antes de dar por terminado. En USD:

| Corte | Riesgos en curso local | Riesgos en curso CNSF | Siniestros reportados (local y CNSF) | No reportados local | No reportados CNSF |
|---|---|---|---|---|---|
| 2025-07-31 | 8,331,317.86 | 11,142,522.97 | 2,818,658.21 | 12,701.88 | 325,011.19 |
| 2025-09-30 | 8,723,257.33 | 11,817,782.40 | 3,722,728.06 | 12,701.88 | 423,226.93 |
| 2025-11-30 | 9,132,774.22 | 12,231,493.45 | 4,246,854.22 | 12,701.88 | 599,138.64 |
| 2025-12-31 | 9,278,886.43 | 12,604,139.42 | 3,890,396.27 | 7,461.06 | 419,827.70 |
| 2026-03-31 | 9,171,744.38 | 12,718,231.37 | 4,187,990.32 | 7,461.06 | 555,998.70 |
| 2026-06-30 | 9,084,601.70 | 12,799,633.26 | 2,773,742.22 | 7,461.06 | 809,405.38 |

La balanza de junio 2026 debe dar exactamente `9,084,601.70`, `2,773,742.22` y `7,461.06`,
que coinciden al centavo con la columna local del archivo de actuarios para ese corte.

La vista, en millones de USD, con los tres últimos cortes:

| | Dic 2025 | | | Mar 2026 | | | Jun 2026 | | | Incremento |
|---|---|---|---|---|---|---|---|---|---|---|
| | local | CNSF | dif. | local | CNSF | dif. | local | CNSF | dif. | vs Mar 2026 |
| Riesgos en Curso | 9.28 | 12.60 | 3.33 | 9.17 | 12.72 | 3.55 | 9.08 | 12.80 | 3.72 | 0.17 |
| Siniestros Reportados | 3.89 | 3.89 | — | 4.19 | 4.19 | — | 2.77 | 2.77 | — | — |
| Siniestros No Reportados | 0.01 | 0.42 | 0.41 | 0.01 | 0.56 | 0.55 | 0.01 | 0.81 | 0.80 | 0.25 |
| **TOTAL RESERVAS** | **13.18** | **16.91** | **3.74** | **13.37** | **17.46** | **4.10** | **11.87** | **16.38** | **4.52** | **0.42** |

Indicadores de junio 2026: diferencia total **USD 4.52 MM** (~MXN 79.09 MM), variación del
trimestre **+USD 0.42 MM**, variación **+10.3 %**.

Nota: la imagen de referencia dice 10.2 % porque ese porcentaje se sacó de cifras ya
redondeadas; con los importes exactos da 10.3 %. Usa el exacto.

## 6. Trampas ya identificadas

Esto ya costó encontrarlo; no lo vuelvas a descubrir:

- **`.xlsb` no lo lee `openpyxl`.** Usa `pyxlsb` (o `pandas.read_excel(..., engine="pyxlsb")`).
  `pyxlsb` entrega las filas dispersas, solo las celdas con contenido: hay que rellenar los
  huecos hasta la columna real antes de indexar por posición.
- **`argparse` en Jupyter revienta.** Si el bloque incluye una interfaz de línea de comandos,
  protégela con `if "ipykernel" not in sys.modules`, o el kernel le pasa su `-f <archivo de
  conexión>` y aborta con `SystemExit: 2`.
- **La ventana no debe pedir un tamaño fijo.** Con `1320x900` en una laptop de 1366×768 la
  tabla queda debajo del borde de la pantalla y parece que «no hizo nada». Dimensiona con
  `winfo_screenwidth()` y `winfo_screenheight()` y centra.
- **Los errores dentro de callbacks de Tk se van a `stderr`**, que desde Jupyter va a la
  consola donde se lanzó el notebook: invisible. Redirígelos a la bitácora de la ventana con
  `root.report_callback_exception`.
- **`mainloop()` deja la celda ocupada** con el `[*]` hasta que se cierre la ventana. Es
  normal, pero conviene avisarlo con un `print` al abrir.
- **Rutas de Windows**: `r"C:\Users\..."`, porque `\U` es un escape de Python.
- **Fecha serial de Excel**: base 1899-12-30.
- **Si el kernel corre en un servidor, WSL o contenedor no hay escritorio** y ninguna ventana
  puede abrir; ahí Tk lanza `TclError: no display name`. Vale la pena atrapar ese caso y
  explicarlo en vez de dejar un rastro crudo.

## 7. Punto de partida

Ya existe una implementación funcionando, verificada contra estos mismos archivos, en la rama
`claude/excel-view-two-files-pz2vxt` del repositorio `maxavil/CV_MAX`:

- `reservas/index.html` — la vista completa en HTML y JavaScript, con la maqueta, la paleta,
  el gráfico de cascada en SVG y los mensajes clave. **Es la referencia visual del HTML que
  hay que generar.**
- `reservas/reservas_qes.py` — la lectura de los dos Excel en Python (`parse_balanza`,
  `parse_actuarios`), el histórico en JSON y una ventana Tk.
- `reservas/README.md` — las reglas de lectura resumidas.

Lo más rápido es partir de ahí: tomar la lectura de `reservas_qes.py`, tomar la maqueta de
`reservas/index.html`, y unirlas en el bloque para Jupyter, de modo que el botón Procesar
escriba ese HTML con los datos ya insertados en vez de pintar la tabla en Tk.

## 8. Cómo comprobar antes de entregar

1. Correr el bloque en un notebook limpio: la ventana abre sin teclear nada más.
2. Cargar los dos Excel y picar Procesar.
3. El HTML se abre en el navegador y coincide con `vista_objetivo.png` en estructura, colores
   y cifras.
4. Todos los números cuadran con la sección 5.
5. Volver a correrlo cargando solo la balanza de otro mes: el histórico conserva los cortes
   anteriores y la vista se recorre al nuevo.
