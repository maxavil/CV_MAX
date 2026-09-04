# Reservas técnicas QES · bloque para Jupyter

Un solo bloque de Python, copiable a **una** celda de Jupyter. Al ejecutarlo abre
una ventana donde se cargan los dos Excel del mes; el botón **Procesar** escribe
la vista comparativa de reservas técnicas en un HTML autocontenido junto al
notebook y lo abre en el navegador.

Nada sale del equipo: todo se lee y se arma en local.

```
reservas_jupyter/
├── celda_jupyter.py     ← el bloque. Es el entregable: se copia entero a una celda
├── Reservas_QES.ipynb   ← el mismo bloque, ya puesto en un notebook
├── verificar.py         ← comprueba la lectura contra las cifras de control
└── README.md
```

## Cómo se usa

1. `pip install openpyxl pyxlsb` (y `pip install tkinterdnd2` si quieres arrastrar
   y soltar; sin esa librería la ventana funciona igual con el botón).
2. Abre `Reservas_QES.ipynb`, o pega el contenido de `celda_jupyter.py` en una
   celda vacía. La última línea ya llama a `abrir_ventana()`: no hay que teclear
   nada más.
3. Carga la balanza (`.xlsx`) y el archivo de actuarios (`.xlsb`). La ventana
   muestra qué leyó de cada uno antes de procesar nada.
4. Pica **Procesar**. Se escribe `vista_reservas_AAAA-MM-DD.html` junto al
   notebook y se abre en el navegador.

La celda queda ocupada con `[*]` mientras la ventana está abierta: es el
`mainloop()` de Tk y es normal. Se libera al cerrarla.

## Cómo se leen los archivos

**Balanza de comprobación (`.xlsx`) → columna «Metodología local».**
No se amarran columnas a letras fijas: se localiza la fila cuya primera celda dice
`CUENTA`, se leen de ahí las columnas que dicen `GRADO` y el ordinal de la fila de
arriba (`7mo. … 1er.`) dice cuál es el 3er grado. De esa columna salen:

| Cuenta | Concepto                            |
|--------|-------------------------------------|
| `2205` | Reserva de Riesgos en Curso         |
| `2301` | Reserva de Siniestros Reportados    |
| `2302` | Reserva de Siniestros No Reportados |

Los saldos vienen en negativo por ser pasivos: se les invierte el signo. Si una
cuenta no tiene importe en 3er grado, se usa el primer grado con importe de esa
misma fila. El periodo sale del encabezado del reporte («Balanza de Comprobación
al 30 de Junio 2026»); si no lo trae, del nombre del archivo (`Balanza_062026.xlsx`).

**Archivo de actuarios (`.xlsb`) → columna «CNSF Método Estatutario».**
Los nombres de hoja no son de fiar (dicen «Marzo 2026» y traen hasta junio), así
que se recorren todas. En cada fila se busca el texto que nombre una de las tres
reservas y se toman los dos primeros importes a su derecha: metodología local y
método estatutario. El periodo sale del serial de fecha de la propia fila (base
1899-12-30) y, si no hay, del título del bloque de arriba. Los seis cortes del
archivo se cargan de una pasada.

Cuidado con «Reserva de Siniestros **No** Reportados»: contiene a «Reportados»
como subcadena, y por eso el emparejador descarta explícitamente ese caso.

**El cálculo.** Diferencia = método estatutario CNSF − metodología local (el exceso
de constitución del estatutario, sale positiva). Total = suma de los tres conceptos
de cada lado. Incremento = diferencia del corte actual − la del corte anterior
mostrado. Variación % = incremento ÷ diferencia anterior. La conversión a pesos usa
el tipo de cambio editable de la ventana (por omisión 17.4986 MXN/USD).

**El histórico.** Cada archivo actualiza su periodo y deja intactos los demás; vive
en `historico_reservas.json`, junto al notebook. La vista muestra los tres últimos
cortes. Cuando un periodo tiene balanza y archivo de actuarios, la balanza manda
para la columna local; la del archivo de actuarios sirve de contraste y, si no
coincide, se avisa en la bitácora y en una banda ámbar del HTML.

Un corte al que solo se le cargó una de las dos fuentes no inventa ceros: sus
celdas de la otra columna quedan en guion y los indicadores dicen qué falta.

## Comprobación

`verificar.py` ejecuta el bloque sin abrir ventana y contrasta 106 cifras contra
los valores de control del cierre de junio 2026 —los seis cortes concepto por
concepto, la vista en millones, los indicadores y la estructura del HTML—, y
además vuelve a cargar la balanza para confirmar que el histórico se conserva.

```
python verificar.py ruta/Balanza_062026.xlsx ruta/ResultadosQES.xlsb
```

Los dos Excel no se guardan en el repositorio: son datos del cliente.

Nota sobre dos cifras del encargo. La variación de junio 2026 es **+10.3 %**
(la imagen de referencia dice 10.2 % porque se sacó de cifras ya redondeadas) y,
por lo mismo, el equivalente en pesos de la diferencia total es **~MXN 79.04 MM**
(79.09 sale de multiplicar el 4.52 ya redondeado). El bloque usa los importes
exactos en los dos casos.

## Trampas ya resueltas

- `.xlsb` no lo lee openpyxl: se usa `pyxlsb`, que además entrega las filas
  dispersas —solo las celdas con contenido—, así que se rellenan los huecos hasta
  la columna real antes de indexar por posición.
- Nada de `argparse`: en Jupyter vería el `-f <archivo de conexión>` del kernel y
  abortaría con `SystemExit: 2`.
- La ventana no pide un tamaño fijo: se dimensiona con `winfo_screenwidth()` /
  `winfo_screenheight()` y se centra, para que en una laptop de 1366×768 no quede
  por debajo del borde de la pantalla.
- Los errores dentro de los callbacks de Tk se van a `stderr`, que desde Jupyter
  acaba en la consola del servidor: invisible. Se redirigen a la bitácora de la
  ventana con `root.report_callback_exception`.
- Si el kernel corre en un servidor, en WSL o en un contenedor no hay escritorio y
  Tk lanza `TclError: no display name`. El bloque lo atrapa y explica las
  alternativas en vez de dejar el rastro crudo; para ese caso está
  `procesar_sin_ventana(balanza, actuarios)`, que hace lo mismo sin interfaz.
- Rutas de Windows: escríbelas como `r"C:\Users\..."`, porque `\U` es un escape
  de Python.
