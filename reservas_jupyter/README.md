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
4. **Conectar** al servidor de auditoría, **Crear tablas** (solo la primera vez)
   y **Subir al servidor** para dejar la copia del mes.
5. Elige las tres tablas de la vista en los desplegables: *Diciembre · t-1 · t*,
   y el tipo de cambio de cierre de cada una. Si quieres más columnas, marca los
   cortes adicionales de abajo.
6. Pica **Procesar**. Se escribe `vista_reservas_AAAA-MM-DD.html` junto al
   notebook y se abre en el navegador. **Ver evolución** escribe
   `evolucion_diferencia.html` con la diferencia mes con mes.

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
como subcadena, y por eso el emparejador descarta explícitamente ese caso. La
fila «Total Reservas» se ignora sola: dice «reservas» pero no nombra ninguno de
los tres conceptos.

Los importes se leen vengan como vengan: número de Excel, o texto en cualquiera
de las dos convenciones —`$8,331,317.86` y `$8.331.317,86` son el mismo importe—.
Si aparecen los dos separadores, el último es el decimal; si uno se repite, ése es
el de miles. Importa: leer `$12.701,88` como `12.70` no truena, sólo mete un
importe mil veces más chico en el reporte.

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

## El servidor de auditoría

Cada archivo que se sube queda como una **carga** —quién, cuándo, qué archivo, con
su huella sha256— y su **detalle** con los importes por corte y concepto:

```
dbo.ReservasQES_Cargas    carga_id · usuario · equipo · fecha_carga · fuente ·
                          archivo · hoja · sha256 · periodo_min/max · filas · nota
dbo.ReservasQES_Detalle   carga_id · periodo · concepto · cuenta · etiqueta ·
                          metodologia_local · metodo_estatutario
dbo.vw_ReservasQES        las dos ya unidas, con la diferencia calculada
```

La vista `vw_ReservasQES` es para consultar desde SSMS o Excel sin armar el join:

```sql
SELECT periodo, etiqueta, metodologia_local, metodo_estatutario, diferencia
FROM   dbo.vw_ReservasQES
WHERE  fuente = 'actuarios'
ORDER  BY periodo, concepto;
```

Nada se pisa: subir otra vez el mismo mes deja una copia nueva y la anterior se
conserva, así que siempre se puede volver a la que se usó en un cierre pasado. Si
el archivo ya está (misma huella), la ventana avisa antes de duplicarlo.

La conexión es la de siempre, con autenticación integrada de Windows:

```
DRIVER={ODBC Driver 17 for SQL Server};SERVER=Qauditinterna;DATABASE=Reservas_QES;Trusted_Connection=yes
```

### La base propia

Una cuenta normal no puede crear tablas en una base ajena —en `PLD_492`, por
ejemplo, `QUALITAS\usuario` no tiene ese permiso—. El botón **Crear base** crea
una base propia en el mismo servidor: se conecta sin base, con `autocommit`
(un `CREATE DATABASE` no corre dentro de una transacción), y al terminar entra
sola a la base nueva, donde ya eres dueño y las tablas se crean sin problema.

El nombre se interpola en el SQL —`CREATE DATABASE` no admite parámetro—, así que
sólo se aceptan identificadores simples: letras, números y guion bajo, empezando
por letra. Cualquier otra cosa se rechaza antes de tocar el servidor.

El orden del primer día es: **Conectar → Crear base → Crear tablas → Subir**.
De ahí en adelante, cada mes es sólo cargar y subir.

Servidor y base se editan en la propia ventana. Hace falta `pip install sqlalchemy pyodbc`.

El driver ODBC **se detecta solo**: si el equipo trae el 18 y no el 17, o sólo el
Native Client, se usa el que haya. Al abrir la ventana, la bitácora dice qué
librerías y qué drivers encontró, y al conectar avisa si la cuenta no tiene
permiso para crear tablas en esa base — todo antes de que te pelees con el
servidor a ciegas.

### Primera corrida en una máquina nueva

```
pip install openpyxl pyxlsb sqlalchemy pyodbc
pip install tkinterdnd2      # opcional, para arrastrar y soltar
```

Luego abre el notebook, ejecuta la celda y mira la bitácora: ahí sale el
diagnóstico del equipo antes de tocar nada.

### Las tres tablas de la vista

Los desplegables listan cada copia subida —`mes · fuente · archivo · usuario · #carga`—
y el usuario decide cuál va en cada columna: **1 Diciembre**, **2 t-1**, **3 t**.
Debajo, las casillas de **cortes adicionales** permiten llevar a la vista cualquier
otro mes del histórico; se acomodan por fecha.
Por omisión se proponen el diciembre más reciente, el último corte y el anterior.
La columna local sale de la carga elegida (la balanza, si la hay) y la estatutaria
de la carga de actuarios más nueva de ese corte; el pie del HTML dice de qué carga
salió cada una. Sin servidor conectado, los desplegables muestran los cortes del
histórico local y todo sigue funcionando igual.

### Dólares o pesos, con un botón

Las dos vistas traen arriba un interruptor **Dólares / Pesos**. El HTML guarda
las cifras en las dos monedas y el CSS enseña la que el lector eligió, así que
sigue sin una línea de JavaScript y se puede mandar por correo tal cual.

Cada corte se convierte con **su propio** tipo de cambio de cierre —diciembre al
de diciembre, no al de hoy—, que se escribe junto a cada columna en la ventana y
viaja con el histórico. El que se deje en blanco usa el tipo de cambio general.

Por eso la variación en pesos no es la variación en dólares multiplicada: lleva
dentro el efecto cambiario. Con 17.6220 al cierre de marzo y 17.4986 al de junio,
la diferencia total sube USD 0.42 MM (+10.3 %) pero MXN 6.88 MM (+9.5 %). Los
indicadores, la matriz, la cascada y los mensajes clave usan todos ese mismo
criterio, para que ninguna parte de la hoja contradiga a otra.

### El gráfico de la vista

El puente de la imagen original abre por reserva el movimiento entre los **dos
últimos** cortes. Con tres columnas eso está bien; con más, dejaría fuera del
dibujo meses que sí salen en la tabla. Por eso el gráfico se elige solo: puente
hasta tres cortes, evolución en cuanto haya más, de modo que nunca falte abajo un
corte que el usuario puso arriba. El desplegable «Gráfico de la vista» permite
forzar uno u otro, y cuando el puente deja cortes fuera lo dice en el pie.

### Evolución de la diferencia

**Ver evolución** escribe una página aparte con la diferencia de todos los cortes en
columnas apiladas por reserva, con el desglose al pasar el cursor y la misma tabla
debajo. Los colores de serie salen de la paleta de la casa, pero elegidos entre los
pasos que separan bien para daltonismo (ΔE 25 en deuteranopia, contra los 2.8 del
vino contra el azul oscuro); aun así van con leyenda y etiqueta directa, para que
el color nunca sea la única pista.

## Comprobación

`verificar.py` ejecuta el bloque sin abrir ventana y contrasta 284 cifras contra
los valores de control del cierre de junio 2026: los seis cortes concepto por
concepto, la vista en millones, los indicadores, la estructura del HTML, el ida y
vuelta completo por la base de datos (contra un SQLite, con el mismo código que
corre en SQL Server), la página de evolución y el interruptor de moneda —celda por
celda, con el tipo de cambio de cada cierre, y comprobando que el indicador y el
mensaje clave digan la misma variación—, la vista plana del servidor y el
rechazo de nombres de base inválidos. Arma además la tablita de los actuarios tal
como la mandan —título del corte, las tres reservas y su Total— en las tres formas
en que puede venir el importe, y comprueba que el cruce trae la columna estatutaria
y que el desajuste contra la balanza se avisa, cargando en los dos órdenes.
Comprueba también que el gráfico de la vista cambia solo según cuántos cortes se
eligieron, y que los globos del cursor no salen abiertos al incrustarlo. Y vuelve
a cargar la balanza para confirmar que el histórico conserva los cortes anteriores.

```
python verificar.py ruta/Balanza_062026.xlsx ruta/ResultadosQES.xlsb
```

Los Excel no se guardan en el repositorio: son datos del cliente.

Para probar la parte del servidor sin red, la ventana acepta un SQLite:

```python
abrir_ventana(url_bd="sqlite:///pruebas.db")
```

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
- Conectar antes de que existan las tablas es lo normal la primera vez: la ventana
  lo dice con palabras y ofrece «Crear tablas», en vez de volcar el error de SQL.
- En SVG no hay `z-index`: lo único que decide qué tapa a qué es el orden. Por eso
  los globos del cursor se dibujan todos al final y cada columna enciende el suyo
  con el `~` de CSS.
