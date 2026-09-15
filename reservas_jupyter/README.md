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
5. Elige las tres tablas de la vista en los desplegables: *Diciembre · t-1 · t*.
   El tipo de cambio de cierre lo trae solo de Banco de México; si prefieres,
   tecléalo. Si quieres más columnas, marca los cortes adicionales de abajo.
6. Pica **Procesar**. Se escribe `vista_reservas_AAAA-MM-DD.html` junto al
   notebook y se abre en el navegador. **Ver evolución** escribe
   `evolucion_diferencia.html` con la diferencia mes con mes, y **Vista dirección**
   escribe `brecha_direccion.html`, la hoja de una página para el comité.

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
un solo tipo de cambio para toda la hoja —el ancla, ver abajo—; el de la ventana
(por omisión 17.4986 MXN/USD) es el que se aplica a los cortes sin uno propio.

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

### Letra grande

La hoja está pensada para leerse sin esfuerzo: la letra más chica es de 13 px y
la de la matriz, de 16. Arriba hay un par de botones **− / +**: cada toque mueve
un 10 % TODO el tablero de golpe —letras y números, matriz, indicadores, cascada
y mensajes— entre el 80 % y el 200 %; arranca en 115 %. Las etiquetas de la
cascada crecieron igual, y por eso el lienzo del gráfico creció con ellas.

### El lila de la casa

Los encabezados de las metodologías van en lila (`#5B3A8C` y `#7355A8`), no en
azul. El vino de la casa (`#5B1A44`) se queda con la columna de reservas, la fila
de totales, la diferencia y la base de la cascada; los movimientos de la cascada
van en el lila claro y lo que baja, en rojo. Sobre los dos tonos oscuros el texto
blanco tiene contraste de sobra (8.6:1 y 5.8:1); sobre el claro nunca va texto
blanco, sólo relleno.

### Los meses los pone y los quita el lector

El HTML lleva dentro **todos los cortes master**, no sólo los tres que se ven al
abrir. Arriba de la matriz hay una casilla por mes: al marcarla o desmarcarla la
hoja se rehace sola —matriz, indicadores, cascada, mensajes clave y diferencias
por reserva— sin volver a Jupyter y sin conexión. Los textos que dependen de un
par de cortes vienen escritos desde Python, así que la redacción sigue viviendo
en un solo lugar.

Con JavaScript apagado la hoja no se puede rehacer, pero no queda en blanco: un
bloque `<noscript>` trae la matriz y la cascada de los cortes con los que se
generó.

### Versiones master

De un mismo corte puede haber varias copias subidas. La **master** es la que
vale: la verificada, la que se usó para el cierre.

La ventana trae un panel con todas las master a la vista —corte, etiqueta, de
qué carga sale cada columna, quién la marcó y cuándo— y cuatro acciones:

| Acción | Qué hace |
|---|---|
| **Marcar los de la vista** | Convierte en master los cortes que están en la vista, con la etiqueta escrita arriba |
| **Cambiar…** | Elige a mano de qué carga sale la metodología local y de cuál la estatutaria, y reetiqueta |
| **Quitar del master** | Saca ese corte. Las cargas no se tocan: sólo deja de ser el bueno |
| **Borrar carga del servidor…** | Deshace una subida equivocada. Se niega si esa carga está sirviendo de master |

Sólo los cortes master viajan dentro del HTML, y sólo con ellos se arma la
vista: si eliges un corte que no es master, se queda fuera y la bitácora lo
dice. Así, dentro del archivo el lector únicamente puede agregar tablas ya
verificadas.

La estructura:

```
dbo.ReservasQES_Master   periodo · carga_local · carga_cnsf · etiqueta ·
                         usuario · fecha · nota
```

Cambiar la master de un mes no toca las cargas: la copia anterior sigue ahí y se
puede volver a ella. Los cortes master son los que viajan dentro del HTML para
que el lector los meta y los saque.

### El tipo de cambio lo trae solo de Banxico

El botón **Traer TC de Banxico** pide al SIE de Banco de México el tipo de cambio
de cierre de cada corte que no lo tenga, y lo guarda en el histórico. Al picar
**Procesar** o **Ver evolución** se hace solo, sin botón de por medio: la idea es
que a fin de mes ya esté ahí y no haya que batallar.

La serie es **SF43718**, que se llama con todas sus letras *«Tipo de cambio Pesos
mexicanos por Dólar E.U.A. para solventar obligaciones en moneda extranjera (fecha
de determinación — Fix)»*: exactamente la que pide el cierre.

**El token.** El SIE es gratis pero pide un token, que se saca en un minuto en
`banxico.org.mx/SieAPIRest/service/v1/token`. Se pega en la casilla **Token SIE**
de la ventana y queda guardado en `banxico_token.txt`, junto al notebook, para no
volver a teclearlo; también se lee de la variable de entorno `BANXICO_TOKEN`. Es
una credencial de sólo lectura de series públicas, pero va en texto plano: el
`.gitignore` ya la excluye.

**Lo que hace bien.** El cierre cae en sábado, domingo o feriado con mucha
frecuencia —el 31 de diciembre, sin ir más lejos— y esos días la serie no trae
dato. Por eso no se pide un día suelto sino una ventana de dos semanas que termina
en el corte, y se toma el último dato que haya: el del último día hábil del mes.
Lo que ya tenga tipo de cambio no se toca, así que lo tecleado a mano siempre
manda. Y si la red está caída se rinde en el primer corte en vez de esperar seis
veces seguidas y dejar la ventana congelada un minuto.

**Lo que pasa si no hay red.** Nada se rompe. Sin token, sin salida a internet o
con el sitio bloqueado por la red de la empresa, la ventana lo dice una vez y todo
sigue funcionando igual que antes: el tipo de cambio se teclea a mano. Esta es la
única parte del bloque que toca la red, y lo único que viaja es una fecha y el
número de serie — los importes nunca salen del equipo.

### Dólares o pesos, con un botón

Las dos vistas traen arriba un interruptor **Dólares / Pesos**. El HTML guarda
las cifras en las dos monedas y el CSS enseña la que el lector eligió, así que
sigue sin una línea de JavaScript y se puede mandar por correo tal cual.

**El ancla del tipo de cambio.** Cada corte viaja con el tipo de cambio de Banco
de México a SU cierre —el de obligaciones a esa fecha—, que el app trae solo (ver
abajo) o que se teclea junto a cada columna en la ventana; el que quede en blanco
y no se haya podido traer usa el tipo de cambio general. Pero los pesos que se ven salen de **uno solo** de
ellos: el ancla, que el HTML trae en un desplegable arriba y que por omisión es
la del último corte de la vista.

Cambiar el ancla vuelve a convertir el tablero entero de golpe: matriz,
indicadores, cascada, mensajes clave y diferencias por reserva. Así los importes
en pesos de todos los meses son comparables entre sí y lo que se lee moviéndose es
la reserva, no el dólar: la variación en pesos es la variación en dólares por el
ancla, y el porcentaje sale idéntico en las dos monedas (+10.3 % entre marzo y
junio, sean USD 0.42 MM o MXN 7.38 MM al ancla de junio). Ninguna parte de la
hoja puede contradecir a otra porque todas leen el mismo número.

### El gráfico de la vista

La cascada **encadena todos los cortes elegidos**: una barra por cada corte con su
diferencia total y, entre cada dos, el movimiento abierto por reserva. Con dos
cortes es el puente de siempre; con cuatro son cuatro cascadas unidas, y se ve
caminar la diferencia corte a corte. Agregar o quitar cortes arriba cambia el
dibujo: el gráfico nunca deja fuera un mes que está en la tabla.

El lienzo crece con los cortes para que las etiquetas no se encimen, y lo que baja
va en rojo. Cuando el gráfico ya no cabe de un vistazo, se desplaza de lado dentro
de su panel y lo dice al pie; no se comprime ni se corta. El desplegable «Gráfico de la vista» ofrece además columnas por corte,
para cuando sólo interesa el nivel y no el movimiento.

### Vista dirección · la hoja de una página

**Vista dirección** escribe `brecha_direccion.html`: no es el tablero con otro
color, es otro documento con otra pregunta. El tablero sirve para revisar cifras;
esta hoja sirve para decidir si hay que reconocer algo en los estados financieros.
Por eso no se arma ni se desarma: se imprime, se proyecta y se manda por correo, y
no cambia entre quien la abre y quien la recibe.

Contesta cuatro preguntas en el orden en que las hace quien firma los estados:

1. **¿De cuánto es la brecha?** La cifra sola, arriba, a 76 px. Una sola en toda la
   página: si hay tres cifras heroicas no hay ninguna.
2. **¿Sobre qué base?** Un gráfico de mancuernas, un renglón por corte: el punto
   claro es lo registrado con la metodología local, el oscuro lo que resultaría del
   método estatutario, y **el trazo que los une es la brecha**. Se eligió por encima
   de dos barras juntas o de una apilada porque lo que importa no es cuánto mide
   cada metodología, sino la distancia entre las dos.
3. **¿Va creciendo?** La trayectoria de la brecha, corte a corte, con la escala
   desde cero: cortar el eje exageraría la brecha, que es justo de lo que nos van a
   acusar. Sólo los extremos llevan cifra; un número sobre cada punto no se lee.
4. **¿De dónde sale?** La apertura por reserva, de mayor a menor, y debajo de cada
   nombre cuánto supera el estatutario a los libros — en por ciento cuando es poco
   (`+41% sobre libros`) y en veces cuando es mucho (`×108 sobre libros`), que es
   como se entiende.

Cierra con la nota relevante y **la tabla completa**: todos los cortes, las tres
reservas, las dos metodologías, la brecha y la proporción. Ningún dato de la página
vive sólo dentro de un gráfico.

Los colores pasaron el validador de daltonismo: lila contra vino separan ΔE 19 en
deuteranopia y 20 a vista normal, por encima del piso. Aun así cada punto va con
leyenda, para que el color nunca sea la única pista. El interruptor de moneda es el
mismo de las otras vistas, y al imprimir la portada se vuelve blanca con tinta vino
en lugar de un bloque de tinta a sangre.

### Evolución de la diferencia

**Ver evolución** escribe una página aparte con la diferencia de todos los cortes en
columnas apiladas por reserva, con el desglose al pasar el cursor y la misma tabla
debajo. Los colores de serie salen de la paleta de la casa, pero elegidos entre los
pasos que separan bien para daltonismo: el peor par es lila contra vino, ΔE 19 en
deuteranopia y 20 a vista normal. El tinta oscuro que llevaba siniestros reportados
se quitaba mal del vino (ΔE 4.6) y se cambió por ocre. Aun así van con leyenda y
etiqueta directa, para que el color nunca sea la única pista.

## Comprobación

`verificar.py` ejecuta el bloque sin abrir ventana y contrasta 465 cifras contra
los valores de control del cierre de junio 2026: los seis cortes concepto por
concepto, la vista en millones, los indicadores, la estructura del HTML, el ida y
vuelta completo por la base de datos (contra un SQLite, con el mismo código que
corre en SQL Server), la página de evolución, el interruptor de moneda y el ancla del tipo
de cambio —celda por celda, comprobando que el indicador y el mensaje clave digan
la misma variación y que ninguna conversión se quede leyendo el cierre del corte
que pinta en vez del ancla—, la vista plana del servidor y el
rechazo de nombres de base inválidos. Para el tipo de cambio de Banxico levanta un
SIE de mentiras con la forma real de la respuesta —no llama a la red— y comprueba
lo que de veras se puede equivocar: el cierre que cae en domingo, el token malo,
la serie sin dato, la red caída, el relevo entre varios tokens y que no se pise lo
tecleado a mano. De la hoja de dirección comprueba que haya una sola cifra heroica,
que las cifras que dice sean las del histórico en las dos monedas, que la tabla de
respaldo lleve todos los cortes, que un corte a medias quede fuera y avisado, y que
sin ningún corte completo se niegue a dibujar una hoja vacía. Arma además la tablita de los actuarios tal
como la mandan —título del corte, las tres reservas y su Total— en las tres formas
en que puede venir el importe, y comprueba que el cruce trae la columna estatutaria
y que el desajuste contra la balanza se avisa, cargando en los dos órdenes.
Comprueba también que la cascada encadena todos los cortes elegidos, que las
versiones master se marcan, se cambian y se quitan sin tocar las cargas, que una
carga que sirve de master no se deja borrar y que borrar una suelta no deja
detalle huérfano, que el
HTML lleva dentro los seis cortes con sus textos precalculados, y que no queda
letra por debajo de 13 px ni en la hoja ni en los gráficos. Y vuelve
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
