# Reservas Técnicas QES

Aplicativo de una sola página que arma la vista comparativa de reservas técnicas
(Metodología local contra Método Estatutario CNSF) a partir de dos archivos de Excel,
y va acumulando el histórico mes a mes.

Abrir `reservas/index.html` en el navegador. No necesita servidor ni instalación;
los archivos se leen en el propio navegador con SheetJS y nada se envía a ningún servidor.

## Los dos archivos de entrada

**1. Balanza de comprobación (`.xlsx`) → columna "Metodología local"**

Se localiza la fila de encabezado que empieza con `CUENTA` y se mapean las columnas de
grado (`7mo.` … `1er.`). De ahí se toma el saldo de la columna **3er grado**:

| Cuenta | Concepto                             |
|--------|--------------------------------------|
| `2205` | Reserva de Riesgos en Curso          |
| `2301` | Reserva de Siniestros Reportados     |
| `2302` | Reserva de Siniestros No Reportados  |

Los saldos vienen en negativo por ser pasivos: el aplicativo invierte el signo.
Si la cuenta no tiene saldo en 3er grado, se toma el primer grado con importe en esa fila.

El periodo sale del encabezado del reporte ("Balanza de Comprobación al 30 de Junio 2026");
si no aparece, se deduce del nombre del archivo (`Balanza_062026.xlsx` → 30-jun-2026).

**2. Archivo de actuarios (`.xlsb` o `.xlsx`) → columna "CNSF Método Estatutario"**

Se recorren todas las hojas buscando filas cuyo texto nombre uno de los tres conceptos de
reserva; de esa fila se toman los dos primeros importes a la derecha (metodología local y
método estatutario). El periodo sale de la fecha de la propia fila o del título del bloque
("30 de junio de 2026"), así que un archivo con varios cortes carga todos de una vez.

La columna local del archivo de actuarios se usa como contraste: si difiere de la balanza,
el periodo se marca en ámbar en el histórico con el detalle de la diferencia.

## La vista

- **Diferencia** = Método Estatutario − Metodología local (exceso de constitución).
- Tabla por periodo, con la columna de incremento contra el periodo anterior seleccionado.
- KPIs: diferencia total, variación del periodo y variación porcentual.
- Puente (waterfall) que descompone el movimiento de la diferencia por reserva.
- Mensajes clave redactados a partir de las cifras del propio corte.
- Unidades intercambiables: millones de USD, USD o millones de MXN (tipo de cambio editable).

## Histórico

Cada archivo cargado actualiza su periodo y conserva los demás; el histórico se guarda en
`localStorage` de ese navegador. El bloque JSON al final permite respaldarlo, moverlo a otro
equipo o restaurarlo. Al abrirlo por primera vez la vista viene precargada con el histórico
de ejemplo (jul-2025 a jun-2026), marcado como tal y sustituible por los datos propios.

## Versión en Python

`reservas/reservas_qes.py` hace la misma ingesta fuera del navegador, para un proceso
mensual por lotes. Lee `.xlsx` con openpyxl y `.xlsb` con pyxlsb, detecta solo cuál archivo
es cuál, funde el resultado en un JSON de histórico e imprime la vista.

```
pip install openpyxl pyxlsb
python reservas/reservas_qes.py Balanza_062026.xlsx ResultadosQES.xlsb --hist historico.json
```

Como módulo:

```python
from reservas_qes import Historico
h = Historico("historico.json")
h.procesar("Balanza_072026.xlsx")     # actualiza solo su periodo
h.procesar("ResultadosQES.xlsb")      # carga todos los cortes que traiga
h.guardar()
print(h.vista())                      # o h.diferencia("2026-06-30")
```
