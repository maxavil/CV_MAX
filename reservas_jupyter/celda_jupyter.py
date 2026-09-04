# =============================================================================
#  RESERVAS TÉCNICAS QES · vista comparativa en HTML
#  ---------------------------------------------------------------------------
#  Pega TODO este bloque en UNA celda de Jupyter y ejecútala.
#
#  Se abre una ventana: cargas la balanza de comprobación (.xlsx) y el archivo
#  de actuarios (.xlsb), picas «Procesar» y se escribe un HTML autocontenido
#  junto al notebook, que se abre solo en el navegador.
#
#  Nada sale del equipo: todo se lee y se arma en local.
#  Requisitos: openpyxl (para .xlsx) y pyxlsb (para .xlsb).
#      pip install openpyxl pyxlsb
#  Opcional, para arrastrar y soltar:  pip install tkinterdnd2
# =============================================================================

from __future__ import annotations

import datetime as dt
import html
import json
import os
import re
import sys
import unicodedata
import webbrowser
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Sequence

# -----------------------------------------------------------------------------
# 1. Conceptos, paleta y parámetros
# -----------------------------------------------------------------------------


@dataclass(frozen=True)
class Concepto:
    id: str
    label: str
    cuenta: str
    patron: "re.Pattern"
    nota: bool = False


CONCEPTOS: list[Concepto] = [
    Concepto("rrc",  "Reserva de Riesgos en Curso",         "2205", re.compile(r"riesgos en curso")),
    Concepto("rsr",  "Reserva de Siniestros Reportados",    "2301", re.compile(r"siniestros reportados")),
    Concepto("rsnr", "Reserva de Siniestros No Reportados", "2302", re.compile(r"no reportados"), nota=True),
]

GRADO = 3                    # columna de la balanza de la que se toma el saldo
FX_DEFAULT = 17.4986         # MXN por USD
PERIODOS_EN_VISTA = 3        # cortes que se muestran por omisión
HIST_JSON = "historico_reservas.json"

NOTA_PIE = ("Esta reserva se calcula una vez al año, al cierre del ejercicio, "
            "en atención a la normativa de El Salvador.")
NOTA_RELEVANTE = ("La diferencia por metodologías de El Salvador aún no ha sido "
                  "reconocida en los EEFF de QC.")

# paleta de la vista objetivo
PLUM, PLUM_DEEP, PLUM_SOFT = "#5B1A44", "#451234", "#7A2A5E"
TEAL, TEAL_2, TEAL_SOFT = "#134E63", "#1B6C86", "#3E8AA6"
ICE, ICE_2, PAPER, GROUND = "#E9F1F6", "#F5F9FB", "#FFFFFF", "#EFF3F6"
INK, INK_2, INK_3 = "#16252D", "#4E6069", "#7B8C95"
LINE, LINE_SOFT = "#C9D8E1", "#E1EAF0"
NEG, WARN = "#B04234", "#9C6A0E"

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
         "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

# -----------------------------------------------------------------------------
# 2. Utilidades
# -----------------------------------------------------------------------------


def norm(v: Any) -> str:
    """Minúsculas, sin acentos y con espacios colapsados."""
    s = "" if v is None else str(v)
    s = unicodedata.normalize("NFD", s.lower())
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s).strip()


def to_num(v: Any) -> "float | None":
    """Número de una celda; acepta texto con comas, signo $ y paréntesis."""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if not isinstance(v, str):
        return None
    s = re.sub(r"[\s$]", "", v).replace(",", "")
    neg = bool(re.fullmatch(r"\(.*\)", s))
    if neg:
        s = s[1:-1]
    if not re.fullmatch(r"-?\d*\.?\d+", s):
        return None
    n = float(s)
    return -n if neg else n


def ultimo_dia(anio: int, mes: int) -> int:
    siguiente = dt.date(anio + (mes == 12), (mes % 12) + 1, 1)
    return (siguiente - dt.timedelta(days=1)).day


def serial_a_fecha(n: float) -> "str | None":
    """Serial de Excel -> 'AAAA-MM-DD'. La base es 1899-12-30."""
    if not isinstance(n, (int, float)) or isinstance(n, bool):
        return None
    if n < 20000 or n > 80000:
        return None
    return (dt.date(1899, 12, 30) + dt.timedelta(days=int(round(n)))).isoformat()


def celda_a_fecha(v: Any) -> "str | None":
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return serial_a_fecha(v)
    return None


def parse_fecha(texto: Any) -> "str | None":
    """'al 30 de Junio 2026' / '30 de junio de 2026' -> '2026-06-30'."""
    m = re.search(r"(\d{1,2})\s+de\s+([a-z]+)\s*(?:de\s*)?(\d{4})", norm(texto))
    if not m or m.group(2) not in MESES:
        return None
    return dt.date(int(m.group(3)), MESES.index(m.group(2)) + 1, int(m.group(1))).isoformat()


def periodo_de_nombre(nombre: str) -> "str | None":
    """'Balanza_062026.xlsx' -> '2026-06-30'."""
    m = re.search(r"(0[1-9]|1[0-2])[_\-.]?(20\d{2})", str(nombre))
    if not m:
        return None
    mes, anio = int(m.group(1)), int(m.group(2))
    return dt.date(anio, mes, ultimo_dia(anio, mes)).isoformat()


def etiqueta_periodo(k: str) -> str:
    a, m, d = k.split("-")
    return f"{int(d)} de {MESES[int(m) - 1]} de {a}"


def etiqueta_corta(k: str) -> str:
    a, m, _ = k.split("-")
    return f"{MESES[int(m) - 1].capitalize()} {a}"


def match_concepto(texto: Any) -> "Concepto | None":
    """Ojo: «no reportados» contiene a «reportados» como subcadena."""
    s = norm(texto)
    if "reserva" not in s:
        return None
    for c in CONCEPTOS:
        if c.id == "rsr" and "no reportados" in s:
            continue
        if c.patron.search(s):
            return c
    return None


def fmt(v: float) -> str:
    return f"{v:,.2f}"


def mm(v: float) -> str:
    """Millones de USD con dos decimales."""
    return f"{v / 1e6:,.2f}"


def celda_mm(v: "float | None") -> str:
    """Millones con dos decimales; guion cuando el valor es cero o no hay dato."""
    if v is None:
        return "—"
    e = v / 1e6
    return "—" if abs(e) < 0.005 else f"{e:,.2f}"


def esc(t: Any) -> str:
    return html.escape("" if t is None else str(t), quote=True)


# -----------------------------------------------------------------------------
# 3. Apertura del libro (.xlsx y .xlsb dan el mismo resultado)
# -----------------------------------------------------------------------------

Hoja = "tuple[str, list[list[Any]]]"


def leer_libro(ruta: "str | Path") -> "list[tuple[str, list[list[Any]]]]":
    """Devuelve [(nombre de hoja, filas densas)]. openpyxl no abre .xlsb: eso es pyxlsb."""
    ruta = Path(ruta)
    ext = ruta.suffix.lower()

    if ext == ".xlsb":
        try:
            from pyxlsb import open_workbook
        except ImportError as err:
            raise RuntimeError(
                "Para leer .xlsb hace falta pyxlsb. En una celda: !pip install pyxlsb"
            ) from err
        hojas: "list[tuple[str, list[list[Any]]]]" = []
        with open_workbook(str(ruta)) as wb:
            for nombre in wb.sheets:
                with wb.get_sheet(nombre) as sh:
                    filas: "list[list[Any]]" = []
                    for fila in sh.rows():
                        # pyxlsb entrega solo las celdas con contenido: hay que
                        # rellenar los huecos hasta la columna real de cada una
                        densa: "list[Any]" = []
                        for celda in fila:
                            while len(densa) < celda.c:
                                densa.append(None)
                            densa.append(celda.v)
                        filas.append(densa)
                    hojas.append((nombre, filas))
        return hojas

    if ext in (".xlsx", ".xlsm", ".xltx"):
        try:
            import openpyxl
        except ImportError as err:
            raise RuntimeError(
                "Para leer .xlsx hace falta openpyxl. En una celda: !pip install openpyxl"
            ) from err
        wb = openpyxl.load_workbook(str(ruta), data_only=True, read_only=True)
        hojas = [(ws.title, [list(f) for f in ws.iter_rows(values_only=True)])
                 for ws in wb.worksheets]
        wb.close()
        return hojas

    raise RuntimeError(f"Extensión no soportada: {ruta.name} (usa .xlsx, .xlsm o .xlsb)")


# -----------------------------------------------------------------------------
# 4. Balanza de comprobación -> columna «Metodología local»
# -----------------------------------------------------------------------------


@dataclass
class LecturaBalanza:
    periodo: "str | None"
    valores: "dict[str, float]"
    detalle: "list[str]"
    faltantes: "list[str]"
    hoja: str


def parse_balanza(hojas: Sequence, fname: str) -> "LecturaBalanza | None":
    """Cuentas 2205 / 2301 / 2302 en la columna de 3er grado.

    Las columnas no se amarran a letras fijas: se localiza la fila cuya primera
    celda dice CUENTA, se leen de ahí las columnas que dicen GRADO y el ordinal
    de la fila de arriba dice cuál de ellas es el 3er grado.
    """
    for nombre, filas in hojas:
        h = next((i for i, f in enumerate(filas[:60]) if f and norm(f[0]) == "cuenta"), None)
        if h is None:
            continue  # sin encabezado CUENTA no es una balanza

        grados: "dict[int, int]" = {}
        hdr = filas[h]
        arriba = filas[h - 1] if h > 0 else []
        for j in range(1, len(hdr)):
            if "grado" in norm(hdr[j]):
                celda_arriba = arriba[j] if j < len(arriba) and arriba[j] is not None else ""
                m = re.search(r"(\d)", str(celda_arriba))
                if m:
                    grados[int(m.group(1))] = j
        cols_grado = list(grados.values()) or list(range(2, 9))

        cuentas: "dict[str, list[Any]]" = {}
        for fila in filas[h + 1:]:
            if not fila or fila[0] is None:
                continue
            clave = str(fila[0]).strip()
            if clave and clave not in cuentas:
                cuentas[clave] = fila

        # periodo: encabezado del reporte; si no lo trae, el nombre del archivo
        periodo = None
        for fila in filas[:h + 1]:
            periodo = parse_fecha(" ".join("" if c is None else str(c) for c in (fila or [])))
            if periodo:
                break
        if not periodo:
            periodo = periodo_de_nombre(fname)

        valores: "dict[str, float]" = {}
        detalle: "list[str]" = []
        faltantes: "list[str]" = []
        for c in CONCEPTOS:
            fila = cuentas.get(c.cuenta)
            if fila is None:
                faltantes.append(c.cuenta)
                continue
            val = col_usada = None
            j = grados.get(GRADO)
            if j is not None and j < len(fila):
                val = to_num(fila[j])
                if val is not None:
                    col_usada = GRADO
            if val is None:  # respaldo: primer grado con importe en esa misma fila
                for col in cols_grado:
                    if col >= len(fila):
                        continue
                    v = to_num(fila[col])
                    if v is not None:
                        val = v
                        col_usada = next((g for g, cc in grados.items() if cc == col), None)
                        break
            if val is None:
                faltantes.append(c.cuenta)
                continue
            valores[c.id] = abs(val)   # el pasivo viene en negativo: se invierte el signo
            detalle.append(f"{c.cuenta} {f'{col_usada}º grado ' if col_usada else ''}{fmt(abs(val))}")

        if not valores:
            continue
        return LecturaBalanza(periodo, valores, detalle, faltantes, nombre)

    return None   # no es balanza: se intentará leer como archivo de actuarios


# -----------------------------------------------------------------------------
# 5. Archivo de actuarios -> columna «CNSF Método Estatutario»
# -----------------------------------------------------------------------------


@dataclass
class LecturaActuarios:
    periodos: "dict[str, dict[str, dict[str, float]]]" = field(default_factory=dict)
    hojas: "list[str]" = field(default_factory=list)


def parse_actuarios(hojas: Sequence) -> LecturaActuarios:
    """Busca el texto de cada reserva y toma los dos primeros importes a su derecha.

    Los nombres de hoja no son de fiar (dicen «Marzo 2026» y traen hasta junio),
    así que se recorren todas y se cargan de una pasada los cortes que haya.
    El periodo sale del serial de fecha de la propia fila; si no, del título del
    bloque de arriba.
    """
    out = LecturaActuarios()
    for nombre, filas in hojas:
        titulo: "str | None" = None
        usada = False
        for fila in filas:
            if not fila:
                continue
            concepto: "Concepto | None" = None
            ci = -1
            for j, v in enumerate(fila):
                if not isinstance(v, str):
                    continue
                c = match_concepto(v)
                if c and ci < 0:
                    concepto, ci = c, j
                elif not c:
                    f = parse_fecha(v)      # título del bloque: «30 de junio de 2026»
                    if f:
                        titulo = f
            if concepto is None:
                continue

            periodo = next((p for p in (celda_a_fecha(fila[j]) for j in range(ci)) if p), None) or titulo
            if not periodo:
                continue

            nums: "list[float]" = []
            for j in range(ci + 1, len(fila)):
                n = to_num(fila[j])
                if n is not None:
                    nums.append(abs(n))
                if len(nums) == 2:
                    break
            if len(nums) < 2:
                continue

            bloque = out.periodos.setdefault(periodo, {"local": {}, "cnsf": {}})
            bloque["local"][concepto.id] = nums[0]   # metodología local (para contraste)
            bloque["cnsf"][concepto.id] = nums[1]    # método estatutario CNSF
            usada = True
        if usada:
            out.hojas.append(nombre)
    return out


# -----------------------------------------------------------------------------
# 6. Histórico: cada archivo actualiza su periodo y conserva los demás
# -----------------------------------------------------------------------------


class Historico:
    """El histórico mes con mes, guardado en un JSON junto al notebook."""

    def __init__(self, ruta: "str | Path" = HIST_JSON):
        self.ruta = Path(ruta)
        self.datos: "dict[str, dict[str, Any]]" = {}
        if self.ruta.exists():
            try:
                self.datos = json.loads(self.ruta.read_text(encoding="utf-8"))
            except (ValueError, OSError):
                self.datos = {}

    # ------------------------------------------------------------------ io
    def guardar(self) -> None:
        self.ruta.write_text(json.dumps(self.datos, indent=2, ensure_ascii=False), encoding="utf-8")

    def periodos(self) -> "list[str]":
        return sorted(self.datos)

    def _entrada(self, p: str) -> "dict[str, Any]":
        return self.datos.setdefault(p, {"periodo": p, "local": {}, "cnsf": {}, "origen": {}})

    # ------------------------------------------------------------- ingesta
    def procesar(self, ruta: "str | Path", hojas: "Sequence | None" = None) -> "list[str]":
        """Lee un Excel, detecta solo de cuál se trata y lo funde al histórico.

        Si el libro ya se leyó antes (la ventana lo hace para la vista previa),
        se pasa en `hojas` y no se vuelve a abrir el archivo.
        """
        ruta = Path(ruta)
        if hojas is None:
            hojas = leer_libro(ruta)
        avisos: "list[str]" = []

        bal = parse_balanza(hojas, ruta.name)
        if bal is not None:
            if not bal.periodo:
                raise ValueError(
                    f"{ruta.name}: se leyó la balanza pero no se identificó el periodo. "
                    "Renombra el archivo como Balanza_MMAAAA.xlsx."
                )
            e = self._entrada(bal.periodo)
            e["local"].update(bal.valores)
            e["origen"]["local"] = f"Balanza · {ruta.name}"
            self._revisar(e)
            avisos.append(
                f"{ruta.name} → balanza al {etiqueta_periodo(bal.periodo)} · " + " · ".join(bal.detalle)
                + (f" · SIN CUENTA {', '.join(bal.faltantes)}" if bal.faltantes else "")
            )
            return avisos

        act = parse_actuarios(hojas)
        if not act.periodos:
            raise ValueError(
                f"{ruta.name}: no se reconoció ni como balanza (falta la columna CUENTA) "
                "ni como archivo de actuarios (faltan los tres conceptos de reserva)."
            )
        for p in sorted(act.periodos):
            e = self._entrada(p)
            src = act.periodos[p]
            e["cnsf"].update(src["cnsf"])
            e["local_actuarios"] = src["local"]
            origen_local = e["origen"].get("local", "")
            if not origen_local or origen_local.startswith("Actuarios"):
                # la balanza manda para la columna local; esto solo rellena los
                # periodos que todavía no tienen balanza
                for cid, val in src["local"].items():
                    e["local"].setdefault(cid, val)
                if not origen_local:
                    e["origen"]["local"] = f"Actuarios · {ruta.name}"
            e["origen"]["cnsf"] = f"Actuarios · {ruta.name}"
            self._revisar(e)
        avisos.append(
            f"{ruta.name} → actuarios, {len(act.periodos)} corte(s): "
            + ", ".join(etiqueta_corta(p) for p in sorted(act.periodos))
            + " · hojas: " + ", ".join(act.hojas)
        )
        avisos += [f"Revisar {etiqueta_periodo(p)}: {self.datos[p]['aviso']}"
                   for p in self.periodos() if self.datos[p].get("aviso")]
        return avisos

    @staticmethod
    def _revisar(e: "dict[str, Any]") -> None:
        """Contrasta la columna local de la balanza contra la del archivo de actuarios."""
        e.pop("aviso", None)
        ref = e.get("local_actuarios")
        if not ref:
            return
        difs = [
            f"{c.label.replace('Reserva de ', '')}: balanza {fmt(e['local'][c.id])} "
            f"vs actuarios {fmt(ref[c.id])}"
            for c in CONCEPTOS
            if c.id in e["local"] and c.id in ref and abs(e["local"][c.id] - ref[c.id]) > 0.5
        ]
        if difs:
            e["aviso"] = "Metodología local no coincide — " + "; ".join(difs)

    # -------------------------------------------------------------- cálculo
    def total(self, periodo: str, lado: str) -> "float | None":
        """Suma de los tres conceptos. None si ese lado del corte aún está vacío:
        un mes al que solo se le cargó la balanza no tiene columna estatutaria,
        y un cero fabricado se leería como «la reserva vale cero»."""
        vals = [self.datos[periodo][lado][c.id] for c in CONCEPTOS
                if c.id in self.datos[periodo][lado]]
        return sum(vals) if vals else None

    def diferencia(self, periodo: str, cid: "str | None" = None) -> "float | None":
        """Método estatutario CNSF − metodología local (exceso de constitución)."""
        e = self.datos[periodo]
        if cid is None:
            tc, tl = self.total(periodo, "cnsf"), self.total(periodo, "local")
            return None if tc is None or tl is None else tc - tl
        if cid not in e["cnsf"] or cid not in e["local"]:
            return None
        return e["cnsf"][cid] - e["local"][cid]

    def completo(self, periodo: str) -> bool:
        """¿El corte tiene las dos columnas y se puede comparar?"""
        return self.diferencia(periodo) is not None

    def incremento(self, actual: str, previo: str, cid: "str | None" = None) -> "float | None":
        a, b = self.diferencia(actual, cid), self.diferencia(previo, cid)
        return None if a is None or b is None else a - b

    # -------------------------------------------------------------- mensajes
    def mensajes(self, actual: str, previo: "str | None") -> "list[str]":
        """Los tres mensajes clave, redactados con las cifras del propio corte."""
        d1 = self.diferencia(actual)
        if d1 is None:
            falta = "el archivo de actuarios" if self.total(actual, "cnsf") is None else "la balanza"
            return [
                f"Al {etiqueta_periodo(actual)} solo se ha cargado una de las dos fuentes: "
                f"falta {falta} de ese corte, así que todavía no hay diferencia que comparar.",
                "Los cortes anteriores del histórico se conservan intactos.",
            ]
        if previo is None or not self.completo(previo):
            return [
                f"Al {etiqueta_periodo(actual)}, la diferencia entre metodologías "
                f"asciende a USD {mm(d1)} MM.",
                "Agrega un corte anterior completo al histórico para comparar la evolución "
                "del diferencial.",
            ]

        d0 = self.diferencia(previo) or 0.0
        v_local = self.total(actual, "local") - self.total(previo, "local")
        v_cnsf = self.total(actual, "cnsf") - self.total(previo, "cnsf")
        rango = f"Entre {etiqueta_corta(previo).lower()} y {etiqueta_corta(actual).lower()}"

        motor = max(
            ((c, self.datos[actual]["local"].get(c.id, 0.0) - self.datos[previo]["local"].get(c.id, 0.0))
             for c in CONCEPTOS), key=lambda x: abs(x[1]), default=None)
        m1 = (f"{rango}, las reservas bajo QES Metodología local "
              f"{'aumentan' if v_local >= 0 else 'disminuyen'} USD {mm(abs(v_local))} MM")
        if motor and abs(motor[1]) > 5000:
            m1 += (f", principalmente por {'el incremento' if motor[1] >= 0 else 'la reducción'} "
                   f"de USD {mm(abs(motor[1]))} MM en la "
                   f"{motor[0].label.replace('Reserva de ', 'reserva de ')}")
        m1 += "."

        sube, baja = [], []
        for c in CONCEPTOS:
            d = self.datos[actual]["cnsf"].get(c.id, 0.0) - self.datos[previo]["cnsf"].get(c.id, 0.0)
            if abs(d) > 5000:
                (sube if d > 0 else baja).append((c, abs(d)))

        def lista(arr):
            return " y de ".join(
                f"USD {mm(d)} MM en {c.label.replace('Reserva de ', '').lower()}" for c, d in arr)

        m2 = (f"Bajo el Método Estatutario, las reservas totales "
              f"{'aumentan' if v_cnsf >= 0 else 'disminuyen'} USD {mm(abs(v_cnsf))} MM.")
        if sube and baja:
            m2 += (f" El incremento de {lista(sube)} fue "
                   f"{'parcialmente compensado' if v_cnsf >= 0 else 'más que compensado'} "
                   f"por la disminución de {lista(baja)}.")
        elif sube:
            m2 += f" El movimiento se concentra en el incremento de {lista(sube)}."
        elif baja:
            m2 += f" El movimiento se concentra en la disminución de {lista(baja)}."

        v = d1 - d0
        m3 = ("El Método Estatutario mantiene una posición superior." if d1 >= 0
              else "La Metodología local se mantiene por encima del Método Estatutario.")
        m3 += (f" La diferencia total por constitución pasa de USD {mm(d0)} MM en "
               f"{etiqueta_corta(previo).lower()} a USD {mm(d1)} MM en "
               f"{etiqueta_corta(actual).lower()}, con "
               f"{'un incremento' if v >= 0 else 'una disminución'} de USD {mm(abs(v))} MM")
        if d0:
            m3 += f" ({'+' if v >= 0 else '−'}{abs(v / d0 * 100):.1f}%)"
        m3 += "."
        return [m1, m2, m3]

    # ---------------------------------------------------------------- texto
    def vista_texto(self, periodos: "Iterable[str] | None" = None) -> str:
        """La misma matriz, en texto, para revisar sin salir del notebook."""
        ps = list(periodos) if periodos else self.periodos()[-PERIODOS_EN_VISTA:]
        if not ps:
            return "Histórico vacío."
        ancho = max(len(c.label) for c in CONCEPTOS) + 2
        sep = "  "
        cab = "RESERVA".ljust(ancho) + sep + sep.join(
            "local".rjust(11) + sep + "CNSF".rjust(11) + sep + "dif.".rjust(11) for _ in ps)
        top = "".ljust(ancho) + sep + sep.join(etiqueta_corta(p).center(37) for p in ps)
        if len(ps) > 1:
            top += sep + f"vs {etiqueta_corta(ps[-2])}".center(11)
            cab += sep + "incremento".rjust(11)
        lineas = [top, cab, "-" * len(cab)]
        for c in CONCEPTOS:
            fila = c.label.ljust(ancho)
            for p in ps:
                e = self.datos[p]
                fila += (sep + celda_mm(e["local"].get(c.id)).rjust(11)
                         + sep + celda_mm(e["cnsf"].get(c.id)).rjust(11)
                         + sep + celda_mm(self.diferencia(p, c.id)).rjust(11))
            if len(ps) > 1:
                fila += sep + celda_mm(self.incremento(ps[-1], ps[-2], c.id)).rjust(11)
            lineas.append(fila)
        fila = "TOTAL RESERVAS".ljust(ancho)
        for p in ps:
            fila += (sep + celda_mm(self.total(p, "local")).rjust(11)
                     + sep + celda_mm(self.total(p, "cnsf")).rjust(11)
                     + sep + celda_mm(self.diferencia(p)).rjust(11))
        if len(ps) > 1:
            a, b = self.diferencia(ps[-1]), self.diferencia(ps[-2])
            fila += sep + celda_mm(None if a is None or b is None else a - b).rjust(11)
        lineas += ["-" * len(cab), fila, "",
                   "Millones de USD. Diferencia = Método Estatutario CNSF − Metodología local."]
        return "\n".join(lineas)


# -----------------------------------------------------------------------------
# 7. La vista en HTML autocontenido
# -----------------------------------------------------------------------------

# Sin tipografías ni scripts externos: el archivo se manda por correo y se ve igual.
CSS = """
:root{
  --plum:%(plum)s; --plum-deep:%(plum_deep)s; --plum-soft:%(plum_soft)s;
  --teal:%(teal)s; --teal-2:%(teal2)s; --teal-soft:%(teal_soft)s;
  --ice:%(ice)s; --ice-2:%(ice2)s; --paper:%(paper)s; --ground:%(ground)s;
  --ink:%(ink)s; --ink-2:%(ink2)s; --ink-3:%(ink3)s;
  --line:%(line)s; --line-soft:%(line_soft)s; --neg:%(neg)s;
  --shadow:0 1px 2px rgba(19,78,99,.08), 0 8px 24px rgba(19,78,99,.06);
  --sans:"Segoe UI",-apple-system,BlinkMacSystemFont,Roboto,Helvetica,Arial,sans-serif;
  --display:"Archivo Narrow","Arial Narrow","Segoe UI Semibold",var(--sans);
  --mono:Consolas,"SFMono-Regular",Menlo,"Liberation Mono",monospace;
}
*{box-sizing:border-box}
body{margin:0;background:var(--ground);color:var(--ink);font-family:var(--sans);
     font-size:15px;line-height:1.5;-webkit-font-smoothing:antialiased}
.wrap{max-width:1260px;margin:0 auto;padding:28px 20px 56px;display:flex;flex-direction:column;gap:22px}
h1,h2,h3{margin:0;font-family:var(--display);letter-spacing:.01em}
p{margin:0}

.board{background:var(--paper);border:1px solid var(--line);border-radius:6px;
       box-shadow:var(--shadow);overflow:hidden}
.board-head{padding:22px 24px 18px;display:grid;grid-template-columns:minmax(280px,1fr) auto;
            gap:20px;align-items:start;border-bottom:3px solid var(--plum)}
@media (max-width:940px){.board-head{grid-template-columns:1fr}}
.board-head h2{font-size:31px;font-weight:700;color:var(--plum);text-transform:uppercase;line-height:1}
.board-head .tag{font-size:14.5px;color:var(--ink-2);margin-top:4px}
.board-head .fx{font-family:var(--mono);font-size:12.5px;color:var(--teal);margin-top:10px}

.kpis{display:flex;gap:12px;flex-wrap:wrap}
.kpi{border:1px solid var(--line);border-radius:5px;padding:11px 16px;min-width:176px;background:var(--ice-2)}
.kpi .k-label{font-size:10.5px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:var(--teal-2)}
.kpi .k-scope{font-size:11.5px;color:var(--ink-3)}
.kpi .k-value{font-family:var(--display);font-size:27px;font-weight:700;color:var(--teal);
              line-height:1.1;margin-top:4px;font-variant-numeric:tabular-nums}
.kpi .k-value .u{font-size:14px;color:var(--teal-2)}
.kpi .k-alt{font-family:var(--mono);font-size:11.5px;color:var(--ink-3)}
.kpi.accent{background:#FBF1F7;border-color:#E4C6D8}
.kpi.accent .k-label{color:var(--plum-soft)}
.kpi.accent .k-value{color:var(--plum)}
.kpi.accent .k-value .u{color:var(--plum-soft)}

.table-wrap{overflow-x:auto}
table.matrix{border-collapse:collapse;width:100%%;min-width:940px;font-variant-numeric:tabular-nums}
table.matrix th,table.matrix td{padding:9px 12px;font-size:13px;border-bottom:1px solid var(--line-soft)}
table.matrix thead th{color:#fff;font-family:var(--sans);font-weight:600;font-size:12px;
                      line-height:1.25;text-align:center;border-bottom:0}
table.matrix thead .grp{background:var(--teal);font-family:var(--display);font-size:16px;
                        letter-spacing:.02em;border-left:2px solid var(--paper)}
table.matrix thead .sub-h{background:var(--teal-2);border-left:1px solid rgba(255,255,255,.25)}
table.matrix thead .sub-h.first{border-left:2px solid var(--paper)}
table.matrix thead .rowhead{background:var(--plum);text-align:left;font-family:var(--display);
                            font-size:19px;text-transform:uppercase;vertical-align:middle;padding-left:16px}
table.matrix thead .delta-h{background:var(--teal);border-left:2px solid var(--paper);vertical-align:middle}
table.matrix tbody th{text-align:left;font-weight:500;font-size:13px;color:var(--ink);
                      padding-left:16px;background:var(--paper)}
table.matrix tbody td{text-align:right;font-family:var(--mono);color:var(--ink)}
table.matrix tbody tr:nth-child(even) th,table.matrix tbody tr:nth-child(even) td{background:var(--ice-2)}
table.matrix td.gstart{border-left:2px solid var(--line)}
table.matrix td.dif{color:var(--plum);font-weight:500}
table.matrix td.delta{border-left:2px solid var(--line);color:var(--teal)}
table.matrix tr.total th,table.matrix tr.total td{background:var(--teal)!important;color:#fff;
                                                  font-weight:700;font-size:14px;border-bottom:0}
table.matrix tr.total th{font-family:var(--display);font-size:17px;text-transform:uppercase}
table.matrix tr.total td.dif,table.matrix tr.total td.delta{color:#fff}
.board-foot{padding:10px 24px 16px;font-size:12px;color:var(--ink-2);display:flex;
            flex-direction:column;gap:3px}

.band{display:grid;grid-template-columns:minmax(330px,1.35fr) minmax(300px,1fr);
      gap:0;border-top:1px solid var(--line-soft)}
@media (max-width:940px){.band{grid-template-columns:1fr}}
.panel{padding:20px 24px;display:flex;flex-direction:column;gap:12px}
.panel+.panel{border-left:1px solid var(--line-soft)}
@media (max-width:940px){.panel+.panel{border-left:0;border-top:1px solid var(--line-soft)}}
.panel h3{font-family:var(--sans);font-size:12px;font-weight:700;letter-spacing:.09em;
          text-transform:uppercase;color:var(--teal-2)}
.panel h3.plum{color:var(--plum-soft)}
.chart-note{font-size:12px;color:var(--ink-3);font-family:var(--mono)}
svg.wf{width:100%%;max-width:660px;height:auto;display:block}

ul.keys{list-style:none;margin:0;padding:0;display:flex;flex-direction:column;gap:11px}
ul.keys li{display:grid;grid-template-columns:16px 1fr;gap:9px;font-size:13.5px;line-height:1.5}
ul.keys li::before{content:"";width:13px;height:13px;margin-top:5px;border-radius:50%%;
                   background:var(--teal-soft)}
ul.byres{list-style:none;margin:0;padding:0}
ul.byres li{display:flex;justify-content:space-between;align-items:baseline;gap:14px;
            padding:8px 0;border-bottom:1px dotted var(--line);font-size:13.5px}
ul.byres li:last-child{border-bottom:0;border-top:2px solid var(--plum);margin-top:2px;
                       padding-top:10px;font-weight:700;color:var(--plum)}
ul.byres .v{font-family:var(--mono);font-size:13px;white-space:nowrap;text-align:right}
ul.byres .v small{display:block;color:var(--ink-3);font-size:11.5px;font-weight:400}
ul.byres li>span>small{font-weight:400;color:var(--ink-3)}
.note-box{background:var(--ice);border-left:3px solid var(--teal-soft);border-radius:0 4px 4px 0;
          padding:12px 14px;font-size:13.5px;color:var(--ink)}
.flags{background:#FDF4E1;border:1px solid #EBD6A5;border-left:4px solid %(warn)s;border-radius:4px;
       padding:10px 14px;font-size:13px;color:#6B4A0A}
.flags ul{margin:6px 0 0;padding-left:18px}
footer.credits{font-size:12px;color:var(--ink-3);text-align:center;font-family:var(--mono)}
@media print{body{background:#fff}.board{border:0;box-shadow:none}footer.credits{display:none}}
""" % {
    "plum": PLUM, "plum_deep": PLUM_DEEP, "plum_soft": PLUM_SOFT,
    "teal": TEAL, "teal2": TEAL_2, "teal_soft": TEAL_SOFT,
    "ice": ICE, "ice2": ICE_2, "paper": PAPER, "ground": GROUND,
    "ink": INK, "ink2": INK_2, "ink3": INK_3,
    "line": LINE, "line_soft": LINE_SOFT, "neg": NEG, "warn": WARN,
}


def _svg_cascada(hist: Historico, actual: str, previo: "str | None") -> str:
    """Gráfico de cascada: corte anterior, incrementos por reserva y corte actual."""
    if not previo or not hist.completo(actual) or not hist.completo(previo):
        return ('<p class="chart-note">Hacen falta dos cortes con las dos fuentes cargadas '
                'para dibujar la cascada.</p>')

    W, H, L, R, T, B = 660, 330, 46, 16, 46, 74
    d0 = (hist.diferencia(previo) or 0.0) / 1e6
    d1 = (hist.diferencia(actual) or 0.0) / 1e6

    pasos = [("base", etiqueta_corta(previo), "Diferencia total", d0)]
    for c in CONCEPTOS:
        v = hist.incremento(actual, previo, c.id)
        if v is None or abs(v / 1e6) < 0.005:
            continue
        v /= 1e6
        pasos.append(("delta", c.label.replace("Reserva de ", ""),
                      "Incremento" if v >= 0 else "Disminución", v))
    pasos.append(("base", etiqueta_corta(actual), "Diferencia total", d1))

    corrida, geo, maxv = 0.0, [], max(d0, d1, 0.0001)
    for tipo, etq, sub, valor in pasos:
        if tipo == "base":
            geo.append((tipo, etq, sub, valor, 0.0, valor))
            corrida = valor
        else:
            geo.append((tipo, etq, sub, valor, corrida, corrida + valor))
            corrida += valor
            maxv = max(maxv, corrida)
    tope = maxv * 1.22

    def y(v: float) -> float:
        return T + (H - T - B) * (1 - v / tope)

    ancho_util = W - L - R
    hueco = ancho_util / len(geo)
    bw = min(70.0, hueco * 0.54)

    partes = [f'<svg class="wf" viewBox="0 0 {W} {H}" role="img" '
              f'aria-label="Puente de la diferencia entre metodologías">',
              f'<line x1="{L - 8}" y1="{y(0):.1f}" x2="{W - R}" y2="{y(0):.1f}" '
              f'stroke="{LINE}" stroke-width="1" />']

    for i, (tipo, etq, sub, valor, y0, y1) in enumerate(geo):
        cx = L + hueco * i + hueco / 2
        ya, yb = y(max(y0, y1)), y(min(y0, y1))
        alto = max(3.0, yb - ya)
        relleno = "#8E2A66" if tipo == "base" else (TEAL_2 if valor >= 0 else NEG)
        partes.append(f'<rect x="{cx - bw / 2:.1f}" y="{ya:.1f}" width="{bw:.1f}" '
                      f'height="{alto:.1f}" fill="{relleno}" rx="1" />')
        signo = "" if tipo == "base" else ("+" if valor >= 0 else "−")
        partes.append(f'<text x="{cx:.1f}" y="{ya - 9:.1f}" text-anchor="middle" '
                      f'fill="{PLUM if tipo == "base" else TEAL}" font-family="Consolas,monospace" '
                      f'font-size="13" font-weight="500">{signo}{abs(valor):,.2f}</text>')

        # etiqueta al pie, partida en renglones de ~18 caracteres
        renglones, actualr = [], ""
        for palabra in etq.split(" "):
            if len((actualr + " " + palabra).strip()) > 18:
                renglones.append(actualr.strip())
                actualr = palabra
            else:
                actualr += " " + palabra
        renglones.append(actualr.strip())
        tspans = "".join(
            f'<tspan x="{cx:.1f}" dy="{0 if k == 0 else 13}">{esc(t)}</tspan>'
            for k, t in enumerate(renglones))
        partes.append(f'<text y="{H - B + 22}" text-anchor="middle" fill="{INK}" '
                      f'font-family="Segoe UI,sans-serif" font-size="11.5">{tspans}</text>')
        partes.append(f'<text x="{cx:.1f}" y="{H - B + 22 + len(renglones) * 13}" '
                      f'text-anchor="middle" fill="{INK_3}" font-family="Segoe UI,sans-serif" '
                      f'font-size="10.5">{esc(sub)}</text>')

        if i < len(geo) - 1:
            yfin = y(y1)
            partes.append(f'<line x1="{cx + bw / 2:.1f}" y1="{yfin:.1f}" '
                          f'x2="{L + hueco * (i + 1) + hueco / 2 - bw / 2:.1f}" y2="{yfin:.1f}" '
                          f'stroke="#9FB6C2" stroke-width="1" stroke-dasharray="4 3" />')

    partes.append(f'<text x="{L - 8}" y="{T - 18}" fill="{INK_3}" '
                  f'font-family="Consolas,monospace" font-size="11">'
                  f'Diferencia acumulada (MM USD)</text>')
    partes.append("</svg>")
    return "\n".join(partes)


def construir_html(hist: Historico, periodos: "Sequence[str] | None" = None,
                   fx: float = FX_DEFAULT, nota: str = NOTA_RELEVANTE) -> str:
    """Arma la vista completa en un solo archivo HTML, sin dependencias externas."""
    ps = list(periodos) if periodos else hist.periodos()[-PERIODOS_EN_VISTA:]
    if not ps:
        raise ValueError("El histórico está vacío: no hay nada que dibujar.")
    actual = ps[-1]
    previo = ps[-2] if len(ps) > 1 else None

    def mmx(v: float) -> str:
        return f"{v * fx / 1e6:,.2f}"

    dif_actual = hist.diferencia(actual)
    d1 = dif_actual or 0.0

    # ---------------------------------------------------------- indicadores
    if dif_actual is None:
        falta = "el archivo de actuarios" if hist.total(actual, "cnsf") is None else "la balanza"
        kpis = [f'<div class="kpi"><div class="k-label">Diferencia total</div>'
                f'<div class="k-scope">{esc(etiqueta_corta(actual))}</div>'
                f'<div class="k-value">n/d</div>'
                f'<div class="k-alt">falta {falta} de este corte</div></div>']
    else:
        kpis = [f'<div class="kpi"><div class="k-label">Diferencia total</div>'
                f'<div class="k-scope">{esc(etiqueta_corta(actual))}</div>'
                f'<div class="k-value">USD {mm(d1)} <span class="u">MM</span></div>'
                f'<div class="k-alt">~MXN {mmx(d1)} MM</div></div>']
    if previo and hist.completo(actual) and hist.completo(previo):
        d0 = hist.diferencia(previo) or 0.0
        v = d1 - d0
        alcance = f"{etiqueta_corta(previo)} → {etiqueta_corta(actual)}"
        kpis.append(f'<div class="kpi accent"><div class="k-label">Variación del periodo</div>'
                    f'<div class="k-scope">{esc(alcance)}</div>'
                    f'<div class="k-value">{"+" if v >= 0 else "−"}USD {mm(abs(v))} '
                    f'<span class="u">MM</span></div>'
                    f'<div class="k-alt">~MXN {mmx(abs(v))} MM</div></div>')
        pct = f'{"+" if v >= 0 else "−"}{abs(v / d0 * 100):.1f}%' if d0 else "n/d"
        kpis.append(f'<div class="kpi"><div class="k-label">Variación % de la diferencia</div>'
                    f'<div class="k-scope">{esc(alcance)}</div>'
                    f'<div class="k-value">{pct}</div>'
                    f'<div class="k-alt">sobre USD {mm(d0)} MM</div></div>')

    # ---------------------------------------------------------------- matriz
    h1 = '<tr><th class="rowhead" rowspan="2">Reserva</th>'
    h2 = "<tr>"
    for p in ps:
        h1 += f'<th class="grp" colspan="3">{esc(etiqueta_periodo(p))}</th>'
        h2 += ('<th class="sub-h first">Metodología<br />local</th>'
               '<th class="sub-h">CNSF<br />Método Estatutario</th>'
               '<th class="sub-h">Diferencia<br />(estatutario − local)</th>')
    if previo:
        h1 += ('<th class="delta-h" rowspan="2">Incremento<br />respecto de<br />'
               f'{esc(etiqueta_corta(previo))}</th>')
    h1 += "</tr>"
    h2 += "</tr>"

    cuerpo = ""
    for c in CONCEPTOS:
        cuerpo += f"<tr><th>{esc(c.label)}{' *' if c.nota else ''}</th>"
        for p in ps:
            e = hist.datos[p]
            cuerpo += (f'<td class="gstart">{celda_mm(e["local"].get(c.id))}</td>'
                       f'<td>{celda_mm(e["cnsf"].get(c.id))}</td>'
                       f'<td class="dif">{celda_mm(hist.diferencia(p, c.id))}</td>')
        if previo:
            cuerpo += f'<td class="delta">{celda_mm(hist.incremento(actual, previo, c.id))}</td>'
        cuerpo += "</tr>"

    cuerpo += '<tr class="total"><th>Total reservas</th>'
    for p in ps:
        cuerpo += (f'<td class="gstart">{celda_mm(hist.total(p, "local"))}</td>'
                   f'<td>{celda_mm(hist.total(p, "cnsf"))}</td>'
                   f'<td class="dif">{celda_mm(hist.diferencia(p))}</td>')
    if previo:
        da, db = hist.diferencia(actual), hist.diferencia(previo)
        cuerpo += ('<td class="delta">'
                   f'{celda_mm(None if da is None or db is None else da - db)}</td>')
    cuerpo += "</tr>"

    # ------------------------------------------------------- mensajes y listas
    mensajes = "".join(f"<li>{esc(m)}</li>" for m in hist.mensajes(actual, previo))

    por_reserva = ""
    for c in CONCEPTOS:
        d = hist.diferencia(actual, c.id)
        if d is None or abs(d) < 5000:
            continue
        por_reserva += (f"<li><span>{esc(c.label)}</span>"
                        f'<span class="v">USD {mm(d)} MM<small>~MXN {mmx(d)} MM</small></span></li>')
    if dif_actual is None:
        por_reserva = ('<li><span>Este corte todavía no tiene las dos fuentes cargadas.</span>'
                       '<span class="v">n/d</span></li>')
    else:
        por_reserva += ('<li><span>Total de reservas<br />'
                        '<small>incremento por constitución</small></span>'
                        f'<span class="v">USD {mm(d1)} MM<small>~MXN {mmx(d1)} MM</small></span></li>')

    avisos = [f"{etiqueta_periodo(p)}: {hist.datos[p]['aviso']}"
              for p in ps if hist.datos[p].get("aviso")]
    banda_avisos = ""
    if avisos:
        banda_avisos = ('<div class="flags"><strong>Revisar la metodología local</strong>'
                        "<ul>" + "".join(f"<li>{esc(a)}</li>" for a in avisos) + "</ul></div>")

    origen = hist.datos[actual].get("origen", {})
    pie = " · ".join(filter(None, [origen.get("local"), origen.get("cnsf")])) or "—"
    sello = dt.datetime.now().strftime("%d/%m/%Y %H:%M")

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Reservas técnicas QES · {esc(etiqueta_corta(actual))}</title>
<style>{CSS}</style>
</head>
<body>
<div class="wrap">
{banda_avisos}
  <section class="board">
    <div class="board-head">
      <div>
        <h2>Resultados reservas técnicas QES</h2>
        <p class="tag">Comparativo de metodologías y evolución del diferencial</p>
        <p class="fx">Tipo de cambio: {fx:,.4f} MXN / USD</p>
      </div>
      <div class="kpis">{"".join(kpis)}</div>
    </div>

    <div class="table-wrap">
      <table class="matrix">
        <thead>{h1}{h2}</thead>
        <tbody>{cuerpo}</tbody>
      </table>
    </div>
    <div class="board-foot">
      <span>Cifras en millones de USD. * {esc(NOTA_PIE)}</span>
      <span>La diferencia se presenta como Método Estatutario CNSF menos Metodología local,
            es decir el exceso de constitución del método estatutario.</span>
    </div>

    <div class="band">
      <div class="panel">
        <h3>Evolución de la diferencia entre metodologías</h3>
        <p class="chart-note">Millones de USD{
            " · " + esc(etiqueta_corta(previo)) + " → " + esc(etiqueta_corta(actual)) if previo else ""}</p>
        {_svg_cascada(hist, actual, previo)}
      </div>
      <div class="panel">
        <h3>Mensajes clave</h3>
        <ul class="keys">{mensajes}</ul>
      </div>
    </div>

    <div class="band">
      <div class="panel">
        <h3 class="plum">Diferencias a {esc(etiqueta_corta(actual).lower())} por reserva
            (estatutario − local)</h3>
        <ul class="byres">{por_reserva}</ul>
      </div>
      <div class="panel">
        <h3>Nota relevante</h3>
        <div class="note-box">{esc(nota)}</div>
      </div>
    </div>
  </section>

  <footer class="credits">Reservas técnicas QES · armado en local el {sello} · fuentes: {esc(pie)}</footer>
</div>
</body>
</html>
"""


def escribir_vista(hist: Historico, destino: "str | Path | None" = None,
                   periodos: "Sequence[str] | None" = None, fx: float = FX_DEFAULT,
                   nota: str = NOTA_RELEVANTE, abrir: bool = True) -> Path:
    """Escribe el HTML junto al notebook y lo abre en el navegador."""
    ps = list(periodos) if periodos else hist.periodos()[-PERIODOS_EN_VISTA:]
    if not ps:
        raise ValueError("El histórico está vacío: carga al menos un archivo.")
    ruta = Path(destino) if destino else hist.ruta.parent / f"vista_reservas_{ps[-1]}.html"
    ruta.write_text(construir_html(hist, ps, fx, nota), encoding="utf-8")
    if abrir:
        try:
            webbrowser.open_new_tab(ruta.resolve().as_uri())
        except Exception:
            pass
    return ruta


# -----------------------------------------------------------------------------
# 8. La ventana
# -----------------------------------------------------------------------------

_SIN_PANTALLA = """
No se pudo abrir la ventana: este Python no tiene una pantalla a la mano ({err}).

Suele pasar cuando el kernel de Jupyter corre en un servidor, en WSL o en un
contenedor: ahí no hay escritorio donde dibujar. Opciones:

  * Ejecutar el notebook en tu propia máquina (Anaconda / Jupyter local).
  * Trabajar sin ventana desde esta misma celda:

        h = Historico("historico_reservas.json")
        h.procesar(r"C:\\ruta\\Balanza_062026.xlsx")
        h.procesar(r"C:\\ruta\\ResultadosQES.xlsb")
        h.guardar()
        print(h.vista_texto())
        escribir_vista(h)          # escribe y abre el HTML
"""


def abrir_ventana(hist_ruta: "str | Path" = HIST_JSON, bloquear: bool = True):
    """Abre la ventana de carga. Es lo que corre la última línea del bloque.

    Se cargan los dos Excel (botón, o arrastrando si está instalado tkinterdnd2),
    se ve qué se leyó de cada uno y el botón «Procesar» escribe el HTML.
    """
    import tkinter as tk
    from tkinter import filedialog, ttk

    try:
        try:
            from tkinterdnd2 import DND_FILES, TkinterDnD
            root = TkinterDnD.Tk()
            arrastre = True
        except ImportError:
            root = tk.Tk()
            arrastre = False
    except tk.TclError as err:
        print(_SIN_PANTALLA.format(err=err))
        return None

    import tkinter.font as tkfont
    familias = set(tkfont.families(root))
    UI = ("Segoe UI" if "Segoe UI" in familias
          else "Helvetica" if "Helvetica" in familias else "TkDefaultFont")
    MONO = "Consolas" if "Consolas" in familias else "TkFixedFont"

    root.title("Reservas técnicas QES")
    root.configure(bg=GROUND)
    # nada de tamaños fijos: en una laptop de 1366x768 la ventana se saldría de la pantalla
    pw, ph = root.winfo_screenwidth(), root.winfo_screenheight()
    ancho = max(560, min(1000, pw - 80))
    alto = max(420, min(720, ph - 120))
    root.geometry(f"{ancho}x{alto}+{max(0, (pw - ancho) // 2)}+{max(0, (ph - alto) // 3)}")
    root.minsize(520, 400)

    hist = Historico(hist_ruta)
    pendientes: "dict[str, dict[str, Any]]" = {}     # ruta -> {tipo, hojas, resumen}
    estado: "dict[str, Any]" = {"fx": FX_DEFAULT, "nota": NOTA_RELEVANTE}

    # ------------------------------------------------------------ encabezado
    cab = tk.Frame(root, bg=GROUND)
    cab.pack(fill="x", padx=16, pady=(14, 6))
    tk.Label(cab, text="RESERVAS TÉCNICAS QES", bg=GROUND, fg=PLUM,
             font=(UI, 16, "bold")).pack(anchor="w")
    tk.Label(cab, text="Carga la balanza de comprobación y el archivo de actuarios; "
                       "«Procesar» escribe la vista en HTML y la abre en el navegador.",
             bg=GROUND, fg=INK_2, font=(UI, 9), wraplength=ancho - 60,
             justify="left").pack(anchor="w")

    # ------------------------------------------------------------ zona de carga
    marco = tk.Frame(root, bg=PAPER, highlightbackground=LINE, highlightthickness=1)
    marco.pack(fill="x", padx=16, pady=6)
    zona = tk.Label(
        marco,
        text=("Arrastra aquí los dos Excel  ·  o pica para elegirlos"
              if arrastre else "Pica aquí para elegir los dos Excel"),
        bg=ICE_2, fg=TEAL, font=(UI, 11, "bold"), height=2,
        relief="ridge", bd=1, cursor="hand2")
    zona.pack(fill="x", padx=14, pady=(14, 4))
    tk.Label(marco, text="Balanza de comprobación (.xlsx) y resultados de actuarios (.xlsb). "
                         "El origen de cada archivo se detecta solo.",
             bg=PAPER, fg=INK_3, font=(UI, 8), wraplength=ancho - 80,
             justify="left").pack(anchor="w", padx=14, pady=(0, 10))

    # ------------------------------------------------------------ lo que se leyó
    lectura = tk.Frame(root, bg=PAPER, highlightbackground=LINE, highlightthickness=1)
    lectura.pack(fill="x", padx=16, pady=6)
    tk.Label(lectura, text="LO QUE SE LEYÓ DE CADA ARCHIVO", bg=PAPER, fg=INK_3,
             font=(UI, 8, "bold")).pack(anchor="w", padx=14, pady=(10, 4))
    tarjetas = tk.Frame(lectura, bg=PAPER)
    tarjetas.pack(fill="x", padx=14, pady=(0, 12))

    # ------------------------------------------------------------ controles
    ctl = tk.Frame(root, bg=GROUND)
    ctl.pack(fill="x", padx=16, pady=(4, 0))
    tk.Label(ctl, text="Tipo de cambio MXN/USD", bg=GROUND, fg=INK_3,
             font=(UI, 8, "bold")).pack(side="left")
    fx_var = tk.StringVar(value=f"{FX_DEFAULT:.4f}")
    tk.Entry(ctl, textvariable=fx_var, width=10, font=(MONO, 9)).pack(side="left", padx=(8, 18))
    btn_procesar = ttk.Button(ctl, text="Procesar")
    btn_procesar.pack(side="right")
    ttk.Button(ctl, text="Vaciar histórico",
               command=lambda: vaciar()).pack(side="right", padx=(0, 8))

    # ------------------------------------------------------------ bitácora
    caja = tk.Frame(root, bg=PAPER, highlightbackground=LINE, highlightthickness=1)
    caja.pack(fill="both", expand=True, padx=16, pady=(10, 16))
    tk.Label(caja, text="BITÁCORA", bg=PAPER, fg=INK_3,
             font=(UI, 8, "bold")).pack(anchor="w", padx=14, pady=(10, 2))
    envoltura = tk.Frame(caja, bg=PAPER)
    envoltura.pack(fill="both", expand=True, padx=14, pady=(0, 12))
    barra = ttk.Scrollbar(envoltura, orient="vertical")
    bitacora = tk.Text(envoltura, bg=PAPER, fg=INK_2, font=(MONO, 8), relief="flat",
                       wrap="word", height=7, yscrollcommand=barra.set,
                       highlightbackground=LINE, highlightthickness=1)
    barra.configure(command=bitacora.yview)
    barra.pack(side="right", fill="y")
    bitacora.pack(side="left", fill="both", expand=True)
    bitacora.tag_configure("err", foreground=NEG)
    bitacora.tag_configure("warn", foreground=WARN)
    bitacora.tag_configure("ok", foreground="#0E7C66")
    bitacora.configure(state="disabled")

    def apunta(texto: str, tag: str = "") -> None:
        bitacora.configure(state="normal")
        bitacora.insert("end", texto + "\n", tag)
        bitacora.see("end")
        bitacora.configure(state="disabled")

    def reporta_error(tipo, valor, rastro) -> None:
        # los errores de los callbacks de Tk se van a stderr, que desde Jupyter
        # acaba en la consola del servidor: invisible. Aquí se ven.
        import traceback
        apunta("! " + "".join(traceback.format_exception_only(tipo, valor)).strip(), "err")

    root.report_callback_exception = reporta_error

    # ------------------------------------------------------------ vista previa
    def pinta_tarjetas() -> None:
        for hijo in tarjetas.winfo_children():
            hijo.destroy()
        if not pendientes:
            tk.Label(tarjetas, text="Todavía no hay archivos cargados en esta sesión."
                                    + (f"  Histórico guardado: {len(hist.periodos())} corte(s)."
                                       if hist.periodos() else ""),
                     bg=PAPER, fg=INK_2, font=(UI, 9), justify="left",
                     wraplength=ancho - 80).pack(anchor="w")
            actualiza_boton()
            return
        for i, (ruta, info) in enumerate(pendientes.items()):
            t = tk.Frame(tarjetas, bg=ICE_2, highlightbackground=LINE, highlightthickness=1)
            t.pack(fill="x", pady=(0 if i == 0 else 6))
            tk.Label(t, text=info["titulo"], bg=ICE_2, fg=TEAL, font=(UI, 9, "bold"),
                     anchor="w").pack(fill="x", padx=10, pady=(6, 0))
            tk.Label(t, text=Path(ruta).name, bg=ICE_2, fg=INK_3, font=(MONO, 8),
                     anchor="w").pack(fill="x", padx=10)
            tk.Label(t, text=info["resumen"], bg=ICE_2, fg=INK, font=(MONO, 8),
                     anchor="w", justify="left", wraplength=ancho - 90).pack(
                fill="x", padx=10, pady=(2, 7))
        actualiza_boton()

    def actualiza_boton() -> None:
        listo = bool(pendientes) or bool(hist.periodos())
        btn_procesar.state(["!disabled"] if listo else ["disabled"])

    def carga_rutas(rutas: "Iterable[str]") -> None:
        for ruta in rutas:
            ruta = str(ruta).strip()
            if not ruta:
                continue
            try:
                hojas = leer_libro(ruta)
                bal = parse_balanza(hojas, Path(ruta).name)
                if bal is not None:
                    per = bal.periodo or periodo_de_nombre(Path(ruta).name)
                    resumen = " · ".join(bal.detalle) or "sin importes"
                    if bal.faltantes:
                        resumen += f"  ·  SIN CUENTA {', '.join(bal.faltantes)}"
                    pendientes[ruta] = {
                        "tipo": "balanza", "hojas": hojas,
                        "titulo": f"Balanza de comprobación  ·  {etiqueta_periodo(per) if per else 'periodo sin identificar'}",
                        "resumen": f"hoja «{bal.hoja}»  ·  {resumen}",
                    }
                    apunta(f"· {Path(ruta).name}: balanza leída "
                           f"({etiqueta_periodo(per) if per else 'periodo sin identificar'})", "ok")
                    if not per:
                        apunta("  no se identificó el periodo: renómbrala como "
                               "Balanza_MMAAAA.xlsx", "warn")
                    continue

                act = parse_actuarios(hojas)
                if not act.periodos:
                    apunta(f"! {Path(ruta).name}: no se reconoció ni como balanza (falta la "
                           "columna CUENTA) ni como archivo de actuarios (faltan los tres "
                           "conceptos de reserva).", "err")
                    continue
                cortes = sorted(act.periodos)
                pendientes[ruta] = {
                    "tipo": "actuarios", "hojas": hojas,
                    "titulo": f"Archivo de actuarios  ·  {len(cortes)} corte(s)",
                    "resumen": ("cortes: " + ", ".join(etiqueta_corta(p) for p in cortes)
                                + "\nhojas: " + ", ".join(act.hojas)),
                }
                apunta(f"· {Path(ruta).name}: actuarios, {len(cortes)} corte(s) "
                       f"({', '.join(etiqueta_corta(p) for p in cortes)})", "ok")
            except Exception as err:      # se reporta en la bitácora, la ventana sigue viva
                apunta(f"! {Path(ruta).name}: {err}", "err")
        pinta_tarjetas()

    def elegir() -> None:
        rutas = filedialog.askopenfilenames(
            title="Elige la balanza (.xlsx) y el archivo de actuarios (.xlsb)",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xlsb"), ("Todos", "*.*")])
        if rutas:
            carga_rutas(root.tk.splitlist(rutas) if isinstance(rutas, str) else rutas)

    zona.bind("<Button-1>", lambda e: elegir())
    if arrastre:
        zona.drop_target_register(DND_FILES)
        zona.dnd_bind("<<Drop>>", lambda e: carga_rutas(root.tk.splitlist(e.data)))
        zona.dnd_bind("<<DragEnter>>", lambda e: zona.configure(bg=ICE))
        zona.dnd_bind("<<DragLeave>>", lambda e: zona.configure(bg=ICE_2))

    # ------------------------------------------------------------ procesar
    def procesar() -> None:
        try:
            fx = float(fx_var.get().replace(",", ""))
            if fx <= 0:
                raise ValueError
        except ValueError:
            apunta(f"! Tipo de cambio no válido: «{fx_var.get()}». Se usa {FX_DEFAULT}.", "warn")
            fx = FX_DEFAULT

        # la balanza manda para la columna local: se funde después de los actuarios
        orden = sorted(pendientes.items(), key=lambda kv: kv[1]["tipo"] != "actuarios")
        for ruta, info in orden:
            try:
                for aviso in hist.procesar(ruta, info["hojas"]):
                    apunta("· " + aviso,
                           "warn" if "Revisar" in aviso or "SIN CUENTA" in aviso else "")
            except Exception as err:
                apunta(f"! {Path(ruta).name}: {err}", "err")
        if not hist.periodos():
            apunta("! No hay ningún corte en el histórico: carga al menos un archivo.", "err")
            return

        hist.guardar()
        pendientes.clear()
        pinta_tarjetas()
        ps = hist.periodos()[-PERIODOS_EN_VISTA:]
        apunta(f"· histórico guardado en {hist.ruta} ({len(hist.periodos())} corte(s))")
        try:
            ruta_html = escribir_vista(hist, periodos=ps, fx=fx, nota=estado["nota"])
        except Exception as err:
            apunta(f"! No se pudo escribir la vista: {err}", "err")
            return
        apunta(f"· vista de {', '.join(etiqueta_corta(p) for p in ps)} escrita en "
               f"{ruta_html}", "ok")
        apunta("· abierta en el navegador. El archivo es autocontenido: se puede mandar "
               "por correo tal cual.", "ok")

    def vaciar() -> None:
        from tkinter import messagebox
        if messagebox.askyesno("Reservas técnicas QES", "¿Vaciar todo el histórico guardado?"):
            hist.datos.clear()
            hist.guardar()
            apunta("· histórico vaciado", "warn")
            pinta_tarjetas()

    btn_procesar.configure(command=procesar)

    apunta(f"· histórico: {Path(hist_ruta).resolve()} ({len(hist.periodos())} corte(s))")
    if hist.periodos():
        apunta("  cortes guardados: " + ", ".join(etiqueta_corta(p) for p in hist.periodos()))
    if not arrastre:
        apunta("  (instala tkinterdnd2 si quieres arrastrar y soltar: pip install tkinterdnd2)",
               "warn")
    pinta_tarjetas()

    # que no nazca detrás del navegador ni del notebook
    try:
        root.update_idletasks()
        root.deiconify()
        root.lift()
        root.attributes("-topmost", True)
        root.after(800, lambda: root.attributes("-topmost", False))
        root.focus_force()
    except Exception:
        pass

    print(f"Ventana abierta ({ancho}x{alto}). La celda queda ocupada con [*] "
          "hasta que la cierres: es normal.")

    if bloquear:
        root.mainloop()
    else:
        root.update()      # con «%gui tk» IPython se encarga del resto
    return root


# -----------------------------------------------------------------------------
# 9. Sin ventana, por si el kernel corre en un servidor
# -----------------------------------------------------------------------------


def procesar_sin_ventana(*archivos: "str | Path", hist_ruta: "str | Path" = HIST_JSON,
                         fx: float = FX_DEFAULT, abrir: bool = True) -> Path:
    """Misma lectura y misma vista, sin interfaz: útil sin escritorio o por lotes."""
    hist = Historico(hist_ruta)
    # los actuarios primero: para la columna local manda la balanza
    for ruta in sorted(archivos, key=lambda r: Path(r).suffix.lower() != ".xlsb"):
        for aviso in hist.procesar(ruta):
            print("·", aviso)
    hist.guardar()
    print()
    print(hist.vista_texto())
    ruta_html = escribir_vista(hist, fx=fx, abrir=abrir)
    print(f"\n· vista escrita en {ruta_html}")
    return ruta_html


# Ni una palabra de argparse: en Jupyter vería el «-f <archivo de conexión>» del
# kernel y abortaría con SystemExit: 2.

abrir_ventana()
