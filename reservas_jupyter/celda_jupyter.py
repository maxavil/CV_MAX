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
import getpass
import hashlib
import html
import json
import os
import platform
import re
import sys
import unicodedata
import urllib.parse
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

# --- servidor de auditoría ---------------------------------------------------
SERVIDOR = "Qauditinterna"                    # nombre del servidor SQL
BASE = "PLD_492"                              # base de datos
DRIVER = "ODBC Driver 17 for SQL Server"
TABLA_CARGAS = "ReservasQES_Cargas"           # una fila por archivo subido
TABLA_DETALLE = "ReservasQES_Detalle"         # importes por corte y concepto

# Las tres ranuras de la vista, en el orden en que salen en la matriz.
RANURAS = [("dic", "Diciembre"), ("t1", "t-1"), ("t", "t")]

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


# --- las dos monedas, una al lado de la otra ---------------------------------
# La vista guarda las dos cifras y el CSS enseña la que el lector eligió: así el
# HTML sigue sin una línea de JavaScript y se puede mandar por correo tal cual.

def dual(usd: str, mxn: str) -> str:
    return f'<span class="v-usd">{usd}</span><span class="v-mxn">{mxn}</span>'


def dual_mm(v: "float | None", fx: float) -> str:
    """La celda de la matriz en las dos monedas, con guion donde el valor es cero."""
    return dual(celda_mm(v), celda_mm(None if v is None else v * fx))


def dual_texto(plantilla: str, v: float, fx: float) -> str:
    """`plantilla` lleva {sim} y {n}; se rellena para USD y para MXN."""
    return dual(
        esc(plantilla.format(sim="USD", n=f"{v / 1e6:,.2f}")),
        esc(plantilla.format(sim="MXN", n=f"{v * fx / 1e6:,.2f}")),
    )


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
# 5 bis. Una lectura, dos destinos: el histórico local y el servidor
# -----------------------------------------------------------------------------


@dataclass
class Lectura:
    """Lo que se sacó de un archivo, ya normalizado y sin importar de cuál venga."""
    fuente: str                                   # "balanza" | "actuarios"
    archivo: str
    hoja: str
    periodos: "dict[str, dict[str, dict[str, float]]]"   # periodo -> {local, cnsf}
    detalle: "list[str]"
    faltantes: "list[str]"
    sha256: str

    @property
    def cortes(self) -> "list[str]":
        return sorted(self.periodos)

    @property
    def titulo(self) -> str:
        if self.fuente == "balanza":
            p = self.cortes[0] if self.cortes else None
            return ("Balanza de comprobación  ·  "
                    + (etiqueta_periodo(p) if p else "periodo sin identificar"))
        return f"Archivo de actuarios  ·  {len(self.cortes)} corte(s)"

    @property
    def resumen(self) -> str:
        if self.fuente == "balanza":
            txt = " · ".join(self.detalle) or "sin importes"
            if self.faltantes:
                txt += f"  ·  SIN CUENTA {', '.join(self.faltantes)}"
            return f"hoja «{self.hoja}»  ·  {txt}"
        return ("cortes: " + ", ".join(etiqueta_corta(p) for p in self.cortes)
                + "\nhojas: " + self.hoja)


def sha256_archivo(ruta: "str | Path") -> str:
    """Huella del archivo, para reconocer una copia ya subida."""
    h = hashlib.sha256()
    with open(ruta, "rb") as f:
        for trozo in iter(lambda: f.read(1 << 20), b""):
            h.update(trozo)
    return h.hexdigest()


def leer_fuente(ruta: "str | Path", hojas: "Sequence | None" = None) -> Lectura:
    """Lee un Excel y devuelve la Lectura, detectando solo de cuál de los dos se trata."""
    ruta = Path(ruta)
    if hojas is None:
        hojas = leer_libro(ruta)

    bal = parse_balanza(hojas, ruta.name)
    if bal is not None:
        if not bal.periodo:
            raise ValueError(
                f"{ruta.name}: se leyó la balanza pero no se identificó el periodo. "
                "Renombra el archivo como Balanza_MMAAAA.xlsx."
            )
        return Lectura("balanza", ruta.name, bal.hoja,
                       {bal.periodo: {"local": dict(bal.valores), "cnsf": {}}},
                       bal.detalle, bal.faltantes, sha256_archivo(ruta))

    act = parse_actuarios(hojas)
    if not act.periodos:
        raise ValueError(
            f"{ruta.name}: no se reconoció ni como balanza (falta la columna CUENTA) "
            "ni como archivo de actuarios (faltan los tres conceptos de reserva)."
        )
    return Lectura("actuarios", ruta.name, ", ".join(act.hojas),
                   {p: {"local": dict(v["local"]), "cnsf": dict(v["cnsf"])}
                    for p, v in act.periodos.items()},
                   [], [], sha256_archivo(ruta))


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
        return self.fundir(leer_fuente(ruta, hojas))

    def fundir(self, lec: Lectura) -> "list[str]":
        """Mete una Lectura al histórico: actualiza su periodo y conserva los demás."""
        avisos: "list[str]" = []

        if lec.fuente == "balanza":
            p = lec.cortes[0]
            e = self._entrada(p)
            e["local"].update(lec.periodos[p]["local"])
            e["origen"]["local"] = f"Balanza · {lec.archivo}"
            self._revisar(e)
            avisos.append(
                f"{lec.archivo} → balanza al {etiqueta_periodo(p)} · " + " · ".join(lec.detalle)
                + (f" · SIN CUENTA {', '.join(lec.faltantes)}" if lec.faltantes else "")
            )
            return avisos

        for p in lec.cortes:
            e = self._entrada(p)
            src = lec.periodos[p]
            e["cnsf"].update(src["cnsf"])
            e["local_actuarios"] = src["local"]
            origen_local = e["origen"].get("local", "")
            if not origen_local or origen_local.startswith("Actuarios"):
                # la balanza manda para la columna local; esto solo rellena los
                # periodos que todavía no tienen balanza
                for cid, val in src["local"].items():
                    e["local"].setdefault(cid, val)
                if not origen_local:
                    e["origen"]["local"] = f"Actuarios · {lec.archivo}"
            e["origen"]["cnsf"] = f"Actuarios · {lec.archivo}"
            self._revisar(e)
        avisos.append(
            f"{lec.archivo} → actuarios, {len(lec.cortes)} corte(s): "
            + ", ".join(etiqueta_corta(p) for p in lec.cortes)
            + " · hojas: " + lec.hoja
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

    def fx(self, periodo: str, por_omision: float = FX_DEFAULT) -> float:
        """El tipo de cambio de ese cierre.

        Cada corte puede llevar el suyo —diciembre se convierte al de diciembre,
        no al de hoy—; el que no tenga usa el de la ventana.
        """
        v = self.datos.get(periodo, {}).get("fx")
        try:
            return float(v) if v else float(por_omision)
        except (TypeError, ValueError):
            return float(por_omision)

    def poner_fx(self, periodo: str, valor: "float | None") -> None:
        """Fija (o borra, con None) el tipo de cambio de ese cierre."""
        if periodo not in self.datos:
            return
        if valor:
            self.datos[periodo]["fx"] = float(valor)
        else:
            self.datos[periodo].pop("fx", None)

    def incremento(self, actual: str, previo: str, cid: "str | None" = None) -> "float | None":
        a, b = self.diferencia(actual, cid), self.diferencia(previo, cid)
        return None if a is None or b is None else a - b

    # -------------------------------------------------------------- mensajes
    def mensajes(self, actual: str, previo: "str | None", fx: "float | None" = None,
                 fx_previo: "float | None" = None) -> "list[str]":
        """Los tres mensajes clave, redactados con las cifras del propio corte.

        Con `fx` se redactan en pesos; sin él, en dólares. Cada corte se convierte
        con SU tipo de cambio, igual que los indicadores: si no fuera así, los
        mensajes dirían una variación y los KPIs otra.
        """
        sim = "USD" if fx is None else "MXN"
        ka = 1.0 if fx is None else float(fx)                       # cierre actual
        kp = 1.0 if fx is None else float(fx_previo or fx)          # cierre anterior

        def q(v: float) -> str:
            return f"{v / 1e6:,.2f}"

        d1_usd = self.diferencia(actual)
        if d1_usd is None:
            falta = "el archivo de actuarios" if self.total(actual, "cnsf") is None else "la balanza"
            return [
                f"Al {etiqueta_periodo(actual)} solo se ha cargado una de las dos fuentes: "
                f"falta {falta} de ese corte, así que todavía no hay diferencia que comparar.",
                "Los cortes anteriores del histórico se conservan intactos.",
            ]
        if previo is None or not self.completo(previo):
            return [
                f"Al {etiqueta_periodo(actual)}, la diferencia entre metodologías "
                f"asciende a {sim} {q(d1_usd * ka)} MM.",
                "Agrega un corte anterior completo al histórico para comparar la evolución "
                "del diferencial.",
            ]

        d0_usd = self.diferencia(previo) or 0.0
        d1, d0 = d1_usd * ka, d0_usd * kp
        v_local = self.total(actual, "local") * ka - self.total(previo, "local") * kp
        v_cnsf = self.total(actual, "cnsf") * ka - self.total(previo, "cnsf") * kp
        rango = f"Entre {etiqueta_corta(previo).lower()} y {etiqueta_corta(actual).lower()}"

        # el umbral se mide siempre en dólares, para que las dos versiones del
        # texto hablen exactamente de las mismas reservas
        def mueve(cid: str, lado: str) -> "tuple[float, float]":
            a = self.datos[actual][lado].get(cid, 0.0)
            b = self.datos[previo][lado].get(cid, 0.0)
            return a - b, a * ka - b * kp

        motor = max(((c, *mueve(c.id, "local")) for c in CONCEPTOS),
                    key=lambda x: abs(x[1]), default=None)
        m1 = (f"{rango}, las reservas bajo QES Metodología local "
              f"{'aumentan' if v_local >= 0 else 'disminuyen'} {sim} {q(abs(v_local))} MM")
        if motor and abs(motor[1]) > 5000:
            m1 += (f", principalmente por {'el incremento' if motor[2] >= 0 else 'la reducción'} "
                   f"de {sim} {q(abs(motor[2]))} MM en la "
                   f"{motor[0].label.replace('Reserva de ', 'reserva de ')}")
        m1 += "."

        sube, baja = [], []
        for c in CONCEPTOS:
            d_usd, d_disp = mueve(c.id, "cnsf")
            if abs(d_usd) > 5000:
                (sube if d_usd > 0 else baja).append((c, abs(d_disp)))

        def lista(arr):
            return " y de ".join(
                f"{sim} {q(d)} MM en {c.label.replace('Reserva de ', '').lower()}" for c, d in arr)

        m2 = (f"Bajo el Método Estatutario, las reservas totales "
              f"{'aumentan' if v_cnsf >= 0 else 'disminuyen'} {sim} {q(abs(v_cnsf))} MM.")
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
        m3 += (f" La diferencia total por constitución pasa de {sim} {q(d0)} MM en "
               f"{etiqueta_corta(previo).lower()} a {sim} {q(d1)} MM en "
               f"{etiqueta_corta(actual).lower()}, con "
               f"{'un incremento' if v >= 0 else 'una disminución'} de {sim} {q(abs(v))} MM")
        if d0:
            m3 += f" ({'+' if v >= 0 else '−'}{abs(v / d0 * 100):.1f}%)"
        m3 += "."
        if fx is not None and abs(ka - kp) > 1e-9:
            m3 += (f" En pesos la variación lleva dentro el efecto cambiario: "
                   f"{kp:,.4f} al cierre de {etiqueta_corta(previo).lower()} contra "
                   f"{ka:,.4f} al de {etiqueta_corta(actual).lower()}.")
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
# 6 bis. El servidor de auditoría: copias etiquetadas por mes y por usuario
# -----------------------------------------------------------------------------
#
# Cada archivo que se sube queda como una CARGA (quién, cuándo, qué archivo) con
# su DETALLE (importes por corte y concepto). No se pisa nada: subir otra vez el
# mismo mes deja una copia nueva y la anterior se conserva, así que siempre se
# puede volver a la que se usó en un cierre pasado.
#
#   dbo.ReservasQES_Cargas    carga_id · usuario · equipo · fecha · fuente ·
#                             archivo · hoja · sha256 · periodo_min/max · filas
#   dbo.ReservasQES_Detalle   carga_id · periodo · concepto · cuenta · local · cnsf
#
# La conexión es la misma de siempre (autenticación integrada de Windows):
#     DRIVER={ODBC Driver 17 for SQL Server};SERVER=...;DATABASE=...;Trusted_Connection=yes


def drivers_odbc() -> "list[str]":
    """Los drivers de SQL Server instalados en este equipo."""
    try:
        import pyodbc
    except ImportError:
        return []
    try:
        return [d for d in pyodbc.drivers() if "SQL Server" in d]
    except Exception:
        return []


def elegir_driver(preferido: str = DRIVER) -> str:
    """El mejor driver disponible.

    En una máquina puede estar el 18 y no el 17, o sólo el Native Client. Sin
    esto la conexión falla con un error que no dice nada útil.
    """
    disponibles = drivers_odbc()
    if not disponibles or preferido in disponibles:
        return preferido
    numerados = sorted(
        (d for d in disponibles if re.search(r"ODBC Driver (\d+)", d)),
        key=lambda d: int(re.search(r"ODBC Driver (\d+)", d).group(1)),
        reverse=True)
    return numerados[0] if numerados else disponibles[0]


def diagnostico() -> "list[str]":
    """Qué hay y qué falta en este equipo, antes de pelearse con el servidor."""
    out = [f"Python {sys.version.split()[0]}"]
    for mod, para in (("openpyxl", "leer .xlsx"), ("pyxlsb", "leer .xlsb"),
                      ("sqlalchemy", "hablar con el servidor"),
                      ("pyodbc", "el driver de SQL Server"),
                      ("tkinterdnd2", "arrastrar y soltar (opcional)")):
        try:
            __import__(mod)
            out.append(f"{mod}: sí ({para})")
        except ImportError:
            out.append(f"{mod}: FALTA · pip install {mod} — {para}")
    ds = drivers_odbc()
    out.append("drivers ODBC: " + (", ".join(ds) if ds else "ninguno encontrado"))
    if ds:
        out.append(f"se usará: {elegir_driver()}")
    return out


class Repositorio:
    """Lectura y escritura de las tablas en el servidor.

    Se apoya en SQLAlchemy, así que el mismo código sirve para SQL Server (lo
    normal) y para un archivo SQLite (útil para probar sin red: pasa
    `url="sqlite:///pruebas.db"`).
    """

    def __init__(self, servidor: str = SERVIDOR, base: str = BASE,
                 driver: "str | None" = None, url: "str | None" = None,
                 esquema: "str | None" = None):
        self.servidor, self.base = servidor, base
        self.driver = driver or elegir_driver()
        self.url = url
        self.engine = None
        self._esquema = esquema
        self._tablas = None

    # ------------------------------------------------------------- conexión
    def cadena(self) -> str:
        """La URL de SQLAlchemy. Sin `url` explícita, SQL Server con Trusted_Connection."""
        if self.url:
            return self.url
        params = urllib.parse.quote_plus(
            f"DRIVER={{{self.driver}}};"
            f"SERVER={self.servidor};"
            f"DATABASE={self.base};"
            "Trusted_Connection=yes;"
        )
        return f"mssql+pyodbc:///?odbc_connect={params}"

    @property
    def esquema(self) -> "str | None":
        if self._esquema is not None:
            return self._esquema or None
        return "dbo" if self.cadena().startswith("mssql") else None

    def conectar(self) -> str:
        """Abre la conexión y devuelve con qué servidor se habló. Lanza si no puede."""
        try:
            from sqlalchemy import create_engine, text
        except ImportError as err:
            raise RuntimeError(
                "Falta SQLAlchemy. En una celda: !pip install sqlalchemy pyodbc"
            ) from err
        try:
            self.engine = create_engine(self.cadena(), future=True)
            with self.engine.connect() as conn:
                if not self.cadena().startswith("mssql"):
                    return f"SQLite · {self.cadena()}"
                fila = conn.execute(text(
                    "SELECT @@SERVERNAME, DB_NAME(), SUSER_SNAME(), "
                    "HAS_PERMS_BY_NAME(NULL, NULL, 'CREATE TABLE')")).one()
                self.puede_crear = bool(fila[3])
                aviso = "" if fila[3] else " · SIN permiso para crear tablas en esta base"
                return f"{fila[0]} · base {fila[1]} · como {fila[2]} · driver {self.driver}{aviso}"
        except Exception as err:
            self.engine = None
            ds = drivers_odbc()
            pista = ""
            if not ds:
                pista = ("\n  No hay ningún driver ODBC de SQL Server instalado "
                         "(o falta pyodbc: pip install pyodbc).")
            elif self.driver not in ds:
                pista = f"\n  Drivers instalados aquí: {', '.join(ds)}."
            raise RuntimeError(f"{err}{pista}") from err

    def _asegura(self):
        if self.engine is None:
            raise RuntimeError("No hay conexión: pica «Conectar» primero.")
        return self.engine

    # -------------------------------------------------------------- esquema
    def tablas(self):
        """Define (sin crear) las dos tablas. Se memoriza para no rehacerlas."""
        if self._tablas is not None:
            return self._tablas
        from sqlalchemy import (Column, DateTime, ForeignKey, Integer, MetaData,
                                Numeric, PrimaryKeyConstraint, String, Table)
        md = MetaData(schema=self.esquema)
        cargas = Table(
            TABLA_CARGAS, md,
            Column("carga_id", Integer, primary_key=True, autoincrement=True),
            Column("usuario", String(128), nullable=False),
            Column("equipo", String(128)),
            Column("fecha_carga", DateTime, nullable=False),
            Column("fuente", String(16), nullable=False),        # balanza | actuarios
            Column("archivo", String(260), nullable=False),
            Column("hoja", String(200)),
            Column("sha256", String(64), nullable=False),
            Column("periodo_min", String(10)),
            Column("periodo_max", String(10)),
            Column("filas", Integer),
            Column("nota", String(400)),
        )
        detalle = Table(
            TABLA_DETALLE, md,
            Column("carga_id", Integer,
                   ForeignKey(f"{cargas.fullname}.carga_id"), nullable=False),
            Column("periodo", String(10), nullable=False),
            Column("concepto", String(8), nullable=False),
            Column("cuenta", String(8)),
            Column("etiqueta", String(80)),
            Column("local", Numeric(20, 6)),
            Column("cnsf", Numeric(20, 6)),
            PrimaryKeyConstraint("carga_id", "periodo", "concepto"),
        )
        self._tablas = (md, cargas, detalle)
        return self._tablas

    def existe_esquema(self) -> bool:
        """¿Están ya las dos tablas? La primera vez, claro que no."""
        from sqlalchemy import inspect
        hay = set(inspect(self._asegura()).get_table_names(schema=self.esquema))
        return {TABLA_CARGAS, TABLA_DETALLE} <= hay

    def crear_esquema(self) -> "list[str]":
        """Crea las tablas si no existen. Es seguro repetirlo."""
        from sqlalchemy import inspect
        eng = self._asegura()
        md, cargas, detalle = self.tablas()
        antes = set(inspect(eng).get_table_names(schema=self.esquema))
        md.create_all(eng, checkfirst=True)
        despues = set(inspect(eng).get_table_names(schema=self.esquema))
        nuevas = sorted(despues - antes)
        return nuevas or []

    # -------------------------------------------------------------- subidas
    def ya_subido(self, sha: str) -> "list[int]":
        """Cargas anteriores con la misma huella: el archivo ya está en el servidor."""
        from sqlalchemy import select
        eng = self._asegura()
        _, cargas, _ = self.tablas()
        with eng.connect() as conn:
            return [r[0] for r in conn.execute(
                select(cargas.c.carga_id).where(cargas.c.sha256 == sha))]

    def subir(self, lec: Lectura, usuario: "str | None" = None,
              nota: "str | None" = None) -> int:
        """Guarda la lectura como una carga nueva y devuelve su carga_id."""
        from sqlalchemy import insert
        eng = self._asegura()
        _, cargas, detalle = self.tablas()
        etiquetas = {c.id: (c.cuenta, c.label) for c in CONCEPTOS}

        filas = []
        for p in lec.cortes:
            src = lec.periodos[p]
            for cid in sorted(set(src["local"]) | set(src["cnsf"])):
                cuenta, label = etiquetas.get(cid, (None, cid))
                filas.append({"periodo": p, "concepto": cid, "cuenta": cuenta,
                              "etiqueta": label,
                              "local": src["local"].get(cid),
                              "cnsf": src["cnsf"].get(cid)})
        if not filas:
            raise ValueError(f"{lec.archivo}: no hay importes que subir.")

        cab = {
            "usuario": usuario or getpass.getuser(),
            "equipo": platform.node()[:128],
            "fecha_carga": dt.datetime.now(),
            "fuente": lec.fuente,
            "archivo": lec.archivo[:260],
            "hoja": (lec.hoja or "")[:200],
            "sha256": lec.sha256,
            "periodo_min": lec.cortes[0],
            "periodo_max": lec.cortes[-1],
            "filas": len(filas),
            "nota": (nota or "")[:400] or None,
        }
        with eng.begin() as conn:      # todo o nada: cabecera y detalle juntos
            carga_id = conn.execute(insert(cargas), cab).inserted_primary_key[0]
            conn.execute(insert(detalle),
                         [dict(f, carga_id=carga_id) for f in filas])
        return int(carga_id)

    # ------------------------------------------------------------ consultas
    def snapshots(self) -> "list[dict[str, Any]]":
        """Una fila por (carga, corte): lo que el usuario puede elegir para la vista.

        Sin tablas todavía devuelve vacío: es el estado normal la primera vez,
        no un error que haya que enseñarle a nadie.
        """
        from sqlalchemy import func, select
        eng = self._asegura()
        if not self.existe_esquema():
            return []
        _, cargas, detalle = self.tablas()
        q = (select(cargas.c.carga_id, detalle.c.periodo, cargas.c.usuario,
                    cargas.c.fecha_carga, cargas.c.fuente, cargas.c.archivo,
                    func.count().label("conceptos"))
             .select_from(cargas.join(detalle, cargas.c.carga_id == detalle.c.carga_id))
             .group_by(cargas.c.carga_id, detalle.c.periodo, cargas.c.usuario,
                       cargas.c.fecha_carga, cargas.c.fuente, cargas.c.archivo)
             .order_by(detalle.c.periodo.desc(), cargas.c.fecha_carga.desc()))
        with eng.connect() as conn:
            return [dict(r._mapping) for r in conn.execute(q)]

    def importes(self, carga_id: int, periodo: str) -> "dict[str, dict[str, float]]":
        """Los importes de un corte dentro de una carga."""
        from sqlalchemy import select
        eng = self._asegura()
        _, _, detalle = self.tablas()
        q = select(detalle.c.concepto, detalle.c.local, detalle.c.cnsf).where(
            (detalle.c.carga_id == carga_id) & (detalle.c.periodo == periodo))
        out = {"local": {}, "cnsf": {}}
        with eng.connect() as conn:
            for cid, loc, cnsf in conn.execute(q):
                if loc is not None:
                    out["local"][cid] = float(loc)
                if cnsf is not None:
                    out["cnsf"][cid] = float(cnsf)
        return out

    # --------------------------------------------------- armado del histórico
    def historico(self, seleccion: "Sequence[tuple[str, int]] | None" = None,
                  ruta_json: "str | Path" = HIST_JSON) -> Historico:
        """Arma un Historico con lo que hay en el servidor.

        `seleccion` son los pares (periodo, carga_id) que el usuario eligió para
        las ranuras de la vista; sin ella se toma, de cada corte, la carga más
        reciente. La columna local sale de la carga elegida (una balanza, si la
        hay) y la estatutaria de la carga de actuarios más nueva de ese corte.
        """
        snaps = self.snapshots()
        h = Historico(ruta_json)
        h.datos = {}
        if not snaps:
            return h

        # la carga de actuarios más nueva por corte: de ahí sale siempre la CNSF
        act = {}
        for s in snaps:
            if s["fuente"] == "actuarios" and s["periodo"] not in act:
                act[s["periodo"]] = s

        if seleccion is None:
            elegidas = {}
            for s in snaps:                       # snapshots ya viene de más nueva a más vieja
                elegidas.setdefault(s["periodo"], s)
            pares = [(p, elegidas[p]["carga_id"]) for p in sorted(elegidas)]
        else:
            pares = [(p, cid) for p, cid in seleccion if p]

        por_id = {(s["carga_id"], s["periodo"]): s for s in snaps}
        for periodo, carga_id in pares:
            e = h._entrada(periodo)
            s = por_id.get((carga_id, periodo))
            if s is not None:
                vals = self.importes(carga_id, periodo)
                e["local"].update(vals["local"])
                e["cnsf"].update(vals["cnsf"])
                e["origen"]["local"] = (f"{s['fuente'].capitalize()} · {s['archivo']} "
                                        f"· {s['usuario']} · #{carga_id}")
                if vals["cnsf"]:
                    e["origen"]["cnsf"] = e["origen"]["local"]
            a = act.get(periodo)
            if a is not None and a["carga_id"] != carga_id:
                vals = self.importes(a["carga_id"], periodo)
                e["cnsf"].update(vals["cnsf"])
                e["local_actuarios"] = vals["local"]
                e["origen"]["cnsf"] = (f"Actuarios · {a['archivo']} · {a['usuario']} "
                                       f"· #{a['carga_id']}")
                for cid, val in vals["local"].items():
                    e["local"].setdefault(cid, val)
            elif a is not None:
                e["local_actuarios"] = self.importes(carga_id, periodo)["local"]
            h._revisar(e)
        return h


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

/* ---- el interruptor de moneda: sin una línea de JavaScript ---- */
input.sw{position:absolute;width:0;height:0;opacity:0;pointer-events:none}
.v-mxn{display:none}
#m-mxn:checked ~ .wrap .v-usd{display:none}
#m-mxn:checked ~ .wrap .v-mxn{display:inline}
#m-mxn:checked ~ .wrap text.v-mxn{display:inline}
.switch{display:inline-flex;border:1px solid var(--line);border-radius:5px;overflow:hidden;
        background:var(--paper);margin-top:10px}
.switch label{padding:6px 18px;font-size:12.5px;font-weight:600;color:var(--ink-2);
              cursor:pointer;user-select:none;border-right:1px solid var(--line);line-height:1.2}
.switch label:last-child{border-right:0}
.switch label:hover{background:var(--ice-2)}
#m-usd:checked ~ .wrap label[for=m-usd],
#m-mxn:checked ~ .wrap label[for=m-mxn]{background:var(--teal);color:#fff}
#m-usd:focus-visible ~ .wrap label[for=m-usd],
#m-mxn:focus-visible ~ .wrap label[for=m-mxn]{outline:2px solid var(--teal-2);outline-offset:-2px}
.switch-note{font-size:11.5px;color:var(--ink-3);margin-top:5px;font-family:var(--mono)}

@media print{body{background:#fff}.board{border:0;box-shadow:none}
             footer.credits,.switch{display:none}}
""" % {
    "plum": PLUM, "plum_deep": PLUM_DEEP, "plum_soft": PLUM_SOFT,
    "teal": TEAL, "teal2": TEAL_2, "teal_soft": TEAL_SOFT,
    "ice": ICE, "ice2": ICE_2, "paper": PAPER, "ground": GROUND,
    "ink": INK, "ink2": INK_2, "ink3": INK_3,
    "line": LINE, "line_soft": LINE_SOFT, "neg": NEG, "warn": WARN,
}


def _svg_cascada(hist: Historico, actual: str, previo: "str | None",
                 fx_a: float = FX_DEFAULT, fx_p: float = FX_DEFAULT) -> str:
    """Gráfico de cascada: corte anterior, incrementos por reserva y corte actual.

    Las barras no cambian al cambiar de moneda —la escala es proporcional—, así
    que solo se guardan las dos versiones de cada número.
    """
    if not previo or not hist.completo(actual) or not hist.completo(previo):
        return ('<p class="chart-note">Hacen falta dos cortes con las dos fuentes cargadas '
                'para dibujar la cascada.</p>')

    W, H, L, R, T, B = 660, 330, 46, 16, 46, 74
    d0 = (hist.diferencia(previo) or 0.0) / 1e6
    d1 = (hist.diferencia(actual) or 0.0) / 1e6

    etq_id = {c.label.replace("Reserva de ", ""): c.id for c in CONCEPTOS}
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
        # el mismo número en pesos: las bases al cambio de su corte, y el
        # incremento por reserva como la resta de los dos cierres convertidos
        if tipo == "base":
            mxn = valor * (fx_a if i == len(geo) - 1 else fx_p)
        else:
            a = hist.diferencia(actual, etq_id.get(etq)) or 0.0
            b = hist.diferencia(previo, etq_id.get(etq)) or 0.0
            mxn = (a * fx_a - b * fx_p) / 1e6
        for clase, num in (("v-usd", valor), ("v-mxn", mxn)):
            partes.append(f'<text class="{clase}" x="{cx:.1f}" y="{ya - 9:.1f}" '
                          f'text-anchor="middle" fill="{PLUM if tipo == "base" else TEAL}" '
                          f'font-family="Consolas,monospace" font-size="13" '
                          f'font-weight="500">{signo}{abs(num):,.2f}</text>')

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

    for clase, unidad in (("v-usd", "MM USD"), ("v-mxn", "MM MXN")):
        partes.append(f'<text class="{clase}" x="{L - 8}" y="{T - 18}" fill="{INK_3}" '
                      f'font-family="Consolas,monospace" font-size="11">'
                      f'Diferencia acumulada ({unidad})</text>')
    partes.append("</svg>")
    return "\n".join(partes)


def construir_html(hist: Historico, periodos: "Sequence[str] | None" = None,
                   fx: float = FX_DEFAULT, nota: str = NOTA_RELEVANTE) -> str:
    """Arma la vista completa en un solo archivo HTML, sin dependencias externas.

    Las cifras van en las dos monedas y el interruptor de arriba decide cuál se
    ve. Cada corte se convierte con SU tipo de cambio de cierre: `fx` es el que
    usan los cortes que no traigan el suyo.
    """
    ps = list(periodos) if periodos else hist.periodos()[-PERIODOS_EN_VISTA:]
    if not ps:
        raise ValueError("El histórico está vacío: no hay nada que dibujar.")
    actual = ps[-1]
    previo = ps[-2] if len(ps) > 1 else None
    tc = {p: hist.fx(p, fx) for p in ps}          # el tipo de cambio de cada cierre
    fx_a = tc[actual]

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
                f'<div class="k-value">{dual_texto("{sim} {n}", d1, fx_a)} '
                f'<span class="u">MM</span></div>'
                f'<div class="k-alt">'
                + dual(f"~MXN {d1 * fx_a / 1e6:,.2f} MM al cierre",
                       f"USD {d1 / 1e6:,.2f} MM · {fx_a:,.4f} MXN/USD")
                + '</div></div>']
    if previo and hist.completo(actual) and hist.completo(previo):
        d0 = hist.diferencia(previo) or 0.0
        fx_p = tc[previo]
        v = d1 - d0
        # en pesos, el movimiento se mide con el cierre de cada mes a su propio
        # tipo de cambio: así lleva dentro el efecto cambiario, como debe ser
        v_mxn = d1 * fx_a - d0 * fx_p
        alcance = f"{etiqueta_corta(previo)} → {etiqueta_corta(actual)}"
        kpis.append(f'<div class="kpi accent"><div class="k-label">Variación del periodo</div>'
                    f'<div class="k-scope">{esc(alcance)}</div>'
                    f'<div class="k-value">'
                    + dual(f'{"+" if v >= 0 else "−"}USD {abs(v) / 1e6:,.2f}',
                           f'{"+" if v_mxn >= 0 else "−"}MXN {abs(v_mxn) / 1e6:,.2f}')
                    + ' <span class="u">MM</span></div>'
                    f'<div class="k-alt">'
                    + dual(f"~MXN {abs(v_mxn) / 1e6:,.2f} MM",
                           "incluye el efecto cambiario")
                    + '</div></div>')
        pct = f'{"+" if v >= 0 else "−"}{abs(v / d0 * 100):.1f}%' if d0 else "n/d"
        pct_mxn = (f'{"+" if v_mxn >= 0 else "−"}{abs(v_mxn / (d0 * fx_p) * 100):.1f}%'
                   if d0 * fx_p else "n/d")
        kpis.append(f'<div class="kpi"><div class="k-label">Variación % de la diferencia</div>'
                    f'<div class="k-scope">{esc(alcance)}</div>'
                    f'<div class="k-value">{dual(pct, pct_mxn)}</div>'
                    f'<div class="k-alt">'
                    + dual(f"sobre USD {d0 / 1e6:,.2f} MM",
                           f"sobre MXN {d0 * fx_p / 1e6:,.2f} MM")
                    + '</div></div>')

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

    def inc_dual(a: "float | None", b: "float | None") -> str:
        """El incremento: en dólares es la resta; en pesos, cada corte a su cambio."""
        if a is None or b is None:
            return dual("—", "—")
        return dual(celda_mm(a - b), celda_mm(a * fx_a - b * tc[previo]))

    cuerpo = ""
    for c in CONCEPTOS:
        cuerpo += f"<tr><th>{esc(c.label)}{' *' if c.nota else ''}</th>"
        for p in ps:
            e = hist.datos[p]
            cuerpo += (f'<td class="gstart">{dual_mm(e["local"].get(c.id), tc[p])}</td>'
                       f'<td>{dual_mm(e["cnsf"].get(c.id), tc[p])}</td>'
                       f'<td class="dif">{dual_mm(hist.diferencia(p, c.id), tc[p])}</td>')
        if previo:
            cuerpo += ('<td class="delta">'
                       + inc_dual(hist.diferencia(actual, c.id),
                                  hist.diferencia(previo, c.id)) + "</td>")
        cuerpo += "</tr>"

    cuerpo += '<tr class="total"><th>Total reservas</th>'
    for p in ps:
        cuerpo += (f'<td class="gstart">{dual_mm(hist.total(p, "local"), tc[p])}</td>'
                   f'<td>{dual_mm(hist.total(p, "cnsf"), tc[p])}</td>'
                   f'<td class="dif">{dual_mm(hist.diferencia(p), tc[p])}</td>')
    if previo:
        cuerpo += ('<td class="delta">'
                   + inc_dual(hist.diferencia(actual), hist.diferencia(previo)) + "</td>")
    cuerpo += "</tr>"

    # ------------------------------------------------------- mensajes y listas
    mensajes = "".join(
        f'<li>{dual(esc(u), esc(m))}</li>'
        for u, m in zip(hist.mensajes(actual, previo),
                        hist.mensajes(actual, previo, fx=fx_a,
                                      fx_previo=tc.get(previo, fx_a))))

    por_reserva = ""
    for c in CONCEPTOS:
        d = hist.diferencia(actual, c.id)
        if d is None or abs(d) < 5000:
            continue
        por_reserva += (f"<li><span>{esc(c.label)}</span>"
                        f'<span class="v">{dual_texto("{sim} {n} MM", d, fx_a)}'
                        f"<small>"
                        + dual(f"~MXN {d * fx_a / 1e6:,.2f} MM", f"USD {d / 1e6:,.2f} MM")
                        + "</small></span></li>")
    if dif_actual is None:
        por_reserva = ('<li><span>Este corte todavía no tiene las dos fuentes cargadas.</span>'
                       '<span class="v">n/d</span></li>')
    else:
        por_reserva += ('<li><span>Total de reservas<br />'
                        '<small>incremento por constitución</small></span>'
                        f'<span class="v">{dual_texto("{sim} {n} MM", d1, fx_a)}'
                        f"<small>"
                        + dual(f"~MXN {d1 * fx_a / 1e6:,.2f} MM", f"USD {d1 / 1e6:,.2f} MM")
                        + "</small></span></li>")

    avisos = [f"{etiqueta_periodo(p)}: {hist.datos[p]['aviso']}"
              for p in ps if hist.datos[p].get("aviso")]
    banda_avisos = ""
    if avisos:
        banda_avisos = ('<div class="flags"><strong>Revisar la metodología local</strong>'
                        "<ul>" + "".join(f"<li>{esc(a)}</li>" for a in avisos) + "</ul></div>")

    origen = hist.datos[actual].get("origen", {})
    pie = " · ".join(filter(None, [origen.get("local"), origen.get("cnsf")])) or "—"
    sello = dt.datetime.now().strftime("%d/%m/%Y %H:%M")

    # el tipo de cambio que se enseña arriba: uno solo si todos coinciden
    distintos = len({round(v, 6) for v in tc.values()}) > 1
    if distintos:
        linea_fx = ("Tipo de cambio de cierre · "
                    + " · ".join(f"{etiqueta_corta(p)} {tc[p]:,.4f}" for p in ps)
                    + " MXN/USD")
    else:
        linea_fx = f"Tipo de cambio: {fx_a:,.4f} MXN / USD"

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Reservas técnicas QES · {esc(etiqueta_corta(actual))}</title>
<style>{CSS}</style>
</head>
<body>
<input type="radio" name="moneda" id="m-usd" class="sw" checked />
<input type="radio" name="moneda" id="m-mxn" class="sw" />
<div class="wrap">
{banda_avisos}
  <section class="board">
    <div class="board-head">
      <div>
        <h2>Resultados reservas técnicas QES</h2>
        <p class="tag">Comparativo de metodologías y evolución del diferencial</p>
        <p class="fx">{esc(linea_fx)}</p>
        <div class="switch" role="group" aria-label="Moneda">
          <label for="m-usd">Dólares</label><label for="m-mxn">Pesos</label>
        </div>
        <p class="switch-note">{
            "Cada corte se convierte con su propio tipo de cambio de cierre."
            if distintos else "Cifras en millones. Pica para cambiar de moneda."}</p>
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
      <span>{dual("Cifras en millones de USD.", "Cifras en millones de MXN.")}
            * {esc(NOTA_PIE)}</span>
      <span>La diferencia se presenta como Método Estatutario CNSF menos Metodología local,
            es decir el exceso de constitución del método estatutario.</span>
    </div>

    <div class="band">
      <div class="panel">
        <h3>Evolución de la diferencia entre metodologías</h3>
        <p class="chart-note">{dual("Millones de USD", "Millones de MXN")}{
            " · " + esc(etiqueta_corta(previo)) + " → " + esc(etiqueta_corta(actual)) if previo else ""}</p>
        {_svg_cascada(hist, actual, previo, fx_a, tc.get(previo, fx_a))}
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


# -----------------------------------------------------------------------------
# 7 bis. La evolución de la diferencia mes con mes (el botón «Ver evolución»)
# -----------------------------------------------------------------------------

# Colores de serie: salen de la paleta de la casa, pero elegidos entre los pasos
# que separan bien para daltonismo (vino oscuro contra azul claro, ΔE 25 en
# deuteranopia). Aun así van con leyenda y etiqueta directa: el color nunca es
# la única pista de qué es cada cosa.
SERIE_COLOR = {"rrc": "#3E8AA6", "rsnr": "#5B1A44", "rsr": "#16252D"}


def _paso_bonito(v: float) -> float:
    """Un escalón de rejilla que caiga en números redondos."""
    if v <= 0:
        return 1.0
    import math
    exp = math.floor(math.log10(v))
    base = v / (10 ** exp)
    for corte in (1, 2, 2.5, 5, 10):
        if base <= corte:
            return corte * (10 ** exp)
    return 10 ** (exp + 1)


def _svg_evolucion(hist: Historico, periodos: "Sequence[str]", fx: float) -> str:  # noqa: C901
    """Columnas apiladas: la diferencia de cada corte, abierta por reserva.

    Las columnas se dibujan primero y los globos del cursor después, todos al
    final: en SVG no hay z-index, así que lo único que decide qué tapa a qué es
    el orden. Cada globo se enciende con el `~` de CSS desde su columna.
    """
    ps = [p for p in periodos if hist.completo(p)]
    if len(ps) < 2:
        return '<p class="chart-note">Hacen falta al menos dos cortes completos.</p>'

    # solo las reservas que mueven la aguja en algún mes
    series = [c for c in CONCEPTOS
              if any(abs(hist.diferencia(p, c.id) or 0) >= 5000 for p in ps)]
    totales = [(hist.diferencia(p) or 0.0) / 1e6 for p in ps]
    tope = (max(totales) * 1.18) or 1.0
    paso = _paso_bonito(tope / 4)

    W, H = 980, 430
    L, R, T, B = 62, 22, 34, 74
    alto_util = H - T - B

    def y(v: float) -> float:
        return T + alto_util * (1 - v / tope)

    hueco = (W - L - R) / len(ps)
    bw = min(64.0, hueco * 0.62)

    reglas = "\n".join(f".ev .c{i}:hover ~ .t{i}{{opacity:1}}" for i in range(len(ps)))
    o = [f'<svg class="ev" viewBox="0 0 {W} {H}" role="img" '
         f'aria-label="Diferencia entre metodologías por corte, en millones de USD">',
         f'<style>{reglas}</style>']

    # rejilla recesiva, detrás de todo
    n = 0
    while n * paso <= tope:
        yy = y(n * paso)
        o.append(f'<line x1="{L}" y1="{yy:.1f}" x2="{W - R}" y2="{yy:.1f}" '
                 f'stroke="{LINE_SOFT}" stroke-width="1" />')
        o.append(f'<text x="{L - 10}" y="{yy + 4:.1f}" text-anchor="end" fill="{INK_3}" '
                 f'font-family="Consolas,monospace" font-size="10.5">{n * paso:,.2f}</text>')
        n += 1
    o.append(f'<line x1="{L}" y1="{y(0):.1f}" x2="{W - R}" y2="{y(0):.1f}" '
             f'stroke="{LINE}" stroke-width="1.5" />')
    for clase, txt in (("v-usd", "MM USD"), ("v-mxn", "MM MXN")):
        o.append(f'<text class="{clase}" x="{L - 10}" y="{T - 14}" text-anchor="end" '
                 f'fill="{INK_3}" font-family="Consolas,monospace" '
                 f'font-size="10.5">{txt}</text>')

    # --- las columnas ---------------------------------------------------
    desglose = []
    for i, p in enumerate(ps):
        cx = L + hueco * i + hueco / 2
        acum, trozos = 0.0, []
        for c in series:
            v = (hist.diferencia(p, c.id) or 0.0) / 1e6
            if abs(v) < 0.005:
                continue
            trozos.append((c, v, acum))
            acum += v
        desglose.append((cx, trozos))

        o.append(f'<g class="col c{i}">')
        # el blanco del cursor es toda la banda de la columna, no solo la barra
        o.append(f'<rect class="hit" x="{cx - hueco / 2:.1f}" y="{T}" '
                 f'width="{hueco:.1f}" height="{alto_util:.1f}" />')
        for k, (c, v, base) in enumerate(trozos):
            y0, y1 = y(base + v), y(base)
            arriba = k == len(trozos) - 1
            alto = max(2.0, y1 - y0 - (0 if arriba else 2))   # 2px de aire entre segmentos
            o.append(f'<rect class="seg" x="{cx - bw / 2:.1f}" y="{y0:.1f}" '
                     f'width="{bw:.1f}" height="{alto:.1f}" '
                     f'fill="{SERIE_COLOR.get(c.id, TEAL)}" rx="{4 if arriba else 0}" />')
        tot = totales[i]
        for clase, num in (("v-usd", tot), ("v-mxn", tot * hist.fx(p, fx))):
            o.append(f'<text class="{clase}" x="{cx:.1f}" y="{y(tot) - 10:.1f}" '
                     f'text-anchor="middle" fill="{INK}" font-family="Consolas,monospace" '
                     f'font-size="12" font-weight="500">{num:,.2f}</text>')
        o.append(f'<text x="{cx:.1f}" y="{H - B + 20}" text-anchor="middle" fill="{INK_2}" '
                 f'font-family="Segoe UI,sans-serif" font-size="11">'
                 f'{esc(etiqueta_corta(p).split(" ")[0][:3])}</text>')
        o.append(f'<text x="{cx:.1f}" y="{H - B + 34}" text-anchor="middle" fill="{INK_3}" '
                 f'font-family="Segoe UI,sans-serif" font-size="10">'
                 f'{esc(p.split("-")[0])}</text>')
        o.append("</g>")

    # --- los globos, hasta el final para que nada los tape ---------------
    for i, p in enumerate(ps):
        cx, trozos = desglose[i]
        tot = totales[i]
        ancho_t, alto_t = 224, 46 + 15 * len(trozos)
        tx = min(max(cx - ancho_t / 2, L), W - R - ancho_t)
        ty = max(y(tot) - alto_t - 20, 2)
        o.append(f'<g class="tip t{i}">')
        o.append(f'<rect x="{tx:.1f}" y="{ty:.1f}" width="{ancho_t}" height="{alto_t}" '
                 f'rx="5" fill="{INK}" opacity="0.97" />')
        o.append(f'<text x="{tx + 11:.1f}" y="{ty + 19:.1f}" fill="#fff" '
                 f'font-family="Segoe UI,sans-serif" font-size="11.5" font-weight="600">'
                 f'{esc(etiqueta_periodo(p))}</text>')
        for k, (c, v, _b) in enumerate(trozos):
            yy = ty + 36 + 15 * k
            o.append(f'<rect x="{tx + 11:.1f}" y="{yy - 8:.1f}" width="8" height="8" rx="2" '
                     f'fill="{SERIE_COLOR.get(c.id, TEAL)}" />')
            o.append(f'<text x="{tx + 25:.1f}" y="{yy:.1f}" fill="#D7E3EA" '
                     f'font-family="Segoe UI,sans-serif" font-size="10.5">'
                     f'{esc(c.label.replace("Reserva de ", ""))}</text>')
            for clase, num in (("v-usd", v), ("v-mxn", v * hist.fx(p, fx))):
                o.append(f'<text class="{clase}" x="{tx + ancho_t - 11:.1f}" y="{yy:.1f}" '
                         f'text-anchor="end" fill="#fff" font-family="Consolas,monospace" '
                         f'font-size="10.5">{num:,.2f}</text>')
        yy = ty + 36 + 15 * len(trozos)
        o.append(f'<text x="{tx + 11:.1f}" y="{yy:.1f}" fill="#9FB6C2" '
                 f'font-family="Segoe UI,sans-serif" font-size="10.5">Total</text>')
        fxp = hist.fx(p, fx)
        for clase, txt in (("v-usd", f"USD {tot:,.2f}  ·  MXN {tot * fxp:,.2f}"),
                           ("v-mxn", f"MXN {tot * fxp:,.2f}  ·  {fxp:,.4f} MXN/USD")):
            o.append(f'<text class="{clase}" x="{tx + ancho_t - 11:.1f}" y="{yy:.1f}" '
                     f'text-anchor="end" fill="#fff" font-family="Consolas,monospace" '
                     f'font-size="11" font-weight="600">{txt}</text>')
        o.append("</g>")

    o.append("</svg>")
    return "\n".join(o)


def construir_html_evolucion(hist: Historico, periodos: "Sequence[str] | None" = None,
                             fx: float = FX_DEFAULT) -> str:
    """La página de la evolución: el gráfico más la tabla que lo respalda."""
    ps = list(periodos) if periodos else hist.periodos()
    completos = [p for p in ps if hist.completo(p)]
    if not completos:
        raise ValueError("No hay ningún corte con las dos fuentes cargadas.")

    series = [c for c in CONCEPTOS
              if any(abs(hist.diferencia(p, c.id) or 0) >= 5000 for p in completos)]
    leyenda = "".join(
        f'<span class="chip"><i style="background:{SERIE_COLOR.get(c.id, TEAL)}"></i>'
        f'{esc(c.label.replace("Reserva de ", ""))}</span>' for c in series)

    tc = {p: hist.fx(p, fx) for p in completos}
    enc = "".join(f"<th>{esc(etiqueta_corta(p))}</th>" for p in completos)
    filas = ""
    for c in series:
        filas += f"<tr><th>{esc(c.label)}</th>" + "".join(
            f"<td>{dual_mm(hist.diferencia(p, c.id), tc[p])}</td>"
            for p in completos) + "</tr>"
    filas += ('<tr class="total"><th>Diferencia total</th>'
              + "".join(f"<td>{dual_mm(hist.diferencia(p), tc[p])}</td>" for p in completos)
              + "</tr>")

    def var_dual(i: int, p: str) -> str:
        if i == 0:
            return dual("—", "—")
        q = completos[i - 1]
        a, b = hist.diferencia(p) or 0.0, hist.diferencia(q) or 0.0
        return dual(celda_mm(a - b), celda_mm(a * tc[p] - b * tc[q]))

    filas += ('<tr class="var"><th>Variación contra el corte anterior</th>'
              + "".join(f"<td>{var_dual(i, p)}</td>" for i, p in enumerate(completos))
              + "</tr>")

    d_ini, d_fin = hist.diferencia(completos[0]) or 0.0, hist.diferencia(completos[-1]) or 0.0
    fx_ini, fx_fin = tc[completos[0]], tc[completos[-1]]
    v = d_fin - d_ini
    v_mxn = d_fin * fx_fin - d_ini * fx_ini
    pct = f'{"+" if v >= 0 else "−"}{abs(v / d_ini * 100):.1f}%' if d_ini else "n/d"
    pct_mxn = (f'{"+" if v_mxn >= 0 else "−"}{abs(v_mxn / (d_ini * fx_ini) * 100):.1f}%'
               if d_ini * fx_ini else "n/d")
    sello = dt.datetime.now().strftime("%d/%m/%Y %H:%M")
    omitidos = [p for p in ps if p not in completos]

    distintos = len({round(v_, 6) for v_ in tc.values()}) > 1
    linea_fx = (("Tipo de cambio de cierre · "
                 + " · ".join(f"{etiqueta_corta(p)} {tc[p]:,.4f}" for p in completos)
                 + " MXN/USD") if distintos
                else f"Tipo de cambio: {fx_fin:,.4f} MXN / USD")

    css_extra = """
.ev{width:100%;max-width:980px;height:auto;display:block}
.ev .hit{fill:transparent}
.ev .tip{opacity:0;pointer-events:none;transition:opacity .12s}
.ev .col:hover .tip{opacity:1}
.ev .col:hover .seg{filter:brightness(1.12)}
.legend{display:flex;flex-wrap:wrap;gap:16px;margin:2px 0 6px}
.chip{display:inline-flex;align-items:center;gap:7px;font-size:12.5px;color:var(--ink-2)}
.chip i{width:11px;height:11px;border-radius:3px;display:inline-block}
table.ev-t{border-collapse:collapse;width:100%;min-width:640px;
           font-variant-numeric:tabular-nums;margin-top:4px}
table.ev-t th,table.ev-t td{padding:7px 11px;font-size:12.5px;
                            border-bottom:1px solid var(--line-soft);text-align:right}
table.ev-t thead th{background:var(--ice);color:var(--teal);font-size:11px;
                    text-transform:uppercase;letter-spacing:.05em}
table.ev-t tbody th{text-align:left;font-weight:500;color:var(--ink)}
table.ev-t tbody td{font-family:var(--mono);color:var(--ink)}
table.ev-t tr.total th,table.ev-t tr.total td{background:var(--teal);color:#fff;font-weight:700}
table.ev-t tr.var td{color:var(--ink-2);font-size:12px}
table.ev-t tr.var th{font-weight:400;color:var(--ink-2);font-size:12px}
"""

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Evolución de la diferencia · Reservas QES</title>
<style>{CSS}{css_extra}</style>
</head>
<body>
<input type="radio" name="moneda" id="m-usd" class="sw" checked />
<input type="radio" name="moneda" id="m-mxn" class="sw" />
<div class="wrap">
  <section class="board">
    <div class="board-head">
      <div>
        <h2>Evolución de la diferencia</h2>
        <p class="tag">Método Estatutario CNSF menos Metodología local, corte a corte</p>
        <p class="fx">{esc(linea_fx)}</p>
        <div class="switch" role="group" aria-label="Moneda">
          <label for="m-usd">Dólares</label><label for="m-mxn">Pesos</label>
        </div>
      </div>
      <div class="kpis">
        <div class="kpi"><div class="k-label">Diferencia al último corte</div>
          <div class="k-scope">{esc(etiqueta_corta(completos[-1]))}</div>
          <div class="k-value">{dual_texto("{sim} {n}", d_fin, fx_fin)} <span class="u">MM</span></div>
          <div class="k-alt">{dual(f"~MXN {d_fin * fx_fin / 1e6:,.2f} MM",
                                   f"al cierre · {fx_fin:,.4f} MXN/USD")}</div></div>
        <div class="kpi accent"><div class="k-label">Desde el primer corte</div>
          <div class="k-scope">{esc(etiqueta_corta(completos[0]))} → {esc(etiqueta_corta(completos[-1]))}</div>
          <div class="k-value">{dual(f'{"+" if v >= 0 else "−"}USD {abs(v) / 1e6:,.2f}',
                                     f'{"+" if v_mxn >= 0 else "−"}MXN {abs(v_mxn) / 1e6:,.2f}')} <span class="u">MM</span></div>
          <div class="k-alt">{dual(f"{pct} sobre USD {d_ini / 1e6:,.2f} MM",
                                   f"{pct_mxn} sobre MXN {d_ini * fx_ini / 1e6:,.2f} MM")}</div></div>
        <div class="kpi"><div class="k-label">Cortes graficados</div>
          <div class="k-scope">con las dos fuentes cargadas</div>
          <div class="k-value">{len(completos)}</div>
          <div class="k-alt">de {len(ps)} en el histórico</div></div>
      </div>
    </div>

    <div class="band" style="grid-template-columns:1fr">
      <div class="panel">
        <h3>Diferencia por corte, abierta por reserva</h3>
        <div class="legend">{leyenda}</div>
        {_svg_evolucion(hist, completos, fx)}
        <p class="chart-note">{dual("Millones de USD", "Millones de MXN")} · pasa el
           cursor por una columna para ver el desglose{'' if not omitidos else ' · sin graficar: ' + esc(', '.join(etiqueta_corta(p) for p in omitidos)) + ' (falta una de las dos fuentes)'}</p>
      </div>
    </div>

    <div class="band" style="grid-template-columns:1fr">
      <div class="panel">
        <h3>Las mismas cifras, en tabla</h3>
        <div class="table-wrap">
          <table class="ev-t">
            <thead><tr><th style="text-align:left">Reserva</th>{enc}</tr></thead>
            <tbody>{filas}</tbody>
          </table>
        </div>
        <p class="chart-note">{dual("Millones de USD.", "Millones de MXN.")}
           El guion marca los cortes sin diferencia.</p>
      </div>
    </div>
  </section>
  <footer class="credits">Reservas técnicas QES · armado en local el {sello}</footer>
</div>
</body>
</html>
"""


def escribir_evolucion(hist: Historico, destino: "str | Path | None" = None,
                       periodos: "Sequence[str] | None" = None,
                       fx: float = FX_DEFAULT, abrir: bool = True) -> Path:
    """Escribe y abre la página de la evolución de la diferencia."""
    ruta = Path(destino) if destino else hist.ruta.parent / "evolucion_diferencia.html"
    ruta.write_text(construir_html_evolucion(hist, periodos, fx), encoding="utf-8")
    if abrir:
        try:
            webbrowser.open_new_tab(ruta.resolve().as_uri())
        except Exception:
            pass
    return ruta


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


def abrir_ventana(hist_ruta: "str | Path" = HIST_JSON, bloquear: bool = True,
                  url_bd: "str | None" = None):
    """Abre la ventana de trabajo. Es lo que corre la última línea del bloque.

    De arriba abajo: se cargan los dos Excel, se ven cuáles se leyeron, se suben
    al servidor de auditoría, se eligen las tres tablas de la vista
    (Diciembre · t-1 · t) y se arma el HTML.

    `url_bd` sirve para apuntar a un SQLite en vez de a SQL Server, para probar
    sin red: abrir_ventana(url_bd="sqlite:///pruebas.db").
    """
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

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
    ancho = max(600, min(1060, pw - 80))
    alto = max(460, min(760, ph - 110))
    root.geometry(f"{ancho}x{alto}+{max(0, (pw - ancho) // 2)}+{max(0, (ph - alto) // 4)}")
    root.minsize(560, 420)

    hist = Historico(hist_ruta)
    repo = Repositorio(url=url_bd)
    pendientes: "dict[str, Lectura]" = {}
    estado: "dict[str, Any]" = {"conectado": False, "snapshots": [], "nota": NOTA_RELEVANTE}
    ancho_texto = ancho - 90

    # ---------------------------------------------------------- lienzo con barra
    lienzo = tk.Canvas(root, bg=GROUND, highlightthickness=0)
    barra_v = ttk.Scrollbar(root, orient="vertical", command=lienzo.yview)
    cuerpo = tk.Frame(lienzo, bg=GROUND)
    cuerpo.bind("<Configure>", lambda e: lienzo.configure(scrollregion=lienzo.bbox("all")))
    ventana_id = lienzo.create_window((0, 0), window=cuerpo, anchor="nw")
    lienzo.bind("<Configure>", lambda e: lienzo.itemconfigure(ventana_id, width=e.width))
    lienzo.configure(yscrollcommand=barra_v.set)
    lienzo.pack(side="left", fill="both", expand=True)
    barra_v.pack(side="right", fill="y")
    root.bind_all("<MouseWheel>", lambda e: lienzo.yview_scroll(int(-e.delta / 120), "units"))
    root.bind_all("<Button-4>", lambda e: lienzo.yview_scroll(-1, "units"))
    root.bind_all("<Button-5>", lambda e: lienzo.yview_scroll(1, "units"))

    def tarjeta(titulo: "str | None" = None) -> tk.Frame:
        marco = tk.Frame(cuerpo, bg=PAPER, highlightbackground=LINE, highlightthickness=1)
        marco.pack(fill="x", padx=14, pady=6)
        if titulo:
            tk.Label(marco, text=titulo, bg=PAPER, fg=INK_3,
                     font=(UI, 8, "bold")).pack(anchor="w", padx=14, pady=(10, 4))
        return marco

    # ------------------------------------------------------------ encabezado
    cab = tk.Frame(cuerpo, bg=GROUND)
    cab.pack(fill="x", padx=14, pady=(12, 2))
    tk.Label(cab, text="RESERVAS TÉCNICAS QES", bg=GROUND, fg=PLUM,
             font=(UI, 16, "bold")).pack(anchor="w")
    tk.Label(cab, text="Carga los Excel, súbelos al servidor de auditoría, elige las tres "
                       "tablas de la vista y arma el HTML.",
             bg=GROUND, fg=INK_2, font=(UI, 9), wraplength=ancho_texto,
             justify="left").pack(anchor="w")

    # ------------------------------------------------------- 1. carga de archivos
    carga = tarjeta()
    zona = tk.Label(
        carga,
        text=("Arrastra aquí los Excel  ·  o pica para elegirlos"
              if arrastre else "Pica aquí para elegir los Excel"),
        bg=ICE_2, fg=TEAL, font=(UI, 11, "bold"), height=2,
        relief="ridge", bd=1, cursor="hand2")
    zona.pack(fill="x", padx=14, pady=(14, 4))
    tk.Label(carga, text="Balanza de comprobación (.xlsx) y resultados de actuarios (.xlsb). "
                         "El origen de cada archivo se detecta solo.",
             bg=PAPER, fg=INK_3, font=(UI, 8), wraplength=ancho_texto,
             justify="left").pack(anchor="w", padx=14, pady=(0, 8))
    tk.Label(carga, text="LO QUE SE LEYÓ DE CADA ARCHIVO", bg=PAPER, fg=INK_3,
             font=(UI, 8, "bold")).pack(anchor="w", padx=14)
    tarjetas = tk.Frame(carga, bg=PAPER)
    tarjetas.pack(fill="x", padx=14, pady=(4, 12))

    # ----------------------------------------------------------- 2. el servidor
    srv = tarjeta("SERVIDOR DE AUDITORÍA")
    fila1 = tk.Frame(srv, bg=PAPER)
    fila1.pack(fill="x", padx=14, pady=(0, 6))
    tk.Label(fila1, text="Servidor", bg=PAPER, fg=INK_2, font=(UI, 8)).pack(side="left")
    srv_var = tk.StringVar(value=SERVIDOR)
    tk.Entry(fila1, textvariable=srv_var, width=18, font=(MONO, 9)).pack(side="left", padx=(6, 14))
    tk.Label(fila1, text="Base", bg=PAPER, fg=INK_2, font=(UI, 8)).pack(side="left")
    base_var = tk.StringVar(value=BASE)
    tk.Entry(fila1, textvariable=base_var, width=14, font=(MONO, 9)).pack(side="left", padx=(6, 14))
    btn_conectar = ttk.Button(fila1, text="Conectar")
    btn_conectar.pack(side="left")
    btn_esquema = ttk.Button(fila1, text="Crear tablas", state="disabled")
    btn_esquema.pack(side="left", padx=(6, 0))
    btn_subir = ttk.Button(fila1, text="Subir al servidor", state="disabled")
    btn_subir.pack(side="left", padx=(6, 0))

    lbl_srv = tk.Label(srv, text="Sin conectar · el histórico vive en el JSON local.",
                       bg=PAPER, fg=INK_3, font=(UI, 8), anchor="w",
                       wraplength=ancho_texto, justify="left")
    lbl_srv.pack(fill="x", padx=14, pady=(0, 4))
    tk.Label(srv, text=f"Autenticación integrada de Windows · tablas {TABLA_CARGAS} y "
                       f"{TABLA_DETALLE} · cada subida queda como una copia nueva, "
                       "etiquetada por mes y por usuario.",
             bg=PAPER, fg=INK_3, font=(UI, 8), anchor="w",
             wraplength=ancho_texto, justify="left").pack(fill="x", padx=14, pady=(0, 12))

    # -------------------------------------------------- 3. las tres de la vista
    sel = tarjeta("TABLAS PARA LA VISTA")
    tk.Label(sel, text="Elige qué tabla va en cada columna. Con el servidor conectado "
                       "aparece cada copia subida (mes · archivo · usuario · #carga); "
                       "sin servidor, los cortes del histórico local.",
             bg=PAPER, fg=INK_3, font=(UI, 8), anchor="w",
             wraplength=ancho_texto, justify="left").pack(fill="x", padx=14, pady=(0, 8))
    rejilla = tk.Frame(sel, bg=PAPER)
    rejilla.pack(fill="x", padx=14, pady=(0, 6))
    rejilla.columnconfigure(1, weight=1)
    tk.Label(rejilla, text="Columna", bg=PAPER, fg=INK_3, font=(UI, 8, "bold"),
             anchor="w").grid(row=0, column=0, sticky="w")
    tk.Label(rejilla, text="Tabla", bg=PAPER, fg=INK_3, font=(UI, 8, "bold"),
             anchor="w").grid(row=0, column=1, sticky="w", padx=(4, 0))
    tk.Label(rejilla, text="TC de cierre", bg=PAPER, fg=INK_3, font=(UI, 8, "bold"),
             anchor="w").grid(row=0, column=2, sticky="w", padx=(8, 0))
    combos: "dict[str, ttk.Combobox]" = {}
    fx_ranura: "dict[str, tk.StringVar]" = {}
    for r, (clave, etq) in enumerate(RANURAS, start=1):
        tk.Label(rejilla, text=f"{r}. {etq}", bg=PAPER, fg=TEAL,
                 font=(UI, 9, "bold"), width=13, anchor="w").grid(row=r, column=0, sticky="w", pady=2)
        cb = ttk.Combobox(rejilla, state="readonly", font=(UI, 9))
        cb.grid(row=r, column=1, sticky="ew", pady=2, padx=(4, 0))
        combos[clave] = cb
        var = tk.StringVar()
        tk.Entry(rejilla, textvariable=var, width=9, font=(MONO, 9)).grid(
            row=r, column=2, sticky="w", pady=2, padx=(8, 0))
        fx_ranura[clave] = var
    tk.Label(sel, text="El tipo de cambio de cada columna es el de SU cierre: diciembre se "
                       "convierte al de diciembre, no al de hoy. En blanco usa el de abajo. "
                       "La vista guarda las dos monedas y trae un botón para cambiar entre ellas.",
             bg=PAPER, fg=INK_3, font=(UI, 8), anchor="w",
             wraplength=ancho_texto, justify="left").pack(fill="x", padx=14, pady=(2, 6))

    # ---- cortes de más, para cuando la vista deba llevar más de tres --------
    extra_marco = tk.Frame(sel, bg=PAPER)
    extra_marco.pack(fill="x", padx=14, pady=(0, 4))
    tk.Label(extra_marco, text="CORTES ADICIONALES EN LA VISTA", bg=PAPER, fg=INK_3,
             font=(UI, 8, "bold")).pack(anchor="w")
    tk.Label(extra_marco, text="Marca los que quieras añadir a las tres de arriba; "
                               "se acomodan por fecha.",
             bg=PAPER, fg=INK_3, font=(UI, 8), anchor="w",
             wraplength=ancho_texto, justify="left").pack(anchor="w")
    extras = tk.Frame(extra_marco, bg=PAPER)
    extras.pack(fill="x", pady=(4, 0))
    extra_vars: "dict[str, tk.BooleanVar]" = {}
    tk.Frame(sel, bg=PAPER, height=6).pack()

    # -------------------------------------------------------- 4. los controles
    ctl = tk.Frame(cuerpo, bg=GROUND)
    ctl.pack(fill="x", padx=14, pady=(4, 0))
    tk.Label(ctl, text="Tipo de cambio MXN/USD", bg=GROUND, fg=INK_3,
             font=(UI, 8, "bold")).pack(side="left")
    fx_var = tk.StringVar(value=f"{FX_DEFAULT:.4f}")
    tk.Entry(ctl, textvariable=fx_var, width=10, font=(MONO, 9)).pack(side="left", padx=(8, 16))
    btn_procesar = ttk.Button(ctl, text="Procesar")
    btn_procesar.pack(side="right")
    btn_evol = ttk.Button(ctl, text="Ver evolución")
    btn_evol.pack(side="right", padx=(0, 6))
    ttk.Button(ctl, text="Vaciar histórico local",
               command=lambda: vaciar()).pack(side="right", padx=(0, 6))

    # ------------------------------------------------------------ 5. bitácora
    caja = tarjeta("BITÁCORA")
    envoltura = tk.Frame(caja, bg=PAPER)
    envoltura.pack(fill="both", expand=True, padx=14, pady=(0, 12))
    barra = ttk.Scrollbar(envoltura, orient="vertical")
    bitacora = tk.Text(envoltura, bg=PAPER, fg=INK_2, font=(MONO, 8), relief="flat",
                       wrap="word", height=8, yscrollcommand=barra.set,
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

    # ====================================================================
    #  Lectura de archivos
    # ====================================================================
    def pinta_tarjetas() -> None:
        for hijo in tarjetas.winfo_children():
            hijo.destroy()
        if not pendientes:
            tk.Label(tarjetas, text="Todavía no hay archivos cargados en esta sesión."
                                    + (f"  Histórico local: {len(hist.periodos())} corte(s)."
                                       if hist.periodos() else ""),
                     bg=PAPER, fg=INK_2, font=(UI, 9), justify="left",
                     wraplength=ancho_texto).pack(anchor="w")
        for i, (ruta, lec) in enumerate(pendientes.items()):
            t = tk.Frame(tarjetas, bg=ICE_2, highlightbackground=LINE, highlightthickness=1)
            t.pack(fill="x", pady=(0 if i == 0 else 6))
            tk.Label(t, text=lec.titulo, bg=ICE_2, fg=TEAL, font=(UI, 9, "bold"),
                     anchor="w").pack(fill="x", padx=10, pady=(6, 0))
            tk.Label(t, text=f"{lec.archivo}   ·   sha {lec.sha256[:12]}…", bg=ICE_2,
                     fg=INK_3, font=(MONO, 8), anchor="w").pack(fill="x", padx=10)
            tk.Label(t, text=lec.resumen, bg=ICE_2, fg=INK, font=(MONO, 8), anchor="w",
                     justify="left", wraplength=ancho_texto - 20).pack(fill="x", padx=10, pady=(2, 7))
        actualiza_botones()

    def actualiza_botones() -> None:
        hay = bool(pendientes)
        listo = hay or bool(hist.periodos()) or bool(estado["snapshots"])
        btn_procesar.state(["!disabled"] if listo else ["disabled"])
        btn_evol.state(["!disabled"] if listo else ["disabled"])
        btn_esquema.state(["!disabled"] if estado["conectado"] else ["disabled"])
        btn_subir.state(["!disabled"] if (estado["conectado"] and hay) else ["disabled"])

    def carga_rutas(rutas: "Iterable[str]") -> None:
        for ruta in rutas:
            ruta = str(ruta).strip()
            if not ruta:
                continue
            try:
                lec = leer_fuente(ruta)
                pendientes[ruta] = lec
                apunta(f"· {lec.archivo}: {lec.fuente}, "
                       + (", ".join(etiqueta_corta(p) for p in lec.cortes) or "sin cortes"), "ok")
                if lec.faltantes:
                    apunta(f"  faltan las cuentas {', '.join(lec.faltantes)}", "warn")
            except Exception as err:     # se reporta en la bitácora, la ventana sigue viva
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

    # ====================================================================
    #  Servidor
    # ====================================================================
    def conectar() -> None:
        repo.servidor, repo.base = srv_var.get().strip(), base_var.get().strip()
        try:
            donde = repo.conectar()
        except Exception as err:
            estado["conectado"] = False
            lbl_srv.configure(text="No se pudo conectar.", fg=NEG)
            apunta(f"! servidor: {err}", "err")
            actualiza_botones()
            return
        estado["conectado"] = True
        lbl_srv.configure(text=f"Conectado a {donde}", fg="#0E7C66")
        apunta(f"· conectado a {donde}", "ok")
        try:
            if not repo.existe_esquema():
                apunta("  las tablas todavía no existen en esta base: "
                       "pica «Crear tablas» una vez y listo.", "warn")
        except Exception as err:
            apunta(f"! no se pudo revisar el esquema: {err}", "err")
        refresca_snapshots()

    def crear_tablas() -> None:
        try:
            nuevas = repo.crear_esquema()
        except Exception as err:
            apunta(f"! no se pudieron crear las tablas: {err}", "err")
            return
        apunta("· tablas creadas: " + ", ".join(nuevas) if nuevas
               else "· las tablas ya existían, no se tocó nada", "ok")
        refresca_snapshots()

    def subir() -> None:
        if not pendientes:
            apunta("! no hay archivos cargados que subir.", "warn")
            return
        for ruta, lec in list(pendientes.items()):
            try:
                previas = repo.ya_subido(lec.sha256)
                if previas and not messagebox.askyesno(
                        "Reservas técnicas QES",
                        f"{lec.archivo} ya está en el servidor "
                        f"(carga{'s' if len(previas) > 1 else ''} "
                        f"{', '.join('#' + str(c) for c in previas)}).\n\n"
                        "¿Subir otra copia de todos modos?"):
                    apunta(f"· {lec.archivo}: no se subió, ya estaba como "
                           + ", ".join(f"#{c}" for c in previas), "warn")
                    continue
                cid = repo.subir(lec)
                apunta(f"· {lec.archivo} → servidor, carga #{cid} "
                       f"({len(lec.cortes)} corte(s): "
                       f"{', '.join(etiqueta_corta(p) for p in lec.cortes)})", "ok")
            except Exception as err:
                apunta(f"! {lec.archivo}: no se pudo subir · {err}", "err")
        refresca_snapshots()

    btn_conectar.configure(command=conectar)
    btn_esquema.configure(command=crear_tablas)
    btn_subir.configure(command=subir)

    # ====================================================================
    #  Las tres tablas de la vista
    # ====================================================================
    def opciones() -> "list[tuple[str, tuple[str, int | None]]]":
        """(etiqueta que se ve, (periodo, carga_id)) para los desplegables."""
        if estado["conectado"] and estado["snapshots"]:
            return [(f"{etiqueta_corta(s['periodo'])}  ·  {s['fuente']}  ·  {s['archivo']}"
                     f"  ·  {s['usuario']}  ·  #{s['carga_id']}",
                     (s["periodo"], s["carga_id"]))
                    for s in estado["snapshots"]]
        return [(f"{etiqueta_periodo(p)}  ·  histórico local", (p, None))
                for p in reversed(hist.periodos())]

    def refresca_snapshots() -> None:
        if estado["conectado"]:
            try:
                estado["snapshots"] = repo.snapshots()
                apunta(f"· servidor: {len(estado['snapshots'])} copia(s) disponibles "
                       f"en {len({s['periodo'] for s in estado['snapshots']})} corte(s)")
            except Exception as err:
                estado["snapshots"] = []
                apunta(f"! no se pudo leer el catálogo: {err}", "err")
        llena_combos()
        actualiza_botones()

    def llena_combos() -> None:
        ops = opciones()
        estado["ops"] = ops
        etiquetas = [""] + [e for e, _ in ops]
        # por omisión: diciembre más reciente, el último corte, y el anterior
        periodos = []
        for _e, (p, _c) in ops:
            if p not in periodos:
                periodos.append(p)          # ya vienen de más nuevo a más viejo
        dic = next((p for p in periodos if p.split("-")[1] == "12"), None)
        t = periodos[0] if periodos else None
        t1 = next((p for p in periodos[1:] if p != dic), None)
        por_defecto = {"dic": dic, "t1": t1, "t": t}
        for clave, cb in combos.items():
            previo = cb.get()
            cb.configure(values=etiquetas)
            if previo and previo in etiquetas:
                cb.set(previo)          # lo que ya eligió el usuario manda
                continue
            objetivo = por_defecto.get(clave)
            cb.set(next((e for e, (p, _c) in ops if p == objetivo), "") if objetivo else "")
        pinta_extras(periodos)
        refresca_fx()

    def pinta_extras(periodos: "Sequence[str]") -> None:
        """Las casillas de los cortes que no están en las tres ranuras."""
        elegidos = {p for p, _c in seleccion()}
        for hijo in extras.winfo_children():
            hijo.destroy()
        sueltos = [p for p in periodos if p not in elegidos]
        if not sueltos:
            tk.Label(extras, text="No queda ningún otro corte.", bg=PAPER, fg=INK_3,
                     font=(UI, 8)).pack(anchor="w")
            return
        for i, per in enumerate(sorted(sueltos, reverse=True)):
            var = extra_vars.setdefault(per, tk.BooleanVar(value=False))
            tk.Checkbutton(extras, text=etiqueta_corta(per), variable=var,
                           bg=PAPER, fg=INK, activebackground=PAPER, selectcolor=PAPER,
                           font=(UI, 8)).grid(row=i // 4, column=i % 4, sticky="w", padx=(0, 12))

    def refresca_fx() -> None:
        """Prellena el tipo de cambio de cada ranura con el guardado del corte."""
        for (clave, _etq), (per, _c) in zip(RANURAS, seleccion() + [(None, None)] * 3):
            var = fx_ranura[clave]
            if var.get().strip():
                continue                      # lo que ya escribió el usuario manda
            guardado = hist.datos.get(per, {}).get("fx") if per else None
            var.set(f"{float(guardado):.4f}" if guardado else "")

    def seleccion() -> "list[tuple[str, int | None]]":
        """Los tres pares (periodo, carga) elegidos, en el orden de la matriz."""
        mapa = dict(estado.get("ops", []))
        out = []
        for clave, _etq in RANURAS:
            v = combos[clave].get()
            if v and v in mapa and mapa[v] not in out:
                out.append(mapa[v])
        return out

    def guarda_fx() -> None:
        """Pasa al histórico el tipo de cambio escrito en cada ranura."""
        for (clave, etq), (per, _c) in zip(RANURAS, seleccion() + [(None, None)] * 3):
            if per is None:
                continue
            txt = fx_ranura[clave].get().strip().replace(",", "")
            if not txt:
                hist.poner_fx(per, None)
                continue
            try:
                v = float(txt)
                if v <= 0:
                    raise ValueError
            except ValueError:
                apunta(f"! tipo de cambio no válido en «{etq}»: «{txt}». Se ignora.", "warn")
                continue
            hist.poner_fx(per, v)

    def historico_de_la_vista() -> "tuple[Historico, list[str]]":
        """El Historico y los periodos que le tocan, según lo elegido arriba."""
        pares = seleccion()
        marcados = sorted(p for p, v in extra_vars.items() if v.get())
        if estado["conectado"] and estado["snapshots"] and pares:
            escogidas = [(p, c) for p, c in pares if c is not None]
            escogidas += [(p, None) for p in marcados]
            h = repo.historico(escogidas, ruta_json=hist.ruta)
            ps = sorted({p for p, _c in escogidas})
            for per in ps:                       # el TC de cierre viaja con la vista
                v = hist.datos.get(per, {}).get("fx")
                if v:
                    h.poner_fx(per, v)
            return h, ps
        ps = sorted(set([p for p, _c in pares] + marcados)) or hist.periodos()[-PERIODOS_EN_VISTA:]
        return hist, [p for p in ps if p in hist.datos]

    # ====================================================================
    #  Acciones
    # ====================================================================
    def tipo_de_cambio() -> float:
        try:
            fx = float(fx_var.get().replace(",", ""))
            if fx <= 0:
                raise ValueError
            return fx
        except ValueError:
            apunta(f"! tipo de cambio no válido: «{fx_var.get()}». Se usa {FX_DEFAULT}.", "warn")
            return FX_DEFAULT

    def funde_pendientes() -> None:
        # los actuarios primero: para la columna local manda la balanza
        orden = sorted(pendientes.values(), key=lambda l: l.fuente != "actuarios")
        for lec in orden:
            try:
                for aviso in hist.fundir(lec):
                    apunta("· " + aviso,
                           "warn" if "Revisar" in aviso or "SIN CUENTA" in aviso else "")
            except Exception as err:
                apunta(f"! {lec.archivo}: {err}", "err")
        if orden:
            hist.guardar()
            pendientes.clear()
            pinta_tarjetas()
            llena_combos()

    def procesar() -> None:
        fx = tipo_de_cambio()
        funde_pendientes()
        guarda_fx()
        hist.guardar()
        h, ps = historico_de_la_vista()
        if not ps:
            apunta("! no hay cortes que mostrar: carga un archivo o elige las tablas.", "err")
            return
        apunta(f"· histórico local en {hist.ruta} ({len(hist.periodos())} corte(s))")
        try:
            ruta_html = escribir_vista(h, periodos=ps, fx=fx, nota=estado["nota"])
        except Exception as err:
            apunta(f"! no se pudo escribir la vista: {err}", "err")
            return
        apunta(f"· vista de {', '.join(etiqueta_corta(p) for p in ps)} escrita en "
               f"{ruta_html}", "ok")
        for p in ps:
            if h.datos.get(p, {}).get("aviso"):
                apunta(f"  revisar {etiqueta_periodo(p)}: {h.datos[p]['aviso']}", "warn")
        apunta("· abierta en el navegador. El archivo es autocontenido: se puede mandar "
               "por correo tal cual.", "ok")

    def evolucion() -> None:
        fx = tipo_de_cambio()
        funde_pendientes()
        guarda_fx()
        hist.guardar()
        if estado["conectado"] and estado["snapshots"]:
            h = repo.historico(ruta_json=hist.ruta)   # todo lo que haya en el servidor
            for per in h.periodos():
                v = hist.datos.get(per, {}).get("fx")
                if v:
                    h.poner_fx(per, v)
        else:
            h = hist
        try:
            ruta_html = escribir_evolucion(h, fx=fx)
        except Exception as err:
            apunta(f"! no se pudo dibujar la evolución: {err}", "err")
            return
        apunta(f"· evolución de {len([p for p in h.periodos() if h.completo(p)])} corte(s) "
               f"escrita en {ruta_html}", "ok")

    def vaciar() -> None:
        if messagebox.askyesno("Reservas técnicas QES",
                               "¿Vaciar el histórico local?\n\n"
                               "Lo que ya está en el servidor no se toca."):
            hist.datos.clear()
            hist.guardar()
            apunta("· histórico local vaciado (el servidor queda intacto)", "warn")
            llena_combos()
            actualiza_botones()

    btn_procesar.configure(command=procesar)
    btn_evol.configure(command=evolucion)

    # ------------------------------------------------------------ arranque
    apunta(f"· histórico local: {Path(hist_ruta).resolve()} ({len(hist.periodos())} corte(s))")
    if hist.periodos():
        apunta("  cortes: " + ", ".join(etiqueta_corta(p) for p in hist.periodos()))
    if not arrastre:
        apunta("  (instala tkinterdnd2 si quieres arrastrar y soltar: pip install tkinterdnd2)",
               "warn")
    apunta(f"· servidor por conectar: {repo.servidor} · base {repo.base}")
    for linea in diagnostico():
        apunta("  " + linea, "warn" if "FALTA" in linea or "ninguno" in linea else "")
    pinta_tarjetas()
    llena_combos()
    actualiza_botones()

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
