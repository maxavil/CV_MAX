#!/usr/bin/env python3
"""Lectura de los dos Excel de reservas técnicas QES y armado del histórico.

Equivalente en Python del bloque de ingesta de reservas/index.html.

    pip install openpyxl pyxlsb
    python reservas_qes.py Balanza_062026.xlsx ResultadosQES.xlsb

Reglas:
  * Metodología local  -> balanza de comprobación, columna de 3er grado de las
    cuentas 2205, 2301 y 2302. Los saldos vienen en negativo por ser pasivos y
    se les invierte el signo.
  * Método Estatutario CNSF -> archivo de actuarios, concepto por concepto,
    tomando todos los cortes que traiga el archivo.
  * Diferencia = estatutario - local.

Cada archivo procesado actualiza su periodo y conserva los demás: el histórico
se guarda en un JSON y se va conversando con el mes que sigue.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Sequence

# =========================================================================
# 1. Configuración de conceptos
# =========================================================================


@dataclass(frozen=True)
class Concepto:
    id: str
    label: str
    cuenta: str
    patron: re.Pattern


CONCEPTOS: list[Concepto] = [
    Concepto("rrc",  "Reserva de Riesgos en Curso",         "2205", re.compile(r"riesgos en curso")),
    Concepto("rsr",  "Reserva de Siniestros Reportados",    "2301", re.compile(r"siniestros reportados")),
    Concepto("rsnr", "Reserva de Siniestros No Reportados", "2302", re.compile(r"no reportados")),
]

GRADO = 3  # columna de la balanza de la que se toma el saldo

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
         "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

# =========================================================================
# 2. Utilidades
# =========================================================================


def norm(v: Any) -> str:
    """Minúsculas, sin acentos y con espacios colapsados."""
    s = "" if v is None else str(v)
    s = unicodedata.normalize("NFD", s.lower())
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s).strip()


def to_num(v: Any) -> float | None:
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


def serial_a_fecha(n: float) -> str | None:
    """Serial de Excel -> 'AAAA-MM-DD' (base 1899-12-30)."""
    if not isinstance(n, (int, float)) or isinstance(n, bool):
        return None
    if n < 20000 or n > 80000:
        return None
    d = dt.date(1899, 12, 30) + dt.timedelta(days=int(round(n)))
    return d.isoformat()


def celda_a_fecha(v: Any) -> str | None:
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return serial_a_fecha(v)
    return None


def parse_fecha(texto: Any) -> str | None:
    """'al 30 de Junio 2026' / '30 de junio de 2026' -> '2026-06-30'."""
    m = re.search(r"(\d{1,2})\s+de\s+([a-z]+)\s*(?:de\s*)?(\d{4})", norm(texto))
    if not m:
        return None
    if m.group(2) not in MESES:
        return None
    return dt.date(int(m.group(3)), MESES.index(m.group(2)) + 1, int(m.group(1))).isoformat()


def periodo_de_nombre(nombre: str) -> str | None:
    """'Balanza_062026.xlsx' -> '2026-06-30'."""
    m = re.search(r"(0[1-9]|1[0-2])[_\-.]?(20\d{2})", str(nombre))
    if not m:
        return None
    mes, anio = int(m.group(1)), int(m.group(2))
    return dt.date(anio, mes, ultimo_dia(anio, mes)).isoformat()


def etiqueta_periodo(k: str) -> str:
    a, m, d = k.split("-")
    return f"{int(d)} {MESES[int(m) - 1]} {a}"


def etiqueta_corta(k: str) -> str:
    a, m, _ = k.split("-")
    return f"{MESES[int(m) - 1].capitalize()} {a}"


def match_concepto(texto: Any) -> Concepto | None:
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


# =========================================================================
# 3. Apertura del libro (.xlsx / .xlsm y .xlsb con el mismo resultado)
# =========================================================================

Hoja = tuple[str, list[list[Any]]]


def leer_libro(ruta: str | Path) -> list[Hoja]:
    """Devuelve [(nombre de hoja, filas como listas densas)] para xlsx o xlsb."""
    ruta = Path(ruta)
    ext = ruta.suffix.lower()

    if ext == ".xlsb":
        try:
            from pyxlsb import open_workbook
        except ImportError as err:  # pragma: no cover
            raise SystemExit("Para leer .xlsb hace falta pyxlsb: pip install pyxlsb") from err
        hojas: list[Hoja] = []
        with open_workbook(str(ruta)) as wb:
            for nombre in wb.sheets:
                with wb.get_sheet(nombre) as sh:
                    filas: list[list[Any]] = []
                    for fila in sh.rows():
                        densa: list[Any] = []
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
        except ImportError as err:  # pragma: no cover
            raise SystemExit("Para leer .xlsx hace falta openpyxl: pip install openpyxl") from err
        wb = openpyxl.load_workbook(str(ruta), data_only=True, read_only=True)
        hojas = [(ws.title, [list(f) for f in ws.iter_rows(values_only=True)]) for ws in wb.worksheets]
        wb.close()
        return hojas

    raise SystemExit(f"Extensión no soportada: {ruta.name} (usa .xlsx, .xlsm o .xlsb)")


# =========================================================================
# 4. Lectura de la balanza de comprobación
# =========================================================================


@dataclass
class LecturaBalanza:
    periodo: str | None
    valores: dict[str, float]
    detalle: list[str]
    faltantes: list[str]
    hoja: str


def parse_balanza(hojas: Sequence[Hoja], fname: str) -> LecturaBalanza | None:
    """Metodología local: cuentas 2205 / 2301 / 2302 en la columna de 3er grado."""
    for nombre, filas in hojas:
        h = next((i for i, f in enumerate(filas[:60]) if f and norm(f[0]) == "cuenta"), None)
        if h is None:
            continue  # sin encabezado CUENTA no es una balanza

        # columnas por grado: la fila de encabezado dice "GRADO" y la de arriba el ordinal
        grados: dict[int, int] = {}
        hdr = filas[h]
        arriba = filas[h - 1] if h > 0 else []
        for j in range(1, len(hdr)):
            if "grado" in norm(hdr[j]):
                m = re.search(r"(\d)", str(arriba[j]) if j < len(arriba) and arriba[j] is not None else "")
                if m:
                    grados[int(m.group(1))] = j
        cols_grado = list(grados.values()) or list(range(2, 9))

        # índice de cuentas
        cuentas: dict[str, list[Any]] = {}
        for fila in filas[h + 1:]:
            if not fila or fila[0] is None:
                continue
            clave = str(fila[0]).strip()
            if clave and clave not in cuentas:
                cuentas[clave] = fila

        # periodo: encabezado del reporte, si no, nombre del archivo
        periodo = None
        for fila in filas[:h + 1]:
            periodo = parse_fecha(" ".join("" if c is None else str(c) for c in (fila or [])))
            if periodo:
                break
        if not periodo:
            periodo = periodo_de_nombre(fname)

        valores: dict[str, float] = {}
        detalle: list[str] = []
        faltantes: list[str] = []
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
            if val is None:  # respaldo: primer grado con importe en esa fila
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
            valores[c.id] = abs(val)  # pasivo en negativo -> se invierte
            grado_txt = f"{col_usada}º grado " if col_usada else ""
            detalle.append(f"{c.cuenta} {grado_txt}{fmt(abs(val))}")

        if not valores:
            continue
        return LecturaBalanza(periodo, valores, detalle, faltantes, nombre)

    return None  # no es balanza -> se intenta como archivo de actuarios


# =========================================================================
# 5. Lectura del archivo de actuarios (todas las fechas que traiga)
# =========================================================================


@dataclass
class LecturaActuarios:
    periodos: dict[str, dict[str, dict[str, float]]] = field(default_factory=dict)
    hojas: list[str] = field(default_factory=list)


def parse_actuarios(hojas: Sequence[Hoja]) -> LecturaActuarios:
    """Busca el texto de cada concepto y toma los dos primeros importes a su derecha."""
    out = LecturaActuarios()
    for nombre, filas in hojas:
        titulo: str | None = None
        usada = False
        for fila in filas:
            if not fila:
                continue
            concepto: Concepto | None = None
            ci = -1
            for j, v in enumerate(fila):
                if not isinstance(v, str):
                    continue
                c = match_concepto(v)
                if c and ci < 0:
                    concepto, ci = c, j
                elif not c:
                    f = parse_fecha(v)  # título del bloque: "30 de junio de 2026"
                    if f:
                        titulo = f
            if concepto is None:
                continue

            periodo = next((p for p in (celda_a_fecha(fila[j]) for j in range(ci)) if p), None) or titulo
            if not periodo:
                continue

            nums: list[float] = []
            for j in range(ci + 1, len(fila)):
                n = to_num(fila[j])
                if n is not None:
                    nums.append(abs(n))
                if len(nums) == 2:
                    break
            if len(nums) < 2:
                continue

            bloque = out.periodos.setdefault(periodo, {"local": {}, "cnsf": {}})
            bloque["local"][concepto.id] = nums[0]  # metodología local (para contraste)
            bloque["cnsf"][concepto.id] = nums[1]   # método estatutario CNSF
            usada = True
        if usada:
            out.hojas.append(nombre)
    return out


# =========================================================================
# 6. Histórico: cada archivo actualiza su periodo y conserva los demás
# =========================================================================


class Historico:
    def __init__(self, ruta: str | Path = "historico_reservas.json"):
        self.ruta = Path(ruta)
        self.datos: dict[str, dict[str, Any]] = {}
        if self.ruta.exists():
            self.datos = json.loads(self.ruta.read_text(encoding="utf-8"))

    # ---------------------------------------------------------------- io
    def guardar(self) -> None:
        self.ruta.write_text(json.dumps(self.datos, indent=2, ensure_ascii=False), encoding="utf-8")

    def periodos(self) -> list[str]:
        return sorted(self.datos)

    def _entrada(self, p: str) -> dict[str, Any]:
        return self.datos.setdefault(p, {"periodo": p, "local": {}, "cnsf": {}, "origen": {}})

    # ----------------------------------------------------------- ingesta
    def procesar(self, ruta: str | Path) -> list[str]:
        """Lee un Excel, detecta de cuál se trata y lo funde al histórico."""
        ruta = Path(ruta)
        hojas = leer_libro(ruta)
        avisos: list[str] = []

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
                # la balanza manda; esto solo rellena periodos que aún no la tienen
                for cid, val in src["local"].items():
                    e["local"].setdefault(cid, val)
                if not origen_local:
                    e["origen"]["local"] = f"Actuarios · {ruta.name}"
            e["origen"]["cnsf"] = f"Actuarios · {ruta.name}"
            self._revisar(e)
        avisos.append(
            f"{ruta.name} → actuarios, {len(act.periodos)} periodo(s): "
            + ", ".join(etiqueta_corta(p) for p in sorted(act.periodos))
            + " · hojas: " + ", ".join(act.hojas)
        )
        avisos += [f"Revisar {etiqueta_periodo(p)}: {self.datos[p]['aviso']}"
                   for p in self.periodos() if self.datos[p].get("aviso")]
        return avisos

    @staticmethod
    def _revisar(e: dict[str, Any]) -> None:
        """Contrasta la columna local de la balanza contra la del archivo de actuarios."""
        e.pop("aviso", None)
        ref = e.get("local_actuarios")
        if not ref:
            return
        difs = [
            f"{c.label.replace('Reserva de ', '')}: balanza {fmt(e['local'][c.id])} vs actuarios {fmt(ref[c.id])}"
            for c in CONCEPTOS
            if c.id in e["local"] and c.id in ref and abs(e["local"][c.id] - ref[c.id]) > 0.5
        ]
        if difs:
            e["aviso"] = "Metodología local no coincide — " + "; ".join(difs)

    # -------------------------------------------------------------- vista
    def total(self, periodo: str, lado: str) -> float:
        return sum(self.datos[periodo][lado].get(c.id, 0.0) for c in CONCEPTOS)

    def diferencia(self, periodo: str, cid: str | None = None) -> float | None:
        e = self.datos[periodo]
        if cid is None:
            return self.total(periodo, "cnsf") - self.total(periodo, "local")
        if cid not in e["cnsf"] or cid not in e["local"]:
            return None
        return e["cnsf"][cid] - e["local"][cid]

    def vista(self, periodos: Iterable[str] | None = None, escala: float = 1e6) -> str:
        """La vista comparativa en texto: conceptos por periodo más el total."""
        ps = list(periodos) if periodos else self.periodos()[-3:]
        if not ps:
            return "Histórico vacío."
        ancho = max(len(c.label) for c in CONCEPTOS) + 2
        sep = "  "

        def celda(v: float | None) -> str:
            return "—".rjust(11) if v is None else f"{v / escala:,.2f}".rjust(11)

        cab1 = "".ljust(ancho) + sep + sep.join(etiqueta_periodo(p).center(37) for p in ps)
        cab2 = "RESERVA".ljust(ancho) + sep + sep.join(
            "local".rjust(11) + sep + "CNSF".rjust(11) + sep + "dif.".rjust(11) for _ in ps)
        if len(ps) > 1:
            cab1 += sep + f"vs {etiqueta_corta(ps[-2])}".center(11)
            cab2 += sep + "incremento".rjust(11)
        lineas = [cab1, cab2, "-" * len(cab2)]

        for c in CONCEPTOS:
            fila = c.label.ljust(ancho)
            for p in ps:
                e = self.datos[p]
                fila += sep + celda(e["local"].get(c.id)) + sep + celda(e["cnsf"].get(c.id)) + sep + celda(self.diferencia(p, c.id))
            if len(ps) > 1:
                a, b = self.diferencia(ps[-1], c.id), self.diferencia(ps[-2], c.id)
                fila += sep + celda(a - b if a is not None and b is not None else None)
            lineas.append(fila)

        fila = "TOTAL RESERVAS".ljust(ancho)
        for p in ps:
            fila += sep + celda(self.total(p, "local")) + sep + celda(self.total(p, "cnsf")) + sep + celda(self.diferencia(p))
        if len(ps) > 1:
            fila += sep + celda(self.diferencia(ps[-1]) - self.diferencia(ps[-2]))
        lineas += ["-" * len(cab2), fila]

        unidad = "millones de USD" if escala == 1e6 else "USD"
        lineas.append("")
        lineas.append(f"Cifras en {unidad}. Diferencia = Método Estatutario CNSF − Metodología local.")
        return "\n".join(lineas)


# =========================================================================
# 7. Línea de comandos
# =========================================================================


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Arma la vista de reservas técnicas QES desde los dos Excel.")
    ap.add_argument("archivos", nargs="*", help="Balanza (.xlsx) y archivo de actuarios (.xlsb/.xlsx), en cualquier orden.")
    ap.add_argument("--hist", default="historico_reservas.json", help="JSON donde vive el histórico acumulado.")
    ap.add_argument("--periodos", nargs="*", help="Periodos a mostrar (AAAA-MM-DD). Por omisión, los tres últimos.")
    ap.add_argument("--usd", action="store_true", help="Mostrar en USD en vez de millones.")
    args = ap.parse_args(argv)

    hist = Historico(args.hist)
    for ruta in args.archivos:
        try:
            for aviso in hist.procesar(ruta):
                print(f"· {aviso}", file=sys.stderr)
        except (ValueError, SystemExit) as err:
            print(f"! {err}", file=sys.stderr)
            return 1
    if args.archivos:
        hist.guardar()
        print(f"· histórico guardado en {hist.ruta} ({len(hist.periodos())} periodo(s))\n", file=sys.stderr)

    print(hist.vista(args.periodos, escala=1.0 if args.usd else 1e6))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
