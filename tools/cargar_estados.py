"""Carga el estado de cada sede desde el archivo de trabajo 'Campos dashborad'
hacia el maestro 'CONTROL DE EQUIPOS MANTENIMIENTO.xlsx' (Hoja3, columna
ESTADO COLSOF), que es la fuente que lee dashboard_web/pipeline.py.

Sin este paso el dashboard queda en 0 ejecutadas aunque el archivo de trabajo
ya tenga sedes finalizadas.

Verifica la alineacion de las sedes (SBAN) posicion por posicion antes de
escribir, y aborta si no coinciden. La edicion se hace por cirugia de XML para
preservar formulas, estilos y validaciones.
"""
from __future__ import annotations

import re
import shutil
import sys
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

REPO = Path(__file__).resolve().parents[1]
TRABAJO = Path(r"C:\Users\Admin\Videos\SHORTS YT\.hermes\desktop-attachments\Campos dashborad 28 sept.xlsx")
MAESTRO = REPO / "CONTROL DE EQUIPOS MANTENIMIENTO.xlsx"

HOJA_TRABAJO = "Dashboard_KPI"
COLS_TRABAJO = {"sban": "A", "estado": "N"}   # Estado de la sede
HOJA_MAESTRO = "Hoja3"
COLS_MAESTRO = {"sban": "C", "estado": "M"}   # ESTADO COLSOF

# El dashboard usa estos nombres canonicos
CANONICO = {
    "EN PROCESO": "En Proceso",
    "EN SITIO": "En Proceso",
    "EN EJECUCION": "En Proceso",
    "FINALIZADA": "Finalizada",
    "PROGRAMADA": "Programada",
    "REPROGRAMADA": "Reprogramada",
    "CANCELADA": "Cancelada",
    "SEDE_CERRADA": "Sede_Cerrada",
}


def mapa_hoja_a_xml(zf: ZipFile) -> dict[str, str]:
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rid = {r.attrib["Id"]: r.attrib["Target"] for r in rels}
    salida = {}
    for h in wb.find(f"{NS}sheets"):
        destino = rid[h.attrib[f"{REL_NS}id"]]
        salida[h.attrib["name"]] = destino if destino.startswith("xl/") else "xl/" + destino.lstrip("/")
    return salida


def leer_strings(zf: ZipFile) -> tuple[str, list[str]]:
    texto = zf.read("xl/sharedStrings.xml").decode("utf-8")
    vals = ["".join(re.findall(r"<t[^>]*>(.*?)</t>", si, flags=re.S))
            for si in re.findall(r"<si>(.*?)</si>", texto, flags=re.S)]
    return texto, vals


def leer_columnas(zf: ZipFile, sheet_xml: str, strings: list[str], cols: set[str]) -> dict[int, dict[str, object]]:
    filas: dict[int, dict[str, object]] = {}
    with zf.open(sheet_xml) as stream:
        for _, row in ET.iterparse(stream, events=("end",)):
            if row.tag != f"{NS}row":
                continue
            n = int(row.attrib.get("r", "0"))
            d: dict[str, object] = {}
            for c in row.findall(f"{NS}c"):
                ref = c.attrib.get("r", "")
                col = re.match(r"[A-Z]+", ref).group(0)
                if col not in cols:
                    continue
                t = c.attrib.get("t")
                if t == "inlineStr":
                    d[col] = "".join(x.text or "" for x in c.iter(f"{NS}t"))
                    continue
                v = c.find(f"{NS}v")
                d[col] = None if v is None else (strings[int(v.text)] if t == "s" else v.text)
            filas[n] = d
            row.clear()
    return filas


def sban_key(valor: object):
    m = re.match(r"^(\d+)", str(valor or "").strip())
    return int(m.group(1)) if m else None


def escapar(v: str) -> str:
    return v.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def normalizar_estado(v: object) -> str | None:
    s = str(v or "").strip()
    if not s:
        return None
    return CANONICO.get(s.upper(), CANONICO.get(s, s))


def main() -> None:
    for ruta in (TRABAJO, MAESTRO):
        if not ruta.exists():
            sys.exit(f"ERROR: no existe {ruta}")

    with ZipFile(TRABAJO) as zt:
        hoja = mapa_hoja_a_xml(zt)[HOJA_TRABAJO]
        _, strings_t = leer_strings(zt)
        filas_t = leer_columnas(zt, hoja, strings_t, set(COLS_TRABAJO.values()))

    datos = []
    for n in sorted(filas_t):
        d = filas_t[n]
        est = normalizar_estado(d.get(COLS_TRABAJO["estado"]))
        sb = sban_key(d.get(COLS_TRABAJO["sban"]))
        if sb is not None and est:
            datos.append((sb, est))

    with ZipFile(MAESTRO) as zm:
        hoja_m = mapa_hoja_a_xml(zm)[HOJA_MAESTRO]
        strings_texto, strings = leer_strings(zm)
        filas_m = leer_columnas(zm, hoja_m, strings, set(COLS_MAESTRO.values()))
        # Ignora filas vacias o de relleno (sin SBAN)
        filas_validas = [n for n in sorted(filas_m)
                         if sban_key(filas_m[n].get(COLS_MAESTRO["sban"])) is not None]
        orden = [sban_key(filas_m[n].get(COLS_MAESTRO["sban"])) for n in filas_validas]
        contenido = zm.read(hoja_m).decode("utf-8")

        if len(orden) != len(datos):
            sys.exit(f"ERROR: filas distintas ({len(orden)} maestro vs {len(datos)} trabajo)")
        desalineadas = [(r, a, b) for r, (a, b) in enumerate(zip(orden, [d[0] for d in datos]), start=2) if a != b]
        if desalineadas:
            sys.exit(f"ERROR: sedes desalineadas: {desalineadas[:5]}")

        cambio: dict[int, str] = {}
        for fila, (_, estado) in zip(filas_validas, datos):
            actual = str(filas_m[fila].get(COLS_MAESTRO["estado"]) or "").strip()
            if actual != estado:
                cambio[fila] = estado

        if not cambio:
            print("Nada por cambiar: el maestro ya tiene los estados actualizados.")
            return

        for est in sorted(set(cambio.values())):
            if est not in strings:
                strings.append(est)
                strings_texto = strings_texto.replace(
                    "</sst>", f'<si><t xml:space="preserve">{escapar(est)}</t></si></sst>', 1)
        col = COLS_MAESTRO["estado"]
        patron = re.compile(r'<c r="%s(\d+)"([^>]*?)(?:/>|>.*?</c>)' % col, re.S)

        def reemplazo(m: re.Match) -> str:
            fila = int(m.group(1))
            if fila not in cambio:
                return m.group(0)
            estilo = re.search(r'\bs="(\d+)"', m.group(2))
            attr = f' s="{estilo.group(1)}"' if estilo else ""
            return f'<c r="{col}{fila}"{attr} t="s"><v>{strings.index(cambio[fila])}</v></c>'

        contenido = patron.sub(reemplazo, contenido)
        total = str(len(strings))
        strings_texto = re.sub(r'(<sst[^>]*?\bcount=")\d+(")', r"\g<1>%s\g<2>" % total, strings_texto, count=1)
        strings_texto = re.sub(r'(<sst[^>]*?\buniqueCount=")\d+(")', r"\g<1>%s\g<2>" % total, strings_texto, count=1)

        temp = MAESTRO.with_suffix(".tmp.xlsx")
        with ZipFile(temp, "w", ZIP_DEFLATED) as out:
            for item in zm.infolist():
                if item.filename == "xl/sharedStrings.xml":
                    data = strings_texto.encode("utf-8")
                elif item.filename == hoja_m:
                    data = contenido.encode("utf-8")
                else:
                    data = zm.read(item.filename)
                out.writestr(item, data)
        shutil.move(str(temp), str(MAESTRO))

    from collections import Counter
    print(f"Estados cargados en {MAESTRO.name} [{HOJA_MAESTRO}] col {col}: {len(cambio)} celdas")
    print("Resultado:", dict(Counter(cambio.values())))


if __name__ == "__main__":
    main()
