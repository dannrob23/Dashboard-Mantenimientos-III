"""Actualiza la columna 'Jefaturas Operaciones Regional' con el nombre real de la
jefatura, usando como fuente los codigos/nombres de la columna AA del archivo
'28 sept 3.xlsx'.

Regla de negocio aplicada:
- Solo se actualizan las filas que representan una JEFATURA (el nombre de la
  oficina contiene 'JEFAT').
- Las filas de tipo 'GERENCIA REGIONAL' / 'REGIONAL REGIONAL' NO se modifican,
  porque su codigo (1300, 3000, ...) agrupa varias jefaturas y no identifican
  una jefatura unica.

La edicion se hace por cirugia de XML dentro del .xlsx para preservar intactos
formulas, estilos, validaciones y graficos de los archivos originales.
"""
from __future__ import annotations

import csv
import re
import shutil
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

REPO = Path(__file__).resolve().parents[1]
ATT = Path(r"C:\Users\Admin\Videos\SHORTS YT\.hermes\desktop-attachments")
FUENTE_JEFATURAS = ATT / "28 sept 3.xlsx"
SALIDA = REPO / "salida_jefaturas"

# (archivo, hoja, columna_codigo, columna_oficina, columna_jefatura, destino)
OBJETIVOS = [
    (ATT / "Campos dashborad 28 sept.xlsx", "Dashboard_KPI", "A", "C", "F",
     SALIDA / "Campos dashborad 28 sept_actualizado.xlsx"),
    (REPO / "CONTROL DE EQUIPOS MANTENIMIENTO.xlsx", "Hoja1", "A", "C", "F",
     SALIDA / "CONTROL DE EQUIPOS MANTENIMIENTO_actualizado.xlsx"),
    (REPO / "CONTROL DE EQUIPOS MANTENIMIENTO.xlsx", "Hoja3", "C", "E", "H",
     None),  # se edita en el mismo archivo que Hoja1
]


def normalizar(texto: object) -> str:
    s = str(texto or "").upper()
    for a, b in (("Á", "A"), ("É", "E"), ("Í", "I"), ("Ó", "O"), ("Ú", "U"), ("Ü", "U")):
        s = s.replace(a, b)
    return re.sub(r"[^A-Z0-9]+", " ", s).strip()


def mapa_hoja_a_xml(zf: ZipFile) -> dict[str, str]:
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rid_a_destino = {r.attrib["Id"]: r.attrib["Target"] for r in rels}
    salida = {}
    for hoja in wb.find(f"{NS}sheets"):
        rid = hoja.attrib.get(f"{REL_NS}id")
        destino = rid_a_destino.get(rid, "")
        if not destino.startswith("xl/"):
            destino = "xl/" + destino.lstrip("/")
        salida[hoja.attrib["name"]] = destino
    return salida


def leer_strings(zf: ZipFile) -> tuple[str, list[str]]:
    texto = zf.read("xl/sharedStrings.xml").decode("utf-8")
    valores = []
    for si in re.findall(r"<si>(.*?)</si>", texto, flags=re.S):
        valores.append("".join(re.findall(r"<t[^>]*>(.*?)</t>", si, flags=re.S)))
    return texto, valores


def leer_celdas(zf: ZipFile, sheet_xml: str, strings: list[str], columnas: set[str]) -> dict[int, dict[str, object]]:
    """Devuelve {fila: {columna: valor}} para las columnas indicadas."""
    filas: dict[int, dict[str, object]] = {}
    with zf.open(sheet_xml) as stream:
        for _, row in ET.iterparse(stream, events=("end",)):
            if row.tag != f"{NS}row":
                continue
            numero = int(row.attrib.get("r", "0"))
            datos: dict[str, object] = {}
            for celda in row.findall(f"{NS}c"):
                ref = celda.attrib.get("r", "")
                col = re.match(r"[A-Z]+", ref).group(0)
                if col not in columnas:
                    continue
                tipo = celda.attrib.get("t")
                if tipo == "inlineStr":
                    datos[col] = "".join(t.text or "" for t in celda.iter(f"{NS}t"))
                    continue
                v = celda.find(f"{NS}v")
                if v is None:
                    datos[col] = None
                    continue
                crudo = v.text or ""
                datos[col] = strings[int(crudo)] if tipo == "s" else crudo
            filas[numero] = datos
            row.clear()
    return filas


def leer_fuente() -> dict[str, list[str]]:
    """codigo (sin ceros a la izquierda) -> nombres de jefatura, desde AA/AB."""
    with ZipFile(FUENTE_JEFATURAS) as zf:
        hoja = mapa_hoja_a_xml(zf).get("Hoja1", "xl/worksheets/sheet1.xml")
        _, strings = leer_strings(zf)
        filas = leer_celdas(zf, hoja, strings, {"AA", "AB"})
    mapa: dict[str, list[str]] = {}
    for _, d in filas.items():
        aa = str(d.get("AA") or "").strip()
        ab = str(d.get("AB") or "").strip()
        m = re.match(r"^(\d+)\s*-\s*(.+)$", aa)
        if not m or not ab.upper().startswith("JEFAT"):
            continue
        codigo = str(int(m.group(1)))
        if ab not in mapa.setdefault(codigo, []):
            mapa[codigo].append(ab)
    return mapa


def elegir_jefatura(codigo: str, oficina: str, mapa: dict[str, list[str]]) -> str | None:
    candidatos = mapa.get(codigo, [])
    nom_oficina = normalizar(oficina)
    coincidencias = []
    for nombre in candidatos:
        cola = normalizar(nombre.replace("JEFATURA", "", 1).replace("JEFATUTA", "", 1))
        if cola and cola in nom_oficina and nombre not in coincidencias:
            coincidencias.append(nombre)
    return coincidencias[0] if len(coincidencias) == 1 else None


def escapar(v: str) -> str:
    return v.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def actualizar_hoja(contenido: str, strings_texto: str, strings: list[str],
                    filas: dict[int, dict[str, object]], col_codigo: str,
                    col_oficina: str, col_jefatura: str,
                    mapa: dict[str, list[str]]) -> tuple[str, str, list[dict[str, str]]]:
    cambios: list[dict[str, str]] = []
    for numero, d in sorted(filas.items()):
        oficina = str(d.get(col_oficina) or "")
        if "JEFAT" not in normalizar(oficina):
            continue
        crudo = d.get(col_codigo)
        m = re.match(r"^(\d+)", str(crudo or "").strip())
        if not m:
            continue
        codigo = str(int(m.group(1)))
        objetivo = elegir_jefatura(codigo, oficina, mapa)
        if not objetivo:
            continue
        actual = str(d.get(col_jefatura) or "")
        if normalizar(actual) == normalizar(objetivo):
            continue
        if objetivo in strings:
            idx = strings.index(objetivo)
        else:
            idx = len(strings)
            strings.append(objetivo)
            strings_texto = strings_texto.replace(
                "</sst>", f'<si><t xml:space="preserve">{escapar(objetivo)}</t></si></sst>', 1)
        # reemplazo quirurgico de la celda conservando su estilo (atributo s)
        patron = re.compile(r'<c r="%s%d"([^>]*?)(?:/>|>.*?</c>)' % (col_jefatura, numero), re.S)
        coincidencia = patron.search(contenido)
        if not coincidencia:
            continue
        estilo = re.search(r'\bs="(\d+)"', coincidencia.group(1))
        attr_estilo = f' s="{estilo.group(1)}"' if estilo else ""
        nueva = f'<c r="{col_jefatura}{numero}"{attr_estilo} t="s"><v>{idx}</v></c>'
        contenido = contenido[:coincidencia.start()] + nueva + contenido[coincidencia.end():]
        cambios.append({
            "fila": str(numero),
            "codigo": codigo,
            "oficina": oficina,
            "jefatura_anterior": actual,
            "jefatura_nueva": objetivo,
        })
    if cambios:
        total = str(len(strings))
        strings_texto = re.sub(r'(<sst[^>]*?\bcount=")\d+(")', r"\g<1>%s\g<2>" % total, strings_texto, count=1)
        strings_texto = re.sub(r'(<sst[^>]*?\buniqueCount=")\d+(")', r"\g<1>%s\g<2>" % total, strings_texto, count=1)
    return contenido, strings_texto, cambios


def procesar(archivo: Path, hojas: list[tuple[str, str, str, str]], destino: Path) -> list[dict[str, str]]:
    if not archivo.exists():
        return []
    mapa = leer_fuente()
    with ZipFile(archivo) as zf:
        nombres = mapa_hoja_a_xml(zf)
        strings_texto, strings = leer_strings(zf)
        nuevos: dict[str, str] = {}
        todos: list[dict[str, str]] = []
        for hoja, col_cod, col_ofi, col_jef in hojas:
            sheet_xml = nombres[hoja]
            filas = leer_celdas(zf, sheet_xml, strings, {col_cod, col_ofi, col_jef})
            contenido = zf.read(sheet_xml).decode("utf-8")
            contenido, strings_texto, cambios = actualizar_hoja(
                contenido, strings_texto, strings, filas, col_cod, col_ofi, col_jef, mapa)
            nuevos[sheet_xml] = contenido
            for c in cambios:
                c["hoja"] = hoja
                todos.append(c)
        if destino is not None and destino != archivo:
            destino.parent.mkdir(parents=True, exist_ok=True)
        destino_final = destino if destino is not None else archivo
        temp = destino_final.with_suffix(".tmp.xlsx")
        with ZipFile(temp, "w", ZIP_DEFLATED) as out:
            for item in zf.infolist():
                if item.filename == "xl/sharedStrings.xml":
                    data = strings_texto.encode("utf-8")
                elif item.filename in nuevos:
                    data = nuevos[item.filename].encode("utf-8")
                else:
                    data = zf.read(item.filename)
                out.writestr(item, data)
        shutil.move(str(temp), str(destino_final))
    return todos


def main() -> None:
    SALIDA.mkdir(parents=True, exist_ok=True)
    mapa = leer_fuente()
    print(f"Jefaturas detectadas en la fuente: {sum(len(v) for v in mapa.values())} "
          f"en {len(mapa)} codigos")
    for codigo in sorted(mapa):
        print(f"  {codigo:>5} -> {', '.join(mapa[codigo])}")

    registro: list[dict[str, str]] = []

    cambios = procesar(
        OBJETIVOS[0][0],
        [(OBJETIVOS[0][1], OBJETIVOS[0][2], OBJETIVOS[0][3], OBJETIVOS[0][4])],
        OBJETIVOS[0][5],
    )
    for c in cambios:
        c["archivo"] = OBJETIVOS[0][0].name
    registro += cambios

    control = OBJETIVOS[1][0]
    cambios = procesar(
        control,
        [("Hoja1", "A", "C", "F"), ("Hoja3", "C", "E", "H")],
        control,
    )
    for c in cambios:
        c["archivo"] = control.name
    registro += cambios

    with (SALIDA / "mapeo_jefaturas.csv").open("w", newline="", encoding="utf-8-sig") as f:
        campos = ["archivo", "hoja", "fila", "codigo", "oficina",
                  "jefatura_anterior", "jefatura_nueva"]
        w = csv.DictWriter(f, fieldnames=campos)
        w.writeheader()
        w.writerows(registro)

    print(f"\nCambios aplicados: {len(registro)}")
    for c in registro:
        print(f"  {c['archivo']} [{c['hoja']}] fila {c['fila']}: "
              f"{c['jefatura_anterior']} -> {c['jefatura_nueva']}")


if __name__ == "__main__":
    main()
