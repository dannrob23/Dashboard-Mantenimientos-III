# -*- coding: utf-8 -*-
"""
descargar_geojson.py — Descarga el GeoJSON de departamentos de Colombia para
activar el mapa choropleth (sin este archivo el mapa funciona en modo burbuja).

Ejecutar UNA VEZ en su maquina (con internet) dentro de la carpeta del proyecto:
    python descargar_geojson.py

El archivo se guarda en public/colombia.geojson y se versiona en el repo.
Fuentes (datos abiertos / dominio publico):
  - https://github.com/john-guerra/colombia_geojson   (departamentos)
  - https://github.com/santiblanko/colombia.geojson   (departamentos y municipios)
"""
import os
import sys
import urllib.request

AQUI = os.path.dirname(os.path.abspath(__file__))
DESTINO = os.path.join(AQUI, "public", "colombia.geojson")

FUENTES = [
    "https://raw.githubusercontent.com/john-guerra/colombia_geojson/master/geojson/departamentos.geojson",
    "https://raw.githubusercontent.com/santiblanko/colombia.geojson/master/colombia.geojson",
]


def descargar(url, destino):
    print(f"Descargando {url} ...")
    req = urllib.request.Request(url, headers={"User-Agent": "pipeline-dashboard"})
    with urllib.request.urlopen(req, timeout=60) as r, open(destino, "wb") as fh:
        fh.write(r.read())
    kb = os.path.getsize(destino) / 1024
    if kb < 20:
        os.remove(destino)
        return False
    print(f"OK -> {destino} ({kb:.0f} KB)")
    return True


def main():
    os.makedirs(os.path.dirname(DESTINO), exist_ok=True)
    for url in FUENTES:
        try:
            if descargar(url, DESTINO):
                return
        except Exception as e:
            print(f"  fallo: {e}")
    sys.exit("No se pudo descargar el GeoJSON. Verifique su conexion o descarguelo "
             "manualmente y guardelo como public/colombia.geojson")


if __name__ == "__main__":
    main()
