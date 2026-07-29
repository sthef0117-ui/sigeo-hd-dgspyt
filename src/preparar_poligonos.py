"""
SIGEO-HD DGSPYT — Preparacion de la cartografia propia.

Toma los limites municipales del Estado de Mexico y produce un GeoJSON
simplificado y con los nombres normalizados, listo para dibujarse en el
tablero sin depender de ningun proveedor de mosaicos.

    python src/preparar_poligonos.py <archivo_origen.json>

Por que dibujar los poligonos en lugar de usar mosaicos: los mosaicos de
Microsoft exigen llave de suscripcion, y los de OpenStreetMap o Esri obligan a
mostrar su credito en el mapa. Con cartografia propia no hay proveedor, no hay
credito de terceros y el mapa se parece al mural de la Direccion.

La simplificacion usa Douglas-Peucker. A 200 m de tolerancia el contorno
municipal se ve identico en pantalla y el archivo baja de 585 KB a una fraccion.
"""

import json
import math
import sys
import unicodedata
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
SALIDA = RAIZ / "poligonos" / "municipios_edomex.geojson"

# Tolerancia en grados. 0.0018 grados son unos 200 m a esta latitud.
TOLERANCIA = 0.0018

SUFIJOS = (" DE MORELOS", " DE JUAREZ", " DE ZARAGOZA", " DE BAZ", " SOLIDARIDAD")

# La fuente de limites trae erratas de captura y grafias antiguas. Se corrigen
# aqui para que el poligono cruce con el nombre que usan los insumos del C5.
CORRECCIONES = {
    "SAN MARTIN DE LAS PIRAAMIDES": "SAN MARTIN DE LAS PIRAMIDES",
    "ZINACATEPEC": "ZINACANTEPEC",
    "JALATLACO": "XALATLACO",
    "TLALNEPANTLA": "TLALNEPANTLA",
}


def norm(texto):
    s = unicodedata.normalize("NFD", str(texto or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return " ".join(s.upper().split())


def canonico(nombre):
    """Mismo criterio que municipio_canonico() del pipeline, para que crucen."""
    n = norm(nombre)
    n = CORRECCIONES.get(n, n)
    for sufijo in SUFIJOS:
        if n.endswith(sufijo):
            return n[: -len(sufijo)].strip()
    return n


def _distancia_perpendicular(p, a, b):
    (px, py), (ax, ay), (bx, by) = p, a, b
    dx, dy = bx - ax, by - ay
    if dx == 0 and dy == 0:
        return math.hypot(px - ax, py - ay)
    t = ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)
    t = max(0.0, min(1.0, t))
    return math.hypot(px - (ax + t * dx), py - (ay + t * dy))


def simplificar(puntos, tol):
    """Douglas-Peucker iterativo: conserva la silueta y tira los vertices planos."""
    if len(puntos) < 3:
        return puntos
    conservar = [False] * len(puntos)
    conservar[0] = conservar[-1] = True
    pila = [(0, len(puntos) - 1)]
    while pila:
        ini, fin = pila.pop()
        peor, indice = 0.0, -1
        for i in range(ini + 1, fin):
            d = _distancia_perpendicular(puntos[i], puntos[ini], puntos[fin])
            if d > peor:
                peor, indice = d, i
        if peor > tol:
            conservar[indice] = True
            pila.append((ini, indice))
            pila.append((indice, fin))
    return [p for p, c in zip(puntos, conservar) if c]


def simplificar_anillo(anillo, tol):
    cerrado = len(anillo) > 2 and anillo[0] == anillo[-1]
    puntos = [tuple(p[:2]) for p in anillo]
    salida = simplificar(puntos, tol)
    if cerrado and salida[0] != salida[-1]:
        salida.append(salida[0])
    # Un anillo necesita al menos cuatro puntos para seguir siendo un poligono.
    return salida if len(salida) >= 4 else [tuple(p[:2]) for p in anillo]


def redondear(punto):
    # Cinco decimales son ~1 m: mas precision solo engorda el archivo.
    return [round(punto[0], 5), round(punto[1], 5)]


def poligonos_de(geometria):
    """Normaliza Polygon, MultiPolygon y GeometryCollection a lista de poligonos."""
    if not geometria:
        return []
    tipo = geometria.get("type")
    if tipo == "Polygon":
        return [geometria["coordinates"]]
    if tipo == "MultiPolygon":
        return list(geometria["coordinates"])
    if tipo == "GeometryCollection":
        salida = []
        for g in geometria.get("geometries", []):
            salida += poligonos_de(g)
        return salida
    return []


def main():
    origen = Path(sys.argv[1]) if len(sys.argv) > 1 else None
    if not origen or not origen.exists():
        print("Uso: python src/preparar_poligonos.py <limites_municipales.json>")
        return 1

    datos = json.loads(origen.read_text(encoding="utf-8"))
    rasgos, antes, despues = [], 0, 0

    for f in datos.get("features", []):
        props = f.get("properties", {})
        nombre = (props.get("NAME_2") or props.get("name")
                  or props.get("NOMGEO") or props.get("municipio") or "")
        if not nombre:
            continue

        poligonos = []
        for poligono in poligonos_de(f.get("geometry")):
            anillos = []
            for anillo in poligono:
                antes += len(anillo)
                simple = simplificar_anillo(anillo, TOLERANCIA)
                despues += len(simple)
                anillos.append([redondear(p) for p in simple])
            if anillos:
                poligonos.append(anillos)
        if not poligonos:
            continue

        geometria = ({"type": "Polygon", "coordinates": poligonos[0]}
                     if len(poligonos) == 1
                     else {"type": "MultiPolygon", "coordinates": poligonos})
        rasgos.append({
            "type": "Feature",
            "properties": {"municipio": canonico(nombre), "nombre": nombre},
            "geometry": geometria,
        })

    SALIDA.parent.mkdir(parents=True, exist_ok=True)
    SALIDA.write_text(json.dumps({"type": "FeatureCollection", "features": rasgos},
                                 ensure_ascii=False, separators=(",", ":")),
                      encoding="utf-8")

    print("SIGEO-HD DGSPYT · cartografía propia")
    print(f"  municipios ................. {len(rasgos)}")
    print(f"  vértices ................... {antes:,} → {despues:,} "
          f"({100 - despues * 100 // max(antes, 1)}% menos)")
    print(f"  archivo .................... {SALIDA.stat().st_size / 1024:.0f} KB")
    return 0


if __name__ == "__main__":
    sys.exit(main())
