"""
SIGEO-HD DGSPYT — Ensamblador del tablero.

Inyecta los JSON calculados por etl_sigeo.py dentro de la plantilla
src/dashboard_template.html y escribe index.html en la raiz del proyecto.

    python src/build_dashboard.py

El resultado es un archivo unico y autocontenido: se abre desde GitHub Pages
o directamente desde una memoria USB en la sala de juntas del C5, sin
servidor ni conexion a los JSON.
"""

import json
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
ANALISIS = RAIZ / "analisis"
PLANTILLA = Path(__file__).parent / "dashboard_template.html"
SALIDA = RAIZ / "index.html"

MARCA = "/*__DATOS__*/ null"

ARCHIVOS = {
    "resumen": "resumen_ejecutivo.json",
    "hd": "corroborados_sigeo.json",
    "llamadas": "llamadas_911_sigeo.json",
    "bases": "bases_dgspyt.json",
    "zonas": "zonas_ciegas.json",
    "auditoria": "auditoria_decesos.json",
    "territorio": "perfil_territorial.json",
    "coordinaciones": "perfil_coordinaciones.json",
    "serie": "serie_temporal.json",
    "cruce": "cruce_hechos_fatales.json",
    "perimetro": "perimetro_edomex.json",
    "municipios": "municipios_edomex.json",
}


def main():
    faltantes = [n for n in ARCHIVOS.values() if not (ANALISIS / n).exists()]
    if faltantes:
        print("Faltan insumos calculados: " + ", ".join(faltantes))
        print("Ejecuta primero:  python src/etl_sigeo.py")
        return 1

    datos = {clave: json.loads((ANALISIS / nombre).read_text(encoding="utf-8"))
             for clave, nombre in ARCHIVOS.items()}

    # El tablero solo necesita las bases georreferenciadas y en uso.
    datos["bases"] = [b for b in datos["bases"] if b["lat"] and b["en_uso"]]

    plantilla = PLANTILLA.read_text(encoding="utf-8")
    if MARCA not in plantilla:
        print(f"La plantilla no contiene el marcador {MARCA}")
        return 1

    # separators compacto: el archivo viaja en USB y se sirve por GitHub Pages.
    carga = json.dumps(datos, ensure_ascii=False, separators=(",", ":"))
    # </script> dentro de una cadena JSON cerraria el bloque antes de tiempo.
    carga = carga.replace("</", "<\\/")

    SALIDA.write_text(plantilla.replace(MARCA, carga), encoding="utf-8")

    mb = SALIDA.stat().st_size / 1_048_576
    print("SIGEO-HD DGSPYT · tablero ensamblado")
    for clave in ARCHIVOS:
        print(f"  {clave:<10} {len(datos[clave]) if isinstance(datos[clave], list) else 1:>6} registros")
    print(f"  salida ..... {SALIDA}  ({mb:.2f} MB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
