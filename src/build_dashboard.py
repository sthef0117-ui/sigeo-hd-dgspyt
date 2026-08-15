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
import re
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
ANALISIS = RAIZ / "analisis"
PLANTILLA = Path(__file__).parent / "dashboard_template.html"
SALIDA = RAIZ / "index.html"

# El tablero web es GitHub Pages (público con cortina de contraseña). Las
# narrativas de las tarjetas traen nombres de víctima y de servidores públicos.
# Aquí se enmascaran SOLO para la copia web; la base local y los productos
# internos (Word, lista, Power BI) conservan el texto íntegro. No se modifica
# analisis/tarjetas_whatsapp.json: el enmascarado se hace en memoria al armar
# index.html.
_MARCA_NOM = "[NOMBRE RESERVADO]"
_NOMBRE = r"[A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ.]+(?:\s+[A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ.]+){1,4}"
_REGLAS_WEB = [
    re.compile(rf"(respond[ií]a al nombre de|responde al nombre de|de nombre|"
               rf"identificad[oa] como)\s+{_NOMBRE}", re.IGNORECASE),
    re.compile(rf"(Primer Respondiente(?:\s+Polic[ií]a\s+Municipal)?)\s+{_NOMBRE}",
               re.IGNORECASE),
    re.compile(rf"(Polic[ií]a(?:\s+(?:Tercero|Primero|Segundo|Municipal|Estatal|"
               rf"Auxiliar))*)\s+{_NOMBRE}", re.IGNORECASE),
    re.compile(rf"(proporcionad[oa]s?\s+por\s+(?:su\s+\w+\s+)?)\s*{_NOMBRE}",
               re.IGNORECASE),
    re.compile(rf"(su\s+(?:novi[ao]|espos[ao]|hij[ao]|herman[ao]|madre|padre|"
               rf"pareja|ti[ao]|sobrin[ao]|niet[ao]|cuñad[ao])\s+){_NOMBRE}",
               re.IGNORECASE),
    re.compile(rf"(al mando del?\s+(?:Param[ée]dic[oa]|Comandante|Oficial)?\s*)"
               rf"{_NOMBRE}", re.IGNORECASE),
]


def enmascarar_nombres_web(texto):
    if not texto:
        return texto
    s = str(texto)
    for patron in _REGLAS_WEB:
        s = patron.sub(lambda m: f"{m.group(1)} {_MARCA_NOM}", s)
    return s


def tarjetas_web(tarjetas):
    """Copia de las tarjetas con la narrativa sin nombres, para la web."""
    salida = []
    for t in tarjetas:
        w = dict(t)
        w["narrativa"] = enmascarar_nombres_web(t.get("narrativa", ""))
        salida.append(w)
    return salida

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
    "fichas": "fichas_inteligencia.json",
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

    # Tarjetas de WhatsApp: homicidios reales reportados por el chat institucional.
    # Es opcional (se genera con ingesta_whatsapp.py); si no está, va vacío.
    ruta_wa = ANALISIS / "tarjetas_whatsapp.json"
    tarjetas = (json.loads(ruta_wa.read_text(encoding="utf-8"))
                if ruta_wa.exists() else [])
    # A la web van sin nombres; el JSON interno queda íntegro.
    datos["tarjetas"] = tarjetas_web(tarjetas)

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
