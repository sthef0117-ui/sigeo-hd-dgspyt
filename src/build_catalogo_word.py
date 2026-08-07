"""
SIGEO-HD DGSPYT — Catálogo territorial en Word editable.

Escribe el contenido del mapa mural de la Dirección como texto: la estructura
Coordinación Regional -> Subdirección -> Municipios, con las cifras del corte.
El mapa no se incluye como imagen; lo que va es su texto, para poder editarlo.

    python src/etl_sigeo.py
    python src/build_catalogo_word.py

Salida: entregables/CATALOGO_TERRITORIAL_DGSPYT.docx

Documento editable de verdad: estilos normales de Word, tablas reales y sin
imágenes ni cajas de texto flotantes, para que el área pueda corregir un
municipio mal ubicado sin pelearse con el formato.
"""

import json
import sys
from collections import defaultdict
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

RAIZ = Path(__file__).resolve().parent.parent
ANALISIS = RAIZ / "analisis"
SALIDA = RAIZ / "entregables"

GUINDA = RGBColor(0x7A, 0x13, 0x27)
TINTA = RGBColor(0x16, 0x20, 0x2C)
GRIS = RGBColor(0x59, 0x65, 0x73)

MENORES = {"de", "del", "la", "las", "el", "los", "en", "y", "con", "por", "a", "al"}


def titulo_es(t):
    palabras = str(t).lower().split()
    return " ".join(p.capitalize() if i == 0 or p not in MENORES else p
                    for i, p in enumerate(palabras))


def leer(nombre):
    ruta = ANALISIS / nombre
    if not ruta.exists():
        print(f"Falta {ruta.name}. Ejecuta primero: python src/etl_sigeo.py")
        sys.exit(1)
    return json.loads(ruta.read_text(encoding="utf-8"))


def sombrear(celda, hex_color):
    sombra = OxmlElement("w:shd")
    sombra.set(qn("w:val"), "clear")
    sombra.set(qn("w:fill"), hex_color)
    celda._tc.get_or_add_tcPr().append(sombra)


def parrafo(doc, texto, tam=10.5, negrita=False, color=TINTA, espacio=6,
            alineacion=WD_ALIGN_PARAGRAPH.LEFT, cursiva=False):
    p = doc.add_paragraph()
    p.alignment = alineacion
    p.paragraph_format.space_after = Pt(espacio)
    r = p.add_run(texto)
    r.font.size = Pt(tam)
    r.font.bold = negrita
    r.font.italic = cursiva
    r.font.color.rgb = color
    r.font.name = "Calibri"
    return p


def main():
    territorio = leer("perfil_territorial.json")
    coord = leer("perfil_coordinaciones.json")["coordinaciones"]
    R = leer("resumen_ejecutivo.json")

    doc = Document()
    for seccion in doc.sections:
        seccion.top_margin = seccion.bottom_margin = Cm(2)
        seccion.left_margin = seccion.right_margin = Cm(2.2)

    estilo = doc.styles["Normal"]
    estilo.font.name = "Calibri"
    estilo.font.size = Pt(10.5)

    # ---------------- Portada ----------------
    parrafo(doc, "DIRECCIÓN GENERAL DE SEGURIDAD PÚBLICA Y TRÁNSITO",
            tam=11, negrita=True, color=GUINDA, espacio=2)
    parrafo(doc, "Unidad de Homicidios Dolosos", tam=10, color=GRIS, espacio=14)

    t = doc.add_heading("Catálogo territorial por coordinación regional", level=0)
    for r in t.runs:
        r.font.color.rgb = TINTA
    parrafo(doc, "Estructura del mapa mural en texto editable · "
                 f"cifras del corte de julio 2026 · generado {R['generado']}",
            tam=9.5, color=GRIS, espacio=16, cursiva=True)

    doc.add_heading("Cómo se construyó este catálogo", level=1)
    parrafo(doc,
            "La estructura Coordinación Regional → Subdirección → Municipios proviene "
            "del inventario oficial de instalaciones de la Dirección "
            "(INSTALACIONES_DGSPYT), que declara la coordinación y la subdirección de "
            "cada inmueble. Es la misma división del mapa mural.")
    parrafo(doc,
            "El inventario es de inmuebles, no de territorio: los municipios sin "
            "instalación propia no traen coordinación declarada. Esos se completaron "
            "de dos formas y quedan marcados en el documento:")

    for marca, glosa in [
        ("Sin marca", "coordinación declarada en el inventario oficial."),
        ("(v)", "asignado por vecindad territorial: se tomó la coordinación "
                "mayoritaria entre sus municipios colindantes."),
        ("(e)", "heredado del municipio del que se escindió, por ser de creación "
                "posterior al levantamiento cartográfico."),
    ]:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(3)
        r1 = p.add_run(f"{marca} — ")
        r1.font.bold = True
        r1.font.size = Pt(10.5)
        r2 = p.add_run(glosa)
        r2.font.size = Pt(10.5)

    parrafo(doc,
            "Las asignaciones marcadas (v) y (e) son inferencias, no dato oficial. "
            "Conviene que el área de planeación las confirme o corrija: este documento "
            "es editable justamente para eso.",
            tam=10, color=GUINDA, espacio=14)

    # ---------------- Resumen ----------------
    doc.add_heading("Resumen por coordinación", level=1)
    columnas = ["Coordinación regional", "Municipios", "Homicidios",
                "Llamadas con violencia", "Bases en uso", "Personal adscrito"]
    tabla = doc.add_table(rows=1, cols=len(columnas))
    tabla.style = "Table Grid"
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, c in enumerate(columnas):
        celda = tabla.rows[0].cells[j]
        celda.text = ""
        p = celda.paragraphs[0]
        r = p.add_run(c)
        r.font.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        if j:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        sombrear(celda, "16202C")

    for c in coord:
        fila = tabla.add_row().cells
        valores = [titulo_es(c["coordinacion"]), c["total_municipios"],
                   c["hd_eventos"], f"{c['llamadas_violentas']:,}",
                   c["bases_en_uso"], f"{c['personal_total']:,}"]
        for j, v in enumerate(valores):
            fila[j].text = ""
            p = fila[j].paragraphs[0]
            r = p.add_run(str(v))
            r.font.size = Pt(9.5)
            r.font.bold = (j == 0)
            if j:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    fila = tabla.add_row().cells
    totales = ["TOTAL", sum(c["total_municipios"] for c in coord),
               sum(c["hd_eventos"] for c in coord),
               f"{sum(c['llamadas_violentas'] for c in coord):,}",
               sum(c["bases_en_uso"] for c in coord),
               f"{sum(c['personal_total'] for c in coord):,}"]
    for j, v in enumerate(totales):
        fila[j].text = ""
        p = fila[j].paragraphs[0]
        r = p.add_run(str(v))
        r.font.size = Pt(9.5)
        r.font.bold = True
        if j:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        sombrear(fila[j], "EEF1F4")

    doc.add_page_break()

    # ---------------- Detalle ----------------
    doc.add_heading("Detalle por coordinación y subdirección", level=1)

    por_coord = defaultdict(lambda: defaultdict(list))
    datos = {t["municipio"]: t for t in territorio}
    for t in territorio:
        sub = t["subdireccion"] or "Sin subdirección declarada"
        por_coord[t["coordinacion"]][sub].append(t["municipio"])

    marca_de = {"vecindad": " (v)", "escision": " (e)"}
    orden = [c["coordinacion"] for c in coord]

    for nombre in orden:
        subs = por_coord.get(nombre, {})
        total = sum(len(v) for v in subs.values())
        doc.add_heading(f"{titulo_es(nombre)} — {total} municipios", level=2)

        ficha = next((c for c in coord if c["coordinacion"] == nombre), None)
        if ficha:
            parrafo(doc,
                    f"{ficha['hd_eventos']} homicidios dolosos · "
                    f"{ficha['llamadas_violentas']:,} llamadas con violencia · "
                    f"{ficha['bases_en_uso']} bases en uso con "
                    f"{ficha['personal_total']:,} elementos adscritos · "
                    f"carga por 100 elementos: {ficha['carga_por_100_elementos']}",
                    tam=9.5, color=GRIS, espacio=8, cursiva=True)

        for sub in sorted(subs, key=lambda s: (s.startswith("Sin"), s)):
            municipios = sorted(subs[sub])
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(2)
            r = p.add_run(f"{titulo_es(sub)} ({len(municipios)})")
            r.font.bold = True
            r.font.size = Pt(10)
            r.font.color.rgb = GUINDA

            for m in municipios:
                d = datos[m]
                p = doc.add_paragraph(style="List Bullet")
                p.paragraph_format.space_after = Pt(0)
                r = p.add_run(titulo_es(m) + marca_de.get(d["coordinacion_origen"], ""))
                r.font.size = Pt(10)
                if d["hd_eventos"]:
                    extra = p.add_run(f"  —  {d['hd_eventos']} homicidio"
                                      f"{'s' if d['hd_eventos'] > 1 else ''}")
                    extra.font.size = Pt(9)
                    extra.font.color.rgb = GUINDA
                    extra.font.bold = True

    # ---------------- Cierre ----------------
    doc.add_page_break()
    doc.add_heading("Notas sobre el alcance", level=1)
    for nota in [
        "Las cifras de homicidios corresponden a los casos corroborados por la "
        "Dirección en el corte de julio de 2026. Las de llamadas provienen del "
        "concentrado del C5 del 22 al 28 de julio de 2026.",
        "Los datos de coordinación miden condiciones del territorio, no el desempeño "
        "de una persona: los insumos no contienen asignación nominal de mando, turnos "
        "ni recorridos.",
        "Este catálogo se regenera con «python src/build_catalogo_word.py» después de "
        "actualizar el pipeline, de modo que las cifras nunca se escriben a mano.",
    ]:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(5)
        r = p.add_run(nota)
        r.font.size = Pt(10)

    SALIDA.mkdir(parents=True, exist_ok=True)
    destino = SALIDA / "CATALOGO_TERRITORIAL_DGSPYT.docx"
    doc.save(destino)

    inferidos = sum(1 for t in territorio if t["coordinacion_origen"] != "catalogo")
    print("SIGEO-HD DGSPYT · catálogo territorial en Word")
    print(f"  coordinaciones ............. {len(coord)}")
    print(f"  municipios ................. {len(territorio)} "
          f"({inferidos} marcados como inferidos)")
    print(f"  archivo .................... {destino}")
    print(f"  tamaño ..................... {destino.stat().st_size / 1024:.0f} KB")
    return 0


if __name__ == "__main__":
    sys.exit(main())
