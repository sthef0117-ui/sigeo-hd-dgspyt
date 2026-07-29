import os
import glob
import json
import re
import pandas as pd
from datetime import datetime
from docx import Document

base_dir = r"C:\Users\xXZaB\.gemini\antigravity\scratch\sigeo_hd_dgspyt"
wa_dir = os.path.join(base_dir, "insumos", "whatsapp")
historico_dir = os.path.join(base_dir, "historico_c5")
ti_dir = os.path.join(historico_dir, "tarjetas_informativas")

master_ti_excel = os.path.join(ti_dir, "CONCENTRADO_MASTER_TARJETAS_INFORMATIVAS.xlsx")
master_911_excel = os.path.join(historico_dir, "concentrados", "CONCENTRADO_HISTORICO_LLAMADAS_911_089_C5.xlsx")

os.makedirs(os.path.join(ti_dir, "texto"), exist_ok=True)
os.makedirs(os.path.join(ti_dir, "imagenes"), exist_ok=True)

print("=========================================================================")
print("  SIGEO-HD DGSPYT: PARSEADOR COMPLETO DE TARJETAS DE WHATSAPP E INSUMOS   ")
print("=========================================================================")

# Load 911 master calls dataset
df_911 = pd.DataFrame()
if os.path.exists(master_911_excel):
    try:
        df_911 = pd.read_excel(master_911_excel, sheet_name='CONCENTRADO_911')
    except Exception as e:
        print(f"Error cargando 911 Master: {e}")

cards_extracted = []

# 1. Parse Excel Corroborados HD
excel_hd = os.path.join(wa_dir, "ACCIONES DE HD JULIO DGSPYT.xlsx")
if os.path.exists(excel_hd):
    df_corr = pd.read_excel(excel_hd, sheet_name='CORROBORADOS')
    for idx, r in df_corr.iterrows():
        cards_extracted.append({
            "origen": "EXCEL_CORROBORADOS",
            "municipio": str(r.get('MUNICIPIO', '')).upper(),
            "victima": str(r.get('POSIBLE VÍCTIMA') or r.get('POSIBLE VCTIMA') or 'DESCONOCIDO').upper(),
            "edad": None,
            "casquillos": "No especificado",
            "policia_unidad": "SSEM DGSPYT",
            "ubicacion": f"{r.get('CALLE', '')}, Col. {r.get('COLONIA', '')}",
            "hora": str(r.get('HORA', '')),
            "narrativa_completa": str(r.get('DESARROLLO DE LOS HECHOS', '')),
            "acciones_ssem": str(r.get('ACCIONES SSEM', ''))
        })

# 2. Parse Word Document INFORME_911_VIOLENCIA...
docx_file = os.path.join(wa_dir, "INFORME_911_VIOLENCIA_21-22_JUL_2026.docx")
if os.path.exists(docx_file):
    doc = Document(docx_file)
    full_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    blocks = re.split(r'\n(?=[A-ZÁÉÍÓÚÑ\s]{3,30}\s*\()', full_text)
    for b in blocks:
        if len(b.strip()) > 50:
            mun_match = re.search(r'([A-ZÁÉÍÓÚÑ\s]+)\s*\(', b)
            mun = mun_match.group(1).strip().upper() if mun_match else "POR CLASIFICAR"
            cards_extracted.append({
                "origen": "INFORME_DOCX",
                "municipio": mun,
                "victima": "DESCONOCIDO",
                "edad": None,
                "casquillos": "No especificado",
                "policia_unidad": "C5 / SSEM",
                "ubicacion": "Ver narrativa",
                "hora": "No especificada",
                "narrativa_completa": b.strip()[:400],
                "acciones_ssem": "Atención C5"
            })

# 3. Add explicit WhatsApp cards from user screenshots (Ecatepec & Tecámac)
cards_extracted.append({
    "origen": "WHATSAPP_JEFE_MIGUEL",
    "municipio": "ECATEPEC",
    "victima": "RICARDO HERNÁNDEZ MAYA REYNOSA",
    "edad": 33,
    "casquillos": "5 casquillos 9mm",
    "policia_unidad": "GUTIERREZ RAMIREZ JUAN CARLOS (ME334A)",
    "ubicacion": "calle Colorines esquina calle Encino, colonia Tierra Blanca",
    "hora": "10:55",
    "narrativa_completa": "A las 10:55 horas, el Policía GUTIERREZ RAMIREZ JUAN CARLOS unidad ME334A, informó que en calle Colorines esquina calle Encino, colonia Tierra Blanca...",
    "acciones_ssem": "Primer Respondiente Policía Municipal FERMÍN GUZMÁN FERNANDO (sector 3-21). PC-2021 certifica deceso."
})

cards_extracted.append({
    "origen": "WHATSAPP_JEFE_MIGUEL",
    "municipio": "TECÁMAC",
    "victima": "LUIS EDUARDO RUIZ ROJAS",
    "edad": 35,
    "casquillos": "7 casquillos 9mm",
    "policia_unidad": "MALDONADO SUÁREZ JUAN TRINIDAD (ME937A5)",
    "ubicacion": "calle Jardines de La Viña esquina 2da. Cerrada, Héroes Tecámac",
    "hora": "23:55",
    "narrativa_completa": "A las 23:55 horas, el Policía MALDONADO SUÁREZ JUAN TRINIDAD, unidad ME937A5, informó que a las 06:15 horas, en calle Jardines de La Viña... vehículo Chevrolet Chevy Monza blanco taxi placas PUB383E...",
    "acciones_ssem": "Primer Respondiente NORMA HERNÁNDEZ CASTRO (C-072). Unidad Médica PC-70."
})

print(f"Total tarjetas brutas procesadas de todos los insumos: {len(cards_extracted)}")

# Deduplicate and Cross-Reference
dedup_cards = []
seen_signatures = set()

for idx, card in enumerate(cards_extracted):
    v = card['victima']
    m = card['municipio']
    
    # Signature for deduplication
    sig = f"{m}_{v}" if v != "DESCONOCIDO" else f"{m}_{card['narrativa_completa'][:30]}"
    if sig in seen_signatures:
        continue
    seen_signatures.add(sig)

    # Cross Reference with 911 Calls
    matched_call = None
    if not df_911.empty and m != "DESCONOCIDO":
        filtered = df_911[df_911['MUNICIPIO'].astype(str).str.upper().str.contains(m, na=False)]
        for _, row in filtered.iterrows():
            notes = str(row.get('NOTAS', '')).upper()
            addr = str(row.get('DIRECCION ', '')).upper()
            full_text = f"{notes} {addr}"

            if (v != "DESCONOCIDO" and v in full_text) or \
               (card['ubicacion'] != "No especificada" and any(w in full_text for w in card['ubicacion'].upper().split() if len(w) > 4)):
                matched_call = str(row.get('FOLIO', ''))
                break

    if matched_call:
        card['folio_911_vinculado'] = matched_call
        card['coincidencia_estatus'] = "VINCULADA 911 C5"
    else:
        card['folio_911_vinculado'] = "N/A"
        card['coincidencia_estatus'] = "SIN REGISTRO PREVIO 911"

    card['id_tarjeta'] = f"TI-MASIVA-{idx+1:03d}"
    card['fecha_registro'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    dedup_cards.append(card)

print(f"Total Tarjetas Informativas únicas consolidadas deduplicadas: {len(dedup_cards)}")

# Write to Master Excel
df_ti_final = pd.DataFrame(dedup_cards)
with pd.ExcelWriter(master_ti_excel, engine='openpyxl') as writer:
    df_ti_final.to_excel(writer, sheet_name='TARJETAS_INFORMATIVAS', index=False)

print(f"[OK] Concentrado Master de Tarjetas Informativas actualizado exitosamente en: {master_ti_excel}")
