import os
import glob
import json
import re
import pandas as pd
from datetime import datetime

base_dir = r"C:\Users\xXZaB\.gemini\antigravity\scratch\sigeo_hd_dgspyt"
wa_dir = os.path.join(base_dir, "insumos", "whatsapp")
historico_dir = os.path.join(base_dir, "historico_c5")
ti_dir = os.path.join(historico_dir, "tarjetas_informativas")

master_ti_excel = os.path.join(ti_dir, "CONCENTRADO_MASTER_TARJETAS_INFORMATIVAS.xlsx")
master_911_excel = os.path.join(historico_dir, "concentrados", "CONCENTRADO_HISTORICO_LLAMADAS_911_089_C5.xlsx")
out_super_fichas_html = os.path.join(historico_dir, "pdf_reportes", "SUPER_FICHAS_360_VINCULADAS.html")

os.makedirs(os.path.join(ti_dir, "texto"), exist_ok=True)
os.makedirs(os.path.join(ti_dir, "imagenes"), exist_ok=True)

print("=========================================================================")
print("  SIGEO-HD DGSPYT: INGESTIÓN MASIVA Y CRUCE DE TODAS LAS TARJETAS WHATSAPP ")
print("=========================================================================")

# 1. Read WhatsApp Chat Log txt file
chat_file = os.path.join(wa_dir, "Chat de WhatsApp con C5.txt")
raw_messages = []

if os.path.exists(chat_file):
    with open(chat_file, 'r', encoding='utf-8', errors='ignore') as f:
        chat_content = f.read()
    
    # Split messages by timestamp pattern
    msg_blocks = re.split(r'\d{1,2}/\d{1,2}/\d{2,4},\s*\d{1,2}:\d{2}\s*-\s*', chat_content)
    print(f"Total mensajes parseados del chat de WhatsApp: {len(msg_blocks)}")

    # Filter blocks containing relevant card keywords
    card_keywords = ["PERSONA MUERTA", "HOMICIDIO", "CADAVER", "IMPACTOS DE BALA", "CASQUILLOS", "OCCISO", "DISPARO", "SUICIDIO"]
    for b in msg_blocks:
        if any(kw in b.upper() for kw in card_keywords) and len(b.strip()) > 80:
            raw_messages.append(b.strip())

print(f"Tarjetas Informativas encontradas en el chat: {len(raw_messages)}")

# Load Master 911 Calls Dataset for Cross-Referencing
df_911 = pd.DataFrame()
if os.path.exists(master_911_excel):
    try:
        df_911 = pd.read_excel(master_911_excel, sheet_name='CONCENTRADO_911')
        print(f"Concentrado Máster 911 cargado con {len(df_911)} folios.")
    except Exception as e:
        print(f"Error cargando Master 911: {e}")

def parse_card_text(text_block):
    # Municipality
    mun_match = re.search(r'([A-ZÁÉÍÓÚÑ\s]+)\s*\((?:PERSONA|HOMICIDIO|CADAVER|SUICIDIO|LESIONADO)', text_block, re.IGNORECASE)
    municipio = mun_match.group(1).strip().upper() if mun_match else "DESCONOCIDO"

    if municipio == "DESCONOCIDO":
        muns = ["ECATEPEC", "NAUCALPAN", "CUAUTITLÁN IZCALLI", "NEZAHUALCÓYOTL", "CHIMALHUACÁN", "TULTITLÁN", "TOLUCA", "CHALCO", "IXTAPALUCA", "TECÁMAC"]
        for m in muns:
            if m.lower() in text_block.lower():
                municipio = m
                break

    # Victim Name
    victim_match = re.search(r'nombre de\s+([A-ZÁÉÍÓÚÑ\s]{5,40})(?:\s+de\s+\d+|\s*,|\s*localiz)', text_block, re.IGNORECASE)
    victima = victim_match.group(1).strip().upper() if victim_match else "DESCONOCIDO"

    # Age
    age_match = re.search(r'(\d{1,2})\s*años', text_block, re.IGNORECASE)
    edad = int(age_match.group(1)) if age_match else None

    # Shells / Caliber
    shells_match = re.search(r'(\d+)\s*casquillos\s*percutidos\s*(?:calibre|de)?\s*([0-9a-zA-Z\s]{2,10})', text_block, re.IGNORECASE)
    casquillos = f"{shells_match.group(1)} casquillos {shells_match.group(2).strip()}" if shells_match else "No especificado"

    # Police Officers / Units
    police_match = re.search(r'Policía\s+([A-ZÁÉÍÓÚÑ\s]{5,35})\s+unidad\s+([A-Z0-9\-]+)', text_block, re.IGNORECASE)
    policia_unidad = f"{police_match.group(1).strip()} ({police_match.group(2).strip()})" if police_match else "No especificado"

    # Location (Street / Colonia)
    loc_match = re.search(r'en\s+(calle\s+[^,]+,\s*colonia\s+[^,]+)', text_block, re.IGNORECASE)
    ubicacion = loc_match.group(1).strip() if loc_match else "No especificada"

    # Time
    time_match = re.search(r'A las\s+(\d{1,2}:\d{2})\s*horas', text_block, re.IGNORECASE)
    hora = time_match.group(1) if time_match else "No especificada"

    return {
        "municipio": municipio,
        "victima": victima,
        "edad": edad,
        "casquillos": casquillos,
        "policia_unidad": policia_unidad,
        "ubicacion": ubicacion,
        "hora": hora,
        "narrativa_completa": text_block[:500],
        "fecha_registro": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

ingested_cards = []
seen_victims = set()

for idx, block in enumerate(raw_messages):
    card_data = parse_card_text(block)
    
    # Deduplication check
    v_key = card_data['victima']
    if v_key != "DESCONOCIDO" and v_key in seen_victims:
        continue
    if v_key != "DESCONOCIDO":
        seen_victims.add(v_key)

    # Cross-reference with 911 Calls
    matched_call = None
    if not df_911.empty and card_data['municipio'] != "DESCONOCIDO":
        filtered = df_911[df_911['MUNICIPIO'].astype(str).str.upper().str.contains(card_data['municipio'], na=False)]
        for _, row in filtered.iterrows():
            notes = str(row.get('NOTAS', '')).upper()
            addr = str(row.get('DIRECCION ', '')).upper()
            full_text = f"{notes} {addr}"

            if (card_data['victima'] != "DESCONOCIDO" and card_data['victima'] in full_text) or \
               (card_data['ubicacion'] != "No especificada" and any(w in full_text for w in card_data['ubicacion'].upper().split() if len(w) > 4)):
                matched_call = str(row.get('FOLIO', ''))
                break

    if matched_call:
        card_data['folio_911_vinculado'] = matched_call
        card_data['coincidencia_estatus'] = "VINCULADA 911 C5"
    else:
        card_data['folio_911_vinculado'] = "N/A"
        card_data['coincidencia_estatus'] = "SIN REGISTRO PREVIO 911"

    card_data['id_tarjeta'] = f"TI-WA-{idx+1:03d}"
    ingested_cards.append(card_data)

print(f"\nTotal Tarjetas Informativas únicas consolidadas: {len(ingested_cards)}")

# Save to Master Excel
df_ti_master = pd.DataFrame(ingested_cards)
with pd.ExcelWriter(master_ti_excel, engine='openpyxl') as writer:
    df_ti_master.to_excel(writer, sheet_name='TARJETAS_INFORMATIVAS', index=False)

print(f"[OK] Concentrado Master de Tarjetas Informativas actualizado en: {master_ti_excel}")
