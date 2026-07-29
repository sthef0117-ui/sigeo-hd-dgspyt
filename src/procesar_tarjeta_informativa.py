import os
import glob
import json
import re
import pandas as pd
from datetime import datetime

base_dir = r"C:\Users\xXZaB\.gemini\antigravity\scratch\sigeo_hd_dgspyt"
historico_dir = os.path.join(base_dir, "historico_c5")
ti_dir = os.path.join(historico_dir, "tarjetas_informativas")

os.makedirs(os.path.join(ti_dir, "texto"), exist_ok=True)
os.makedirs(os.path.join(ti_dir, "imagenes"), exist_ok=True)

master_ti_excel = os.path.join(ti_dir, "CONCENTRADO_MASTER_TARJETAS_INFORMATIVAS.xlsx")

def extract_metadata_from_text(raw_text):
    """ Extrae campos estructurados de la narrativa de una Tarjeta Informativa """
    date_match = re.search(r'(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})', raw_text)
    fecha = date_match.group(1) if date_match else datetime.now().strftime("%d/%m/%Y")
    
    time_match = re.search(r'(\d{1,2}:\d{2}\s*(?:hrs|HRS|am|pm|AM|PM)?)', raw_text)
    hora = time_match.group(1) if time_match else "No especificada"

    muns = ["ECATEPEC", "NAUCALPAN", "CUAUTITLÁN IZCALLI", "NEZAHUALCÓYOTL", "CHIMALHUACÁN", "TULTITLÁN", "TOLUCA", "CHALCO", "IXTAPALUCA", "TEPOZOTLÁN"]
    found_mun = "POR CLASIFICAR"
    for m in muns:
        if m.lower() in raw_text.lower():
            found_mun = m
            break

    delito = "HOMICIDIO DOLOSO" if any(w in raw_text.upper() for w in ["HOMICIDIO", "OCCISO", "CADAVER", "EJECUTADO", "DISPARO"]) else "INCIDENTE RELEVANTE"

    return {
        "fecha": fecha,
        "hora": hora,
        "municipio": found_mun,
        "delito": delito,
        "narrativa_completa": raw_text.strip(),
        "fecha_registro": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

def ingestar_tarjeta_texto(raw_text):
    data = extract_metadata_from_text(raw_text)
    timestamp_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    data["id_tarjeta"] = f"TI-{timestamp_id}"

    json_path = os.path.join(ti_dir, "texto", f"{data['id_tarjeta']}.json")
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    df_new = pd.DataFrame([data])
    if os.path.exists(master_ti_excel):
        df_master = pd.read_excel(master_ti_excel)
        df_combined = pd.concat([df_master, df_new], ignore_index=True)
    else:
        df_combined = df_new

    with pd.ExcelWriter(master_ti_excel, engine='openpyxl') as writer:
        df_combined.to_excel(writer, sheet_name='TARJETAS_INFORMATIVAS', index=False)

    print(f"[OK] Tarjeta Informativa #{data['id_tarjeta']} ingestada y guardada en el concentrado Master.")
    return data

if __name__ == "__main__":
    test_text = "TARJETA INFORMATIVA - 28/07/2026. En Ecatepec a las 11:30 hrs se reportó detonaciones de arma de fuego en colonia Centro."
    ingestar_tarjeta_texto(test_text)
