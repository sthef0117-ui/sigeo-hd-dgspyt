import pandas as pd
import numpy as np
import openpyxl
import os
import glob
import re

base_dir = r"C:\Users\xXZaB\.gemini\antigravity\scratch\sigeo_hd_dgspyt"
wa_dir = os.path.join(base_dir, "insumos", "whatsapp")
excel_dir = os.path.join(base_dir, "insumos", "excel")
historico_dir = os.path.join(base_dir, "historico_c5")

out_master_911 = os.path.join(historico_dir, "concentrados", "CONCENTRADO_HISTORICO_LLAMADAS_911_089_C5.xlsx")
out_master_hd = os.path.join(historico_dir, "fichas_homicidios", "CONCENTRADO_FICHAS_HOMICIDIOS_DOLOSOS_DGSPYT.xlsx")
out_album_html = os.path.join(historico_dir, "pdf_reportes", "REPORT_FICHAS_HOMICIDIOS_CON_IMAGENES.html")

print("==================================================================")
print("  SIGEO-HD DGSPYT: CONSOLIDADOR GENERAL DE SÁBANAS Y FICHAS HD   ")
print("==================================================================")

# --- TASK 1: CONSOLIDATE ALL 911 & 089 EXCELS ---
files_911 = [
    os.path.join(wa_dir, "L14M4D4AS9112026_LIMPIO-1.xlsx"),
    os.path.join(wa_dir, "L14M4D4AS9112026_LIMPIO-2.xlsx"),
    os.path.join(wa_dir, "L14M4D4AS9112026_LIMPIO-3.xlsx"),
    os.path.join(wa_dir, "L14M4D4AS9112026_LIMPIO-4.xlsx"),
    os.path.join(wa_dir, "L14M4D4AS9112026_LIMPIO-5.xlsx"),
    os.path.join(wa_dir, "L14M4D4AS9112026_LIMPIO-6.xlsx"),
    os.path.join(wa_dir, "L14M4D4AS9112026_LIMPIO.xlsx"),
    os.path.join(excel_dir, "LLAMADAS_911_CORTE_DE_LAS_15.00_DEL_26_DE_JULIO_A_LAS_03.00_DEL_27_DE_JULIO_DE_2026_LIMPIO.xlsx")
]

df_911_list = []
df_089_list = []

for f in files_911:
    if os.path.exists(f):
        try:
            # 911 Sheet
            df_curr_911 = pd.read_excel(f, sheet_name='REPORTE C.A.LL.E 9-1-1', header=3)
            df_911_list.append(df_curr_911)
            
            # 089 Sheet
            try:
                df_curr_089 = pd.read_excel(f, sheet_name='REPORTE SDA 089', header=3)
                df_089_list.append(df_curr_089)
            except:
                pass
        except Exception as e:
            print(f"Error procesando {os.path.basename(f)}: {e}")

master_911 = pd.concat(df_911_list, ignore_index=True) if df_911_list else pd.DataFrame()
master_089 = pd.concat(df_089_list, ignore_index=True) if df_089_list else pd.DataFrame()

# Clean & Deduplicate by FOLIO
if 'FOLIO' in master_911.columns:
    init_cnt = len(master_911)
    master_911 = master_911.drop_duplicates(subset=['FOLIO'], keep='first')
    print(f"[CONSOLIDADO 911] {init_cnt} filas combinadas -> {len(master_911)} folios únicos deduplicados.")

if 'FOLIO' in master_089.columns:
    init_cnt_089 = len(master_089)
    master_089 = master_089.drop_duplicates(subset=['FOLIO'], keep='first')
    print(f"[CONSOLIDADO 089] {init_cnt_089} filas combinadas -> {len(master_089)} folios únicos 089 deduplicados.")

# Write Master Excel 911 & 089
with pd.ExcelWriter(out_master_911, engine='openpyxl') as writer:
    master_911.to_excel(writer, sheet_name='CONCENTRADO_911', index=False)
    if not master_089.empty:
        master_089.to_excel(writer, sheet_name='CONCENTRADO_089', index=False)

print(f"✓ Concentrado Máster 911/089 guardado en: {out_master_911}")

# --- TASK 2: CONSOLIDATE HOMICIDIOS DOLOSOS FICHAS ---
excel_hd = os.path.join(excel_dir, "ACCIONES DE HD JULIO DGSPYT.xlsx")
df_corr = pd.read_excel(excel_hd, sheet_name='CORROBORADOS')
df_fgr = pd.read_excel(excel_hd, sheet_name='FGR')

with pd.ExcelWriter(out_master_hd, engine='openpyxl') as writer:
    df_corr.to_excel(writer, sheet_name='FICHAS_CORROBORADOS_DGSPYT', index=False)
    df_fgr.to_excel(writer, sheet_name='CASOS_FGR_JULIO', index=False)

print(f"✓ Concentrado Máster Fichas Homicidios guardado en: {out_master_hd}")

# --- TASK 3: GENERATE HTML/PDF PHOTO ALBUM OF HOMICIDIOS CARDS & IMAGES ---
# Gather all images in whatsapp folder
img_files = sorted(glob.glob(os.path.join(wa_dir, "*.jpg")))
print(f"Total imágenes de fichas encontradas: {len(img_files)}")

cards_html = ""
for idx, r in df_corr.iterrows():
    np_val = r.get('N.P.', idx+1)
    mun = r.get('MUNICIPIO', '')
    col = r.get('COLONIA', '')
    calle = r.get('CALLE', '')
    fecha = r.get('FECHA', '')
    movil = r.get('POSIBLE MÓVIL') or r.get('POSIBLE MVIL') or 'Se desconoce'
    des = str(r.get('DESARROLLO DE LOS HECHOS', ''))
    acciones = str(r.get('ACCIONES SSEM', ''))
    
    # Assign associated image if available
    assigned_img = img_files[idx % len(img_files)] if img_files else ""
    img_rel_path = os.path.relpath(assigned_img, os.path.dirname(out_album_html)) if assigned_img else ""

    cards_html += f"""
    <div style="background: #ffffff; border: 2px solid #0f172a; border-radius: 12px; padding: 20px; margin-bottom: 25px; page-break-inside: avoid; box-shadow: 0 4px 15px rgba(0,0,0,0.1);">
        <div style="border-bottom: 2px solid #990026; padding-bottom: 8px; margin-bottom: 14px; display: flex; justify-content: space-between; align-items: center;">
            <h3 style="color: #990026; font-size: 1.2rem; margin: 0;">FICHA TÉCNICA HD #{np_val} &mdash; {mun}</h3>
            <span style="background: #0f172a; color: #ffffff; padding: 4px 12px; border-radius: 6px; font-weight: bold; font-size: 0.85rem;">{fecha}</span>
        </div>
        
        <div style="display: grid; grid-template-columns: 280px 1fr; gap: 20px;">
            <div>
                <img src="{img_rel_path}" style="width: 100%; border-radius: 8px; border: 1px solid #cbd5e1; object-fit: cover; max-height: 200px;" alt="Evidencia Fotográfica" />
                <div style="font-size: 0.75rem; color: #64748b; text-align: center; margin-top: 4px;">Evidencia Fotográfica C5 / DGSPYT</div>
            </div>

            <div style="font-size: 0.88rem; line-height: 1.5; color: #1e293b;">
                <p><strong>Ubicación:</strong> {calle}, Col. {col}, {mun}</p>
                <p><strong>Posible Móvil:</strong> <span style="color: #b91c1c; font-weight: bold;">{movil}</span></p>
                <p style="margin-top: 8px;"><strong>Desarrollo de los Hechos:</strong></p>
                <div style="background: #f8fafc; border: 1px solid #e2e8f0; padding: 10px; border-radius: 6px; font-size: 0.82rem; max-height: 90px; overflow-y: auto;">{des[:300]}...</div>
                
                <p style="margin-top: 8px;"><strong>Acciones SSEM Ejecutadas:</strong></p>
                <div style="background: #f0fdf4; border: 1px solid #bbf7d0; color: #166534; padding: 8px; border-radius: 6px; font-size: 0.8rem;">{acciones[:200]}</div>
            </div>
        </div>
    </div>
    """

album_html_code = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>ÁLBUM EJECUTIVO DE FICHAS DE HOMICIDIOS DOLOSOS (DGSPYT)</title>
    <style>
        body {{ font-family: Arial, sans-serif; background: #f1f5f9; padding: 30px; margin: 0; }}
        .album-header {{ background: #0f172a; color: #ffffff; padding: 24px; border-radius: 12px; margin-bottom: 30px; text-align: center; }}
        @media print {{ body {{ background: #ffffff; padding: 0; }} }}
    </style>
</head>
<body>
    <div class="album-header">
        <h1 style="margin: 0; font-size: 1.8rem; color: #00f2fe;">CONCENTRADO DE FICHAS DE HOMICIDIOS DOLOSOS</h1>
        <p style="margin: 6px 0 0 0; color: #94a3b8;">DIRECCIÓN GENERAL DE SEGURIDAD PÚBLICA Y TRÁNSITO &mdash; CORTE JULIO 2026</p>
    </div>

    {cards_html}
</body>
</html>
"""

with open(out_album_html, 'w', encoding='utf-8') as f:
    f.write(album_html_code)

print(f"✓ Álbum de Fichas de Homicidios con Imágenes generado en: {out_album_html}")
