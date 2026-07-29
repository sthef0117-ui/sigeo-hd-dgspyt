import pandas as pd
import openpyxl
import os
import shutil
import glob
from datetime import datetime

base_dir = r"C:\Users\xXZaB\.gemini\antigravity\scratch\sigeo_hd_dgspyt"
downloads_dir = r"C:\Users\xXZaB\Downloads"
historico_dir = os.path.join(base_dir, "historico_c5")

master_path = os.path.join(historico_dir, "concentrados", "CONCENTRADO_HISTORICO_LLAMADAS_911_089_C5.xlsx")
diario_dir = os.path.join(historico_dir, "diario")

os.makedirs(diario_dir, exist_ok=True)

print("=========================================================================")
print("  SIGEO-HD DGSPYT: PROCESADOR AUTOMÁTICO DE NUEVAS SÁBANAS Y LIMPIEZA    ")
print("=========================================================================")

# Look for downloaded 911 / 089 Excel files in Downloads
pattern = os.path.join(downloads_dir, "*.xlsx")
downloaded_excels = [f for f in glob.glob(pattern) if "LLAMADAS" in f.upper() or "911" in f.upper() or "089" in f.upper()]

if not downloaded_excels:
    print("[INFO] No se encontraron nuevas sábanas descargadas en Downloads.")
else:
    print(f"Archivos de sábanas detectados ({len(downloaded_excels)}):")
    for f in downloaded_excels:
        fname = os.path.basename(f)
        print(f"  - Procesando: {fname}")
        
        # 1. Archive daily copy
        today_str = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        archive_name = f"{today_str}_{fname}"
        archive_path = os.path.join(diario_dir, archive_name)
        shutil.copy2(f, archive_path)
        print(f"    ✓ Archivado en histórico diario: {archive_name}")

        # 2. Append non-duplicate rows to Master Excel
        try:
            df_new_911 = pd.read_excel(f, sheet_name='REPORTE C.A.LL.E 9-1-1', header=3)
            df_master_911 = pd.read_excel(master_path, sheet_name='CONCENTRADO_911')
            
            # Combine and deduplicate by FOLIO
            combined_911 = pd.concat([df_master_911, df_new_911], ignore_index=True)
            combined_911 = combined_911.drop_duplicates(subset=['FOLIO'], keep='first')

            # Process 089 if exists
            try:
                df_new_089 = pd.read_excel(f, sheet_name='REPORTE SDA 089', header=3)
                df_master_089 = pd.read_excel(master_path, sheet_name='CONCENTRADO_089')
                combined_089 = pd.concat([df_master_089, df_new_089], ignore_index=True)
                combined_089 = combined_089.drop_duplicates(subset=['FOLIO'], keep='first')
            except:
                combined_089 = pd.read_excel(master_path, sheet_name='CONCENTRADO_089')

            # Write updated Master Excel
            with pd.ExcelWriter(master_path, engine='openpyxl') as writer:
                combined_911.to_excel(writer, sheet_name='CONCENTRADO_911', index=False)
                combined_089.to_excel(writer, sheet_name='CONCENTRADO_089', index=False)

            print(f"    ✓ Concentrado Máster actualizado ({len(combined_911)} folios 911).")

            # 3. Delete raw download file after successful ingestion
            os.remove(f)
            print(f"    ✓ Archivo único eliminado de Downloads: {fname}")

        except Exception as e:
            print(f"    [ERROR] No se pudo procesar {fname}: {e}")

print("=========================================================================")
