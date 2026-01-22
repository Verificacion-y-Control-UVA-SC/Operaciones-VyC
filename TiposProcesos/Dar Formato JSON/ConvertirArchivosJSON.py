import pandas as pd
import os

# Carpeta donde están los Excel
BASE_PATH = os.path.join(os.getcwd(), "archivos")

# Archivos Excel y su JSON de salida
archivos_excel = {
    "BASE DECATHLON GENERAL ADVANCE II.xlsx": "base_general.json",
    "codigos_cumple.xlsx": "codigos_cumple.json",
    "HISTORIAL_PROCESOS.xlsx": "historial.json"
}

# Carpeta donde se guardarán los JSON
RESOURCES_PATH = os.path.join(os.getcwd(), "resources")
if not os.path.exists(RESOURCES_PATH):
    os.makedirs(RESOURCES_PATH)

# Procesar cada archivo
for excel_file, json_file in archivos_excel.items():
    excel_path = os.path.join(BASE_PATH, excel_file)  # Ruta completa al Excel
    json_path = os.path.join(RESOURCES_PATH, json_file)  # Ruta completa al JSON
    
    if os.path.exists(excel_path):
        df = pd.read_excel(excel_path)
        df.to_json(json_path, orient="records", force_ascii=False, indent=4)
        print(f"{excel_file} → resources/{json_file}")
    else:
        print(f"No se encontró el archivo: {excel_path}")
