import pandas as pd
import os

# Definir rutas
carpeta_archivos = "archivos"
carpeta_salida = "resources"

# Crear carpeta resources si no existe
os.makedirs(carpeta_salida, exist_ok=True)

# -----------------------------
# ARCHIVO BASE GENERAL
# -----------------------------
archivo_base = os.path.join(carpeta_archivos, "BASE_GENERAL_ULTA.xlsx")
df_base = pd.read_excel(archivo_base, sheet_name=0)  # primera hoja
archivo_base_json = os.path.join(carpeta_salida, "base_general.json")
df_base.to_json(archivo_base_json, orient="records", force_ascii=False, indent=4)

print("✅ Archivos convertidos a JSON y guardados en carpeta 'resources'")
