import os, json, pandas as pd, unicodedata
from pathlib import Path

DATA_DIR = Path("data")
DATA_DIR.mkdir(parents=True, exist_ok=True)
RESOURCES_DIR = Path("data")
RESOURCES_DIR.mkdir(parents=True, exist_ok=True)

def normalizar_texto(texto):
    """Elimina acentos y espacios innecesarios de los encabezados"""
    if not isinstance(texto, str):
        return texto
    texto = texto.strip()
    texto = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('utf-8')
    texto = texto.replace("  ", " ")
    return texto

def convertir_a_json(archivo_excel, sheet_name=0, nombre_json="data.json", persist=True):
    """
    Convierte una hoja de Excel a JSON.
    Si el archivo es un layout, lee desde la fila 3 (header=2), limpia columnas vacías,
    normaliza encabezados y guarda un layout_preview.json.
    """
    try:
        # Detectar si es un layout
        es_layout = "layout" in nombre_json.lower()

        if es_layout:
            df = pd.read_excel(archivo_excel, sheet_name=sheet_name, header=2, dtype=str)
            df.columns = [normalizar_texto(c) for c in df.columns]
            df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
            df.dropna(how="all", inplace=True)
            df.fillna("", inplace=True)

            # Renombrar columnas normalizadas a sus nombres originales esperados
            columnas_correctas = {
                "folio de solicitud": "Folio de Solicitud",
                "nom": "NOM",
                "numero de acreditacion": "Número de Acreditación",
                "rfc": "RFC",
                "denominacion social o nombre": "Denominación social o nombre",
                "tipo de persona": "Tipo de persona",
                "marca del producto": "Marca del producto",
                "descripcion del producto": "Descripción del producto",
                "fraccion arancelaria": "Fracción arancelaria",
                "fecha de envio de la solicitud": "Fecha de envío de la solicitud",
                "vigencia de la solicitud": "Vigencia de la Solicitud",
                "modalidad de etiquetado": "Modalidad de etiquetado",
                "modelo": "Modelo",
                "umc": "UMC",
                "cantidad": "Cantidad",
                "numero de etiquetas a verificar": "Número de etiquetas a verificar",
                "parte": "Parte",
                "partida": "Partida",
                "pais origen": "Pais Origen",
                "pais comprador": "Pais Comprador",
            }

            df.rename(columns=columnas_correctas, inplace=True)

        else:
            df = pd.read_excel(archivo_excel, sheet_name=sheet_name, header=0, dtype=str)
            df.columns = [c.strip() for c in df.columns]
            df.fillna("", inplace=True)

        records = df.to_dict(orient="records")

        # Guardar solo si persist=True
        destino = DATA_DIR / nombre_json
        if persist:
            with open(destino, "w", encoding="utf-8") as f:
                json.dump(records, f, indent=4, ensure_ascii=False)

        # Ya no se guarda layout_preview.json, solo el archivo principal

        return records

    except Exception as e:
        print(f"❌ Error convertir_a_json: {e}")
        return None

def leer_json(ruta):
    try:
        p = Path(ruta)
        if p.exists():
            with open(p, "r", encoding="utf-8") as f:
                return pd.read_json(f, orient="records")
        return pd.DataFrame()
    except Exception as e:
        print(f"Error leer_json: {e}")
        return pd.DataFrame()

# --- Falta la persistencia de datos ya que no se guardan cuando se cierra el programa --- #
def guardar_config(config_dict):
    try:
        config_path = RESOURCES_DIR / "config.json"
        existing = {}
        if config_path.exists():
            with open(config_path, "r", encoding="utf-8") as f:
                existing = json.load(f)
        existing.update(config_dict)
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(existing, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Error guardar_config: {e}")
