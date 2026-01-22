# Rutas.py
import os
import sys
import json

def ruta_base():
    """Devuelve la carpeta base, ya sea del .exe o del script en desarrollo"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.abspath(".")

def archivo_datos(nombre_archivo):
    """
    Devuelve la ruta absoluta de un archivo dentro de /datos
    y crea la carpeta si no existe.
    """
    ruta_datos = os.path.join(ruta_base(), "datos")
    os.makedirs(ruta_datos, exist_ok=True)
    return os.path.join(ruta_datos, nombre_archivo)

# --- Inicialización automática de archivos ---
ARCHIVOS_INICIALES = {
    "archivos_procesados.json": [],
    "base_general.json": [],
    "config.json": {},
    "codigos_cumple.json": [],
    "codigos_cumple.xlsx": None  # Aquí puedes luego cargar o reemplazar el Excel real
}

def inicializar_archivos():
    """
    Crea los archivos de datos si no existen.
    Para los JSON escribe estructura vacía.
    Para el Excel, solo se asegura que la ruta exista (archivo puede ser cargado después).
    """
    for nombre, contenido in ARCHIVOS_INICIALES.items():
        ruta = archivo_datos(nombre)
        if not os.path.exists(ruta):
            if nombre.endswith(".json"):
                with open(ruta, "w", encoding="utf-8") as f:
                    json.dump(contenido, f, indent=4, ensure_ascii=False)
            else:
                # Para Excel u otros archivos no JSON, se deja vacío (el usuario lo cargará)
                open(ruta, "a").close()

# Llamar automáticamente al importar
inicializar_archivos()
