# Configuracion.py
import os
import json
import pandas as pd
import re, shutil
from tkinter import filedialog, messagebox, simpledialog
from PIL import Image
import customtkinter as ctk


# ---------------- CONFIGURACIÓN ---------------- #
STYLE = {
    "primario": "#ECD925",        # Amarillo dorado más vibrante
    "secundario": "#282828",      # Azul oscuro elegante en lugar de negro puro
    "exito": "#27AE60",           # Verde más vivo
    "advertencia": "#d57067",     # Naranja cálido
    "peligro": "#d74a3d",         # Rojo más intenso
    "fondo": "#F8F9FA",           # Fondo gris muy claro (mantenido)
    "surface": "#FFFFFF",         # Superficies blancas (mantenido)
    "texto_oscuro": "#282828",    # Texto principal - azul oscuro
    "texto_claro": "#4b4b4b",     # Texto secundario - gris azulado
    "borde": "#BDC3C7",           # Bordes gris claro
    "header_texto": "#282828",    # Texto del header
    "hover_primario": "#ECD925",  # Hover para el amarillo
    "hover_boton": "#4b4b4b"      # Hover para los botones
}

FONT_TITLE = ("Inter", 20, "bold")
FONT_SUB = ("Inter", 16, "bold")
FONT_LABEL = ("Inter", 14)
FONT_TEXT = ("Inter", 12)


# ---------------- CONFIGURACIÓN DE RUTAS ---------------- #
DATA_DIR = "data"
JSON_FILE = os.path.join(DATA_DIR, "base_etiquetado.json")
CONFIG_FILE = os.path.join(DATA_DIR, "config.json")
os.makedirs(DATA_DIR, exist_ok=True)

# ---------------- GESTIÓN DE CONFIGURACIÓN ---------------- #
def cargar_configuracion():
    """Carga la configuración desde el archivo JSON"""
    config_default = {"image_dir": "", "ultima_carpeta": ""}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
            for key, value in config_default.items():
                if key not in config:
                    config[key] = value
            return config
        except Exception as e:
            print(f"Error cargando configuración: {e}")
    return config_default

def guardar_configuracion(config):
    """Guarda la configuración en el archivo JSON"""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
        return True
    except Exception as e:
        print(f"Error guardando configuración: {e}")
        return False

def seleccionar_carpeta_imagenes():
    """Permite al usuario seleccionar la carpeta de imágenes"""
    config = cargar_configuracion()
    carpeta_inicial = config.get("ultima_carpeta", "")
    carpeta = filedialog.askdirectory(title="Seleccionar carpeta de imágenes", initialdir=carpeta_inicial)
    if carpeta:
        config["image_dir"] = carpeta
        config["ultima_carpeta"] = carpeta
        guardar_configuracion(config)
        return carpeta
    return None

# ---------------- FUNCIONES BASE ---------------- #
def normalizar_columna(col):
    return str(col).strip().upper().replace(" ", "_")

def cargar_base_excel(path_excel):
    try:
        df = pd.read_excel(path_excel)
        df.columns = [normalizar_columna(c) for c in df.columns]

        columnas_requeridas = [
            "CATEGORIA","UPC","DENOMINACION","DENOMINACION_AXO","MARCA",
            "OBSERVACIONES_REVISION","LEYENDAS_PRECAUTORIAS","INSTRUCCIONES_DE_USO","OBSERVACIONES",
            "TAMAÑO_DE_LA_DECLARACION_DE_CONTENIDO","CONTENIDO","PAIS_ORIGEN",
            "IMPORTADOR","NORMA","INGREDIENTES_Y_LOTE","MEDIDAS","TIPO_DE_ETIQUETA"
        ]

        faltantes = [col for col in columnas_requeridas if col not in df.columns]
        if faltantes:
            messagebox.showerror("Error en columnas", f"Faltan las siguientes columnas:\n{', '.join(faltantes)}")
            return None

        df = df[columnas_requeridas].fillna("")
        registros = df.to_dict(orient="records")

        with open(JSON_FILE, "w", encoding="utf-8") as f:
            json.dump(registros, f, ensure_ascii=False, indent=4)

        excel_info = {
            "excel_path": path_excel,
            "fecha_carga": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        with open(os.path.join(DATA_DIR, "excel_info.json"), "w", encoding="utf-8") as f:
            json.dump(excel_info, f, ensure_ascii=False, indent=4)

        messagebox.showinfo("Carga completada", f"Archivo JSON guardado en: {JSON_FILE}")
        return df, path_excel
    except Exception as e:
        messagebox.showerror("Error", f"Error al cargar el Excel:\n{e}")
        return None

def cargar_json():
    if not os.path.exists(JSON_FILE):
        return []
    try:
        with open(JSON_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, list):
            messagebox.showerror("Error", "El archivo JSON no tiene formato correcto.")
            return []
        productos_validos = [d for d in data if isinstance(d, dict) and "UPC" in d]
        return productos_validos
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo cargar el JSON:\n{e}")
        return []

def buscar_producto_por_upc(upc_or_categoria, productos):
    """Busca un producto por su UPC o por CATEGORIA.

    Comportamiento:
    - Si se encuentra una coincidencia exacta por UPC, devuelve ese producto.
    - Si no hay coincidencia por UPC, busca por CATEGORIA (coincidencia exacta
      ignorando mayúsculas). Si hay varias coincidencias se devuelve la primera.
    - Devuelve None si no encuentra nada.
    """
    if not productos or not upc_or_categoria:
        return None

    termino = str(upc_or_categoria).strip()
    # Primero intentar buscar por UPC (igual que antes)
    for p in productos:
        upc_val = p.get("UPC") or p.get("upc") or next((p[k] for k in p if "UPC" in k.upper()), None)
        if upc_val and str(upc_val).strip() == termino:
            return p

    # Si no se encontró por UPC, intentar buscar por categoría
    termino_upper = termino.upper()
    for p in productos:
        cat = p.get("CATEGORIA") or p.get("categoria") or p.get("Categoria")
        if cat and str(cat).strip().upper() == termino_upper:
            return p

    # Como último recurso, buscar por inclusión parcial en la categoría
    for p in productos:
        cat = p.get("CATEGORIA") or p.get("categoria") or p.get("Categoria")
        if cat and termino_upper in str(cat).strip().upper():
            return p

    return None

def mostrar_imagen(label, categoria):
    config = cargar_configuracion()
    image_dir = config.get("image_dir")
    
    if not image_dir or not os.path.isdir(image_dir):
        label.configure(image="", text="❌ Carpeta de imágenes no configurada")
        return
    
    # Buscar imagen que coincida con la categoría
    imagen_path = None
    categoria_norm = str(categoria).strip().lower()
    # Primera pasada: coincidencia exacta por nombre de archivo (sin extensión)
    for archivo in os.listdir(image_dir):
        nombre, ext = os.path.splitext(archivo)
        if ext.lower() not in ['.jpg', '.jpeg', '.png', '.bmp', '.gif']:
            continue
        if nombre.strip().lower() == categoria_norm:
            imagen_path = os.path.join(image_dir, archivo)
            break
    # Nota: NO hacer fallback por inclusión — solo coincidencia exacta para evitar mostrar imágenes equivocadas
    
    if imagen_path and os.path.exists(imagen_path):
        try:
            # Usar CTkImage en lugar de PhotoImage
            imagen_pil = Image.open(imagen_path)
            # Redimensionar manteniendo aspecto
            width, height = imagen_pil.size
            if width > 400 or height > 400:
                ratio = min(400/width, 400/height)
                new_size = (int(width * ratio), int(height * ratio))
                imagen_pil = imagen_pil.resize(new_size, Image.Resampling.LANCZOS)
            
            imagen_ctk = ctk.CTkImage(imagen_pil, size=imagen_pil.size)
            label.configure(image=imagen_ctk, text="")
            label.image = imagen_ctk  # Mantener referencia
        except Exception as e:
            print(f"Error al cargar imagen: {e}")
            label.configure(image="", text=f"❌ Error al cargar imagen: {e}")
    else:
        # No mostrar imagen cuando no se encuentra archivo acorde a la categoría
        label.configure(image="", text="")
        try:
            label.image = None
        except Exception:
            pass

def extraer_valor_numerico(texto):
    """Extrae el número de un texto como '50 ml' -> 50"""
    if not texto:
        return 0
    try:
        s = str(texto).strip()
        # Reemplazar coma decimal por punto para aceptar formatos como '1,5 mm'
        s = s.replace(',', '.')
        numeros = re.findall(r"(\d+\.?\d*)", s)
        return float(numeros[0]) if numeros else 0
    except:
        return 0

def _quitar_acentos(s: str) -> str:
    """Quita tildes y caracteres especiales básicos para normalizar claves."""
    if not isinstance(s, str):
        return s
    s = s.upper()
    replacements = {
        'Á': 'A', 'É': 'E', 'Í': 'I', 'Ó': 'O', 'Ú': 'U', 'Ñ': 'N', 'À': 'A', 'Ä': 'A'
    }
    for k, v in replacements.items():
        s = s.replace(k, v)
    return s

def obtener_campo(d: dict, campo_base: str):
    """Devuelve el valor de un campo buscando variantes de nombre.

    Prueba variantes: exacta, con espacios/guiones/underscores, sin tildes,
    y busca por inclusión parcial si no encuentra una coincidencia exacta.
    """
    if not isinstance(d, dict):
        return None

    if campo_base in d:
        return d.get(campo_base)

    # Variantes a probar
    variantes = set()
    base = str(campo_base)
    variantes.add(base)
    variantes.add(base.replace(' ', '_'))
    variantes.add(base.replace('_', ' '))
    variantes.add(base.replace(' ', '').replace('_', ''))

    # Normalizar acentos y mayúsculas
    base_sin = _quitar_acentos(base)
    variantes.add(base_sin)
    variantes.add(base_sin.replace(' ', '_'))
    variantes.add(base_sin.replace('_', ' '))
    variantes.add(base_sin.replace(' ', '').replace('_', ''))

    # Probar variantes directas (mayúsculas/minúsculas)
    for v in list(variantes):
        # probar como está
        if v in d:
            return d.get(v)
        # probar version mayúsculas
        if v.upper() in d:
            return d.get(v.upper())
        if v.lower() in d:
            return d.get(v.lower())

    # Si no hay coincidencias exactas, buscar por clave que contenga el texto base
    buscar = base_sin
    for k in d.keys():
        k_norm = _quitar_acentos(str(k))
        if buscar in k_norm or _quitar_acentos(str(buscar)) in k_norm:
            return d.get(k)

    return None

def subir_imagen(self):
    """Permite seleccionar y guardar una imagen sin requerir un UPC."""
    try:
        # --- Verificar carpeta de imágenes ---
        config = cargar_configuracion()
        image_dir = config.get("image_dir", "")
        if not image_dir:
            messagebox.showwarning(
                "Carpeta no configurada",
                "Primero selecciona una carpeta de imágenes usando el botón 'Seleccionar Carpeta'."
            )
            return

        # --- Seleccionar imagen ---
        ruta_img = filedialog.askopenfilename(
            title="Seleccionar imagen",
            filetypes=[("Imágenes", "*.jpg *.png *.jpeg")]
        )
        if not ruta_img:
            return  # Usuario canceló

        # --- Solicitar nombre o categoría para guardar ---
        categoria = simpledialog.askstring(
            "Guardar imagen como",
            "Ingresa el nombre o categoría de la imagen (sin extensión):"
        )

        if not categoria:
            messagebox.showwarning("Atención", "Debes ingresar un nombre o categoría.")
            return

        # --- Guardar en carpeta configurada: conservar la extensión original ---
        ext = os.path.splitext(ruta_img)[1].lower()
        if ext not in ['.jpg', '.jpeg', '.png', '.bmp', '.gif']:
            # Forzar a .jpg si la extensión no es estándar
            ext = '.jpg'

        destino = os.path.join(image_dir, f"{categoria}{ext}")

        # Si ya existe un archivo con ese nombre, preguntar al usuario.
        # Permitir sobrescribir o ingresar un nuevo nombre; si cancela, abortar.
        while os.path.exists(destino):
            respuesta = messagebox.askyesno(
                "Archivo existente",
                f"Ya existe un archivo llamado {os.path.basename(destino)} en la carpeta de imágenes.\n¿Deseas sobrescribirlo?"
            )
            if respuesta:
                # Usuario confirmó sobrescribir
                break
            # Permitir al usuario ingresar un nuevo nombre
            nuevo = simpledialog.askstring(
                "Nuevo nombre",
                "Ingresa un nuevo nombre para la imagen (sin extensión) o deja vacío para cancelar:"
            )
            if not nuevo:
                messagebox.showinfo("Cancelado", "La operación de subida fue cancelada.")
                return
            # Recalcular destino con el nuevo nombre proporcionado
            categoria = nuevo.strip()
            destino = os.path.join(image_dir, f"{categoria}{ext}")

        # Copiar archivo al destino final
        shutil.copyfile(ruta_img, destino)

        # --- Mostrar imagen cargada (buscar por nombre de categoria) ---
        # Llamamos a mostrar_imagen para que haga el thumbnail y la carga correcta.
        mostrar_imagen(self.label_imagen, categoria)

        # Actualizar estados visuales (si existen en el objeto pasado)
        try:
            if hasattr(self, 'label_estado_imagen'):
                self.label_estado_imagen.configure(
                    text=f"✅ Imagen guardada como {os.path.basename(destino)}",
                    text_color=STYLE["exito"]
                )
        except Exception:
            pass

        try:
            if hasattr(self, 'label_estado'):
                self.label_estado.configure(
                    text=f"✅ Imagen guardada como {os.path.basename(destino)}",
                    text_color=STYLE["exito"]
                )
        except Exception:
            pass

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo guardar la imagen: {e}")

