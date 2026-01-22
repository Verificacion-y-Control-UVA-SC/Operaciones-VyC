import os
import sys
import json

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.utils import ImageReader
from datetime import datetime
from tkinter import filedialog, messagebox
from io import BytesIO
import matplotlib.pyplot as plt

# Detectar el directorio base compatible con .py y .exe
if getattr(sys, 'frozen', False):
    # Ejecutable: carpeta junto al .exe
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'datos')
CODIGOS_PATH = os.path.join(DATA_DIR, 'codigos_cumple.json')
ARCHIVOS_PROCESADOS_PATH = os.path.join(DATA_DIR, 'archivos_procesados.json')

def guardar_codigos(codigos):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(CODIGOS_PATH, 'w', encoding='utf-8') as f:
        json.dump(codigos, f, ensure_ascii=False, indent=2)

def cargar_codigos():
    if os.path.exists(CODIGOS_PATH):
        with open(CODIGOS_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def guardar_archivos_procesados(lista):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(ARCHIVOS_PROCESADOS_PATH, 'w', encoding='utf-8') as f:
        json.dump(lista, f, ensure_ascii=False, indent=2)

def cargar_archivos_procesados():
    if os.path.exists(ARCHIVOS_PROCESADOS_PATH):
        with open(ARCHIVOS_PROCESADOS_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []

def borrar_archivos_procesados():
    if os.path.exists(ARCHIVOS_PROCESADOS_PATH):
        os.remove(ARCHIVOS_PROCESADOS_PATH)
import tkinter as tk
from tkinter import filedialog, messagebox
import json
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas as pdf_canvas
import os
import sys
from datetime import datetime
import matplotlib.pyplot as plt
from io import BytesIO
from reportlab.lib.utils import ImageReader
import threading
import time
import pandas as pd

# ---------------- Configuración ---------------- #
COL_BG = "#FFFFFF"  # Fondo blanco
COL_TEXT = "#282828"  # Texto oscuro
COL_BTN = "#ECD925"  # Amarillo para botones
COL_LIST_BG = "#d8d8d8"  # Gris claro para lista
COL_BAR = "#ECD925"  # Amarillo para barras
COL_BTN_CERRAR = "#282828"  # Oscuro para botón cerrar

# Colores para las tarjetas
COL_CARD_BG = "#FFFFFF"  # Fondo de tarjetas blanco
COL_BORDER = "#E2E8F0"  # Bordes grises suaves
COL_SUCCESS = "#4CAF50"  # Verde para "Cumple"
COL_DANGER = "#F44336"  # Rojo para "No cumple"
COL_TEXT_LIGHT = "#666666"  # Texto secundario

# ------------------ Rutas seguras para PyInstaller ------------------
def recurso_path(ruta_relativa):
    """Devuelve la ruta absoluta correcta dentro del ejecutable o desarrollo"""
    try:
        base_path = sys._MEIPASS  # PyInstaller
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, ruta_relativa)

# Rutas de archivos
ARCHIVO_JSON = CODIGOS_PATH
ARCHIVO_EXCEL = os.path.join(DATA_DIR, "codigos_cumple.xlsx")
CONFIG_DIR = DATA_DIR
ARCHIVOS_PROCESADOS_FILE = ARCHIVOS_PROCESADOS_PATH
LOGO_PATH = recurso_path("img/logo_empresarial.png")

# Crear directorios si no existen
os.makedirs(CONFIG_DIR, exist_ok=True)

# Lista global
archivos_procesados = []

# Variables globales para las etiquetas
lbl_total_valor = None
lbl_cumple_valor = None
lbl_cumple_porcentaje = None
lbl_no_cumple_valor = None
lbl_no_cumple_porcentaje = None
canvas_grafica = None
lst_archivos = None
lbl_totales = None

# ---------------- Sistema de Monitoreo ---------------- #
class MonitorCambios:
    def __init__(self, intervalo=2):  # 2 segundos entre verificaciones
        self.intervalo = intervalo
        self.ultima_modificacion_json = 0
        self.ultima_modificacion_excel = 0
        self.ejecutando = False
        self.thread = None
    
    def iniciar_monitoreo(self):
        """Inicia el monitoreo en segundo plano"""
        self.ejecutando = True
        self.thread = threading.Thread(target=self._monitorear, daemon=True)
        self.thread.start()
    
    def detener_monitoreo(self):
        """Detiene el monitoreo"""
        self.ejecutando = False
    
    def _monitorear(self):
        """Monitorea cambios en los archivos"""
        while self.ejecutando:
            try:
                # Verificar cambios en JSON
                if os.path.exists(ARCHIVO_JSON):
                    mod_time = os.path.getmtime(ARCHIVO_JSON)
                    if mod_time > self.ultima_modificacion_json:
                        self.ultima_modificacion_json = mod_time
                        print("📁 Cambio detectado en JSON - Actualizando dashboard...")
                        # Actualizar interfaz desde el hilo principal
                        root.after(0, actualizar_interfaz_completa)
                
                # Verificar cambios en Excel
                if os.path.exists(ARCHIVO_EXCEL):
                    mod_time = os.path.getmtime(ARCHIVO_EXCEL)
                    if mod_time > self.ultima_modificacion_excel:
                        self.ultima_modificacion_excel = mod_time
                        print("📁 Cambio detectado en Excel - Actualizando dashboard...")
                        root.after(0, actualizar_interfaz_completa)
                
                time.sleep(self.intervalo)
                
            except Exception as e:
                print(f"Error en monitoreo: {e}")
                time.sleep(self.intervalo)

# Instancia global del monitor
monitor = MonitorCambios()

# ---------------- Funciones ---------------- #

class Dashboard:
    def __init__(self):
        self.archivo_json = ARCHIVO_JSON
        self.ultima_modificacion = 0
        self.cargar_datos()
        self.iniciar_verificacion()
    
    def iniciar_verificacion(self):
        """Verifica periódicamente si el archivo ha cambiado"""
        self.verificar_cambios()
        # Verificar cada 2 segundos
        self.root.after(2000, self.iniciar_verificacion)
    
    def verificar_cambios(self):
        """Verifica si el archivo JSON ha sido modificado"""
        try:
            if os.path.exists(self.archivo_json):
                mod_time = os.path.getmtime(self.archivo_json)
                if mod_time > self.ultima_modificacion:
                    self.ultima_modificacion = mod_time
                    self.cargar_datos()
                    print("Dashboard actualizado automáticamente")
        except Exception as e:
            print(f"Error al verificar cambios: {e}")
    
    def cargar_datos(self):
        """Carga los datos desde el archivo JSON"""
        try:
            if os.path.exists(self.archivo_json):
                with open(self.archivo_json, 'r', encoding='utf-8') as f:
                    datos = json.load(f)
                
                # Convertir a DataFrame
                self.df_codigos = pd.DataFrame(datos)
                
                # Actualizar estadísticas
                self.actualizar_estadisticas()
                
                # Actualizar la interfaz
                self.actualizar_interfaz()
        except Exception as e:
            print(f"Error al cargar datos: {e}")
    
    def actualizar_estadisticas(self):
        """Actualiza las estadísticas del dashboard"""
        total_codigos = len(self.df_codigos)
        codigos_cumple = len(self.df_codigos[self.df_codigos['OBSERVACIONES'].str.upper() == 'CUMPLE'])
        otros_codigos = total_codigos - codigos_cumple
        
        print(f"Total: {total_codigos}, Cumple: {codigos_cumple}, Otros: {otros_codigos}")
        
        # Aquí actualizas tus labels o gráficos
        if hasattr(self, 'label_total'):
            self.label_total.config(text=f"Total: {total_codigos}")
        if hasattr(self, 'label_cumple'):
            self.label_cumple.config(text=f"Cumple: {codigos_cumple}")
        if hasattr(self, 'label_otros'):
            self.label_otros.config(text=f"Otros: {otros_codigos}")


def actualizar_interfaz_completa():
    """Actualiza toda la interfaz con los datos más recientes"""
    global lbl_total_valor, lbl_cumple_valor, lbl_cumple_porcentaje
    global lbl_no_cumple_valor, lbl_no_cumple_porcentaje, canvas_grafica, lst_archivos, lbl_totales
    
    try:
        total_codigos, codigos_cumple, codigos_no_cumple = leer_datos()
        
        # Actualizar tarjetas
        if lbl_total_valor:
            lbl_total_valor.config(text=f"{total_codigos}")
        if lbl_cumple_valor:
            lbl_cumple_valor.config(text=f"{codigos_cumple}")
        if lbl_no_cumple_valor:
            lbl_no_cumple_valor.config(text=f"{codigos_no_cumple}")
        
        # Calcular porcentajes
        porcentaje_cumple = (codigos_cumple / total_codigos * 100) if total_codigos > 0 else 0
        porcentaje_no_cumple = (codigos_no_cumple / total_codigos * 100) if total_codigos > 0 else 0
        
        if lbl_cumple_porcentaje:
            lbl_cumple_porcentaje.config(text=f"{porcentaje_cumple:.1f}%")
        if lbl_no_cumple_porcentaje:
            lbl_no_cumple_porcentaje.config(text=f"{porcentaje_no_cumple:.1f}%")
        
        if lbl_totales:
            lbl_totales.config(text=f"Total: {total_codigos}  |  Cumple: {codigos_cumple}  |  No cumple: {codigos_no_cumple}")
        
        # Redibujar gráfica
        if canvas_grafica:
            dibujar_grafica_estatica(canvas_grafica, total_codigos, codigos_cumple, codigos_no_cumple)
        
        # Actualizar lista de archivos
        if lst_archivos:
            actualizar_lista_archivos(lst_archivos)
            
    except Exception as e:
        print(f"Error actualizando interfaz: {e}")

def cargar_archivos_procesados():
    """Carga la lista de archivos procesados desde el JSON, crea el archivo si no existe"""
    global archivos_procesados
    try:
        if os.path.exists(ARCHIVOS_PROCESADOS_FILE):
            with open(ARCHIVOS_PROCESADOS_FILE, "r", encoding="utf-8") as f:
                datos = json.load(f)
                if isinstance(datos, list):
                    archivos_procesados = datos
                else:
                    archivos_procesados = []
                    print("Formato inválido en archivo de procesados")
        else:
            archivos_procesados = []
            # Crear archivo vacío
            with open(ARCHIVOS_PROCESADOS_FILE, "w", encoding="utf-8") as f:
                json.dump([], f, indent=4, ensure_ascii=False)
            print(f"Archivo {ARCHIVOS_PROCESADOS_FILE} no encontrado. Se creó uno nuevo.")
    except Exception as e:
        archivos_procesados = []
        print(f"Error cargando archivos procesados: {e}")
    
    return archivos_procesados

def guardar_archivos_procesados():
    """Guarda la lista de archivos procesados en el archivo JSON"""
    try:
        with open(ARCHIVOS_PROCESADOS_FILE, "w", encoding="utf-8") as f:
            json.dump(archivos_procesados, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Error guardando archivos procesados: {e}")

def registrar_archivo_procesado(nombre_archivo, fecha_proceso):
    """Registra un archivo procesado en el JSON"""
    try:
        cargar_archivos_procesados()
        
        # Evitar duplicados
        if any(a["nombre"] == nombre_archivo for a in archivos_procesados):
            print(f"[INFO] Archivo ya registrado: {nombre_archivo}")
            return
        
        # Agregar nuevo archivo
        archivo_info = {
            "nombre": nombre_archivo,
            "fecha_proceso": fecha_proceso,
            "fecha_archivo": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        archivos_procesados.append(archivo_info)
        
        # Guardar cambios en el JSON
        guardar_archivos_procesados()
        
        print(f"[OK] Archivo registrado: {nombre_archivo}")
    
    except Exception as e:
        print(f"[ERROR] Error registrando archivo: {e}")

def borrar_archivo_procesados():
    """Elimina físicamente el archivo JSON de archivos procesados"""
    try:
        if os.path.exists(ARCHIVOS_PROCESADOS_FILE):
            os.remove(ARCHIVOS_PROCESADOS_FILE)
            return True
        return False
    except Exception as e:
        print(f"❌ Error al borrar archivo de procesados: {e}")
        return False

def actualizar_lista_archivos(lst_archivos):
    """Actualiza la lista visual con los archivos procesados"""
    lst_archivos.delete(0, tk.END)
    cargar_archivos_procesados()  # Asegurarse de tener datos actualizados
    
    for archivo in archivos_procesados:
        # Si el archivo es un diccionario, mostrar nombre y fecha
        if isinstance(archivo, dict) and 'nombre' in archivo:
            nombre = archivo['nombre']
            fecha = archivo.get('fecha_proceso', archivo.get('fecha_archivo', 'Fecha desconocida'))
            # Formatear la fecha si es necesario
            if isinstance(fecha, str) and len(fecha) > 10:
                try:
                    fecha_dt = datetime.strptime(fecha, "%Y-%m-%d %H:%M:%S")
                    fecha = fecha_dt.strftime("%d/%m/%Y %H:%M")
                except:
                    pass
            lst_archivos.insert(tk.END, f"{nombre} - {fecha}")
        # Si es un string (solo nombre), mostrarlo tal cual
        elif isinstance(archivo, str):
            lst_archivos.insert(tk.END, archivo)
        else:
            # Mostrar representación string para otros tipos
            lst_archivos.insert(tk.END, str(archivo))

def eliminar_archivo_seleccionado(lst_archivos):
    """Elimina el archivo seleccionado de la lista y del disco"""
    global archivos_procesados
    
    seleccion = lst_archivos.curselection()
    if not seleccion:
        messagebox.showwarning("Selección requerida", "Por favor seleccione un archivo de la lista para eliminar")
        return
    
    indice = seleccion[0]
    
    # Obtener información del archivo
    archivo_info = archivos_procesados[indice]
    if isinstance(archivo_info, dict) and 'nombre' in archivo_info and 'ruta' in archivo_info:
        nombre_archivo = archivo_info['nombre']
        ruta_archivo = archivo_info['ruta']
    else:
        nombre_archivo = str(archivo_info)
        ruta_archivo = str(archivo_info)
    
    # Confirmar eliminación
    respuesta = messagebox.askyesno(
        "Confirmar eliminación",
        f"¿Está seguro de que desea eliminar el archivo:\n\n{nombre_archivo}\n\nEsta acción no se puede deshacer."
    )
    if not respuesta:
        return
    
    # Eliminar archivo físico si existe
    if os.path.exists(ruta_archivo):
        try:
            os.remove(ruta_archivo)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo eliminar el archivo del disco:\n{str(e)}")
            return
    
    # Eliminar el archivo de la lista
    archivos_procesados.pop(indice)
    
    # Guardar cambios en JSON
    guardar_archivos_procesados()
    
    # Actualizar la lista visual
    actualizar_lista_archivos(lst_archivos)
    
    messagebox.showinfo("Eliminado", f"Archivo eliminado correctamente: {nombre_archivo}")

def limpiar_lista(lst_archivos):
    """Limpia la lista de archivos procesados y elimina el archivo JSON"""
    global archivos_procesados
    archivos_procesados = []

    # Intentar borrar el JSON
    if borrar_archivo_procesados():
        messagebox.showinfo("🗑️ Lista Limpiada", "Se han eliminado todos los archivos de la lista")
    else:
        # Si no se pudo borrar el archivo, al menos guardar lista vacía
        guardar_archivos_procesados()
        messagebox.showinfo("🗑️ Lista Limpiada", "Se han eliminado todos los archivos de la lista")

    # Actualizar visualización de la lista en la interfaz
    actualizar_lista_archivos(lst_archivos)

def leer_datos():
    """Lee los datos del archivo JSON and calcula estadísticas"""
    total_codigos = 0
    codigos_cumple = 0
    codigos_no_cumple = 0
    
    try:
        print(f"Leyendo archivo JSON desde: {ARCHIVO_JSON}")
        
        if not os.path.exists(ARCHIVO_JSON):
            print("❌ El archivo JSON no existe. Creando uno vacío...")
            # Crear archivo JSON vacío si no existe
            with open(ARCHIVO_JSON, "w", encoding="utf-8") as f:
                json.dump([], f, indent=4, ensure_ascii=False)
            return 0, 0, 0
        
        # Intentar leer desde JSON primero
        with open(ARCHIVO_JSON, "r", encoding="utf-8") as f:
            codigos_data = json.load(f)
            print(f"Datos leídos desde JSON: {len(codigos_data)} registros")
            
        for d in codigos_data:
            if not isinstance(d, dict):
                continue
                
            total_codigos += 1
            obs = str(d.get("OBSERVACIONES", "")).upper().strip()
            
            if obs == "CUMPLE":
                codigos_cumple += 1
            else:
                codigos_no_cumple += 1
                
        print(f"Estadísticas: Total={total_codigos}, Cumple={codigos_cumple}, No Cumple={codigos_no_cumple}")
        
    except json.JSONDecodeError as e:
        print(f"❌ Error: El archivo JSON está corrupto o vacío: {e}")
        # Intentar leer desde Excel como respaldo
        return leer_datos_desde_excel()
    except Exception as e:
        print(f"❌ Error leyendo JSON: {e}")
        return leer_datos_desde_excel()
        
    return total_codigos, codigos_cumple, codigos_no_cumple

def leer_datos_desde_excel():
    """Lee datos desde Excel como respaldo"""
    try:
        if not os.path.exists(ARCHIVO_EXCEL):
            print("❌ El archivo Excel tampoco existe")
            return 0, 0, 0
            
        df = pd.read_excel(ARCHIVO_EXCEL)
        total_codigos = len(df)
        codigos_cumple = len(df[df["OBSERVACIONES"].astype(str).str.upper() == "CUMPLE"])
        codigos_no_cumple = total_codigos - codigos_cumple
        
        print(f"Datos leídos desde Excel: {total_codigos} registros")
        return total_codigos, codigos_cumple, codigos_no_cumple
        
    except Exception as e:
        print(f"❌ Error leyendo Excel: {e}")
        return 0, 0, 0

def dibujar_grafica_estatica(canvas, total_codigos, codigos_cumple, codigos_no_cumple):
    """Dibuja la gráfica con datos específicos (versión estática)"""
    try:
        canvas.delete("all")
        
        # Solo dibujar gráfica si hay datos
        if total_codigos > 0:
            # --- Datos para las barras ---
            datos = [
                ("Total de Códigos", total_codigos),
                ("Códigos Cumple", codigos_cumple),
                ("Códigos No cumple", codigos_no_cumple)
            ]

            # --- Ajustes de espacio dinámicos ---
            ancho, alto = int(canvas["width"]), int(canvas["height"])
            margen_sup = 30
            margen_inf = 60
            margen_lat = 20
            ancho_barra = 80
            espacio = 60

            altura_max = alto - (margen_sup + margen_inf)
            max_valor = max([v for _, v in datos], default=1)

            # --- Dibujar ejes ---
            eje_x_y = alto - margen_inf
            canvas.create_line(margen_lat, eje_x_y, ancho - margen_lat, eje_x_y, fill=COL_TEXT, width=2)
            canvas.create_line(margen_lat, margen_sup, margen_lat, eje_x_y, fill=COL_TEXT, width=2)

            # --- Dibujar barras ---
            x_inicio = margen_lat + espacio
            for i, (nombre, valor) in enumerate(datos):
                altura_barra = (valor / max_valor) * altura_max if valor > 0 else 0
                x1 = x_inicio + i * (ancho_barra + espacio)
                y1 = eje_x_y - altura_barra
                x2 = x1 + ancho_barra
                y2 = eje_x_y
                
                # Color por categoría
                if nombre == "Códigos Cumple":
                    color = COL_SUCCESS
                elif nombre == "Códigos No cumple":
                    color = COL_DANGER
                else:
                    color = COL_BAR
                    
                # Barra
                canvas.create_rectangle(x1, y1, x2, y2, fill=color, outline=COL_TEXT, width=1.5)
                # Valor encima
                canvas.create_text((x1 + x2) / 2, y1 - 10, text=str(valor), font=("INTER", 9, "bold"), fill=COL_TEXT)
                # Etiqueta abajo
                canvas.create_text((x1 + x2) / 2, eje_x_y + 20, text=nombre, font=("INTER", 8, "bold"), 
                                  fill=COL_TEXT, width=100, justify='center')
        else:
            # Mostrar mensaje si no hay datos
            ancho, alto = int(canvas["width"]), int(canvas["height"])
            canvas.create_text(ancho/2, alto/2, text="No hay datos disponibles\nVerifique el archivo JSON", 
                              font=("INTER", 10), fill=COL_TEXT_LIGHT, justify='center')

    except Exception as e:
        print(f"Error en dibujar_grafica_estatica: {e}")

def dibujar_grafica(canvas, lbl_totales_ref, lst_archivos_ref):
    """Función original para compatibilidad (ahora usa la versión estática)"""
    total_codigos, codigos_cumple, codigos_no_cumple = leer_datos()
    
    # Actualizar labels
    if lbl_total_valor:
        lbl_total_valor.config(text=f"{total_codigos}")
    if lbl_cumple_valor:
        lbl_cumple_valor.config(text=f"{codigos_cumple}")
    if lbl_no_cumple_valor:
        lbl_no_cumple_valor.config(text=f"{codigos_no_cumple}")
    
    porcentaje_cumple = (codigos_cumple / total_codigos * 100) if total_codigos > 0 else 0
    porcentaje_no_cumple = (codigos_no_cumple / total_codigos * 100) if total_codigos > 0 else 0
    
    if lbl_cumple_porcentaje:
        lbl_cumple_porcentaje.config(text=f"{porcentaje_cumple:.1f}%")
    if lbl_no_cumple_porcentaje:
        lbl_no_cumple_porcentaje.config(text=f"{porcentaje_no_cumple:.1f}%")
    
    if lbl_totales_ref:
        lbl_totales_ref.config(text=f"Total: {total_codigos}  |  Cumple: {codigos_cumple}  |  No cumple: {codigos_no_cumple}")

    # Dibujar gráfica
    dibujar_grafica_estatica(canvas, total_codigos, codigos_cumple, codigos_no_cumple)
    
    # Actualizar lista de archivos
    actualizar_lista_archivos(lst_archivos_ref)
    
    # Programar próxima actualización
    canvas.after(2000, lambda: dibujar_grafica(canvas, lbl_totales_ref, lst_archivos_ref))

def crear_tarjeta(parent, titulo, valor, porcentaje=None, color=COL_BAR):
    """Crea una tarjeta de estadística moderna"""
    frame = tk.Frame(parent, bg=COL_CARD_BG, relief="flat", bd=1, 
                    highlightbackground=COL_BORDER, highlightthickness=1)
    
    # Título
    lbl_titulo = tk.Label(frame, text=titulo, bg=COL_CARD_BG, fg=COL_TEXT_LIGHT, 
                         font=("INTER", 9))
    lbl_titulo.pack(pady=(8, 3))
    
    # Valor principal
    lbl_valor = tk.Label(frame, text=valor, bg=COL_CARD_BG, fg=color, 
                        font=("INTER", 14, "bold"))
    lbl_valor.pack(pady=3)
    
    # Porcentaje (opcional)
    if porcentaje:
        lbl_porcentaje = tk.Label(frame, text=porcentaje, bg=COL_CARD_BG, fg=COL_TEXT_LIGHT,
                                 font=("INTER", 8))
        lbl_porcentaje.pack(pady=(0, 8))
    
    return frame, lbl_valor, lbl_porcentaje if porcentaje else (frame, lbl_valor, None)

# ---------------- Exportar PDF ---------------- 
def exportar_pdf_simple():
    """Genera un PDF simple con estadísticas, varias páginas, encabezado, footer y numeración"""
    try:
        total_codigos, codigos_cumple, codigos_no_cumple = leer_datos()
        porcentaje_cumple = (codigos_cumple / total_codigos * 100) if total_codigos > 0 else 0
        porcentaje_no_cumple = (codigos_no_cumple / total_codigos * 100) if total_codigos > 0 else 0

        ruta = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Archivos PDF", "*.pdf")],
            title="Guardar Reporte de Estadísticas"
        )
        if not ruta:
            return

        c = pdf_canvas.Canvas(ruta, pagesize=letter)
        ancho, alto = letter
        pagina_actual = 1

        # Función para dibujar encabezado
        def dibujar_encabezado(titulo):
            c.setFillColor("#ecd925")
            c.rect(0, alto - 20, ancho, 20, fill=1, stroke=0)
            c.setFillColor("#282828")
            c.setFont("Helvetica-Bold", 16)
            c.drawString(50, alto - 50, titulo)
            # Logo
            try:
                if os.path.exists(LOGO_PATH):
                    logo = ImageReader(LOGO_PATH)
                    c.drawImage(logo, ancho - 100, alto - 70, width=50, height=50, preserveAspectRatio=True)
            except:
                pass
            c.setFont("Helvetica", 10)

        # Función para dibujar footer
        def dibujar_footer(pagina):
            c.setFillColor("#282828")
            c.rect(0, 0, ancho, 30, fill=1, stroke=0)
            c.setFillColor("#FFFFFF")
            c.setFont("Helvetica", 8)
            c.drawString(50, 15, "Sistema de Tipos de Procesos V&C")
            texto_centro = "www.vandc.com"
            ancho_texto_centro = c.stringWidth(texto_centro, "Helvetica", 8)
            c.drawString((ancho - ancho_texto_centro) / 2, 15, texto_centro)
            texto_derecho = f"Página {pagina}"
            ancho_texto_derecho = c.stringWidth(texto_derecho, "Helvetica", 8)
            c.drawString(ancho - ancho_texto_derecho - 50, 15, texto_derecho)

        # Función para crear nueva página
        def nueva_pagina(titulo):
            nonlocal pagina_actual, y
            dibujar_footer(pagina_actual)
            c.showPage()
            pagina_actual += 1
            dibujar_encabezado(titulo)
            y = alto - 100
            return y

        # Iniciar primera página
        y = alto - 120
        dibujar_encabezado("REPORTE DE ESTADÍSTICAS")
        c.drawString(50, alto - 70, f"Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

        # Estadísticas principales
        c.setFont("Helvetica-Bold", 12)
        c.drawString(50, y, "ESTADÍSTICAS PRINCIPALES")
        y -= 30
        c.setFont("Helvetica", 10)
        lineas = [
            f"• Total de códigos: {total_codigos}",
            f"• Códigos que cumplen: {codigos_cumple} ({porcentaje_cumple:.1f}%)",
            f"• Códigos que no cumplen: {codigos_no_cumple} ({porcentaje_no_cumple:.1f}%)"
        ]
        for linea in lineas:
            if y < 100:
                y = nueva_pagina("REPORTE DE ESTADÍSTICAS")
            c.drawString(70, y, linea)
            y -= 20

        # Archivos procesados
        archivos = archivos_procesados
        if archivos:
            if y < 100:
                y = nueva_pagina("REPORTE DE ESTADÍSTICAS")
            c.setFont("Helvetica-Bold", 12)
            c.drawString(50, y, "ARCHIVOS PROCESADOS")
            y -= 25

            # <-- Aquí agregas el total de archivos procesados -->
            total_archivos = len(archivos)
            c.setFont("Helvetica", 10)
            c.drawString(70, y, f"Total de archivos procesados: {total_archivos}")
            y -= 20

            # Listado de archivos
            c.setFont("Helvetica-Bold", 12)
            c.drawString(50, y, "CARGA SEMANAL:")
            y -=  25

            c.setFont("Helvetica",10)
            for archivo in archivos:
                if y < 100:
                    y = nueva_pagina("REPORTE DE ESTADÍSTICAS")
                nombre = archivo if isinstance(archivo, str) else archivo.get('nombre', str(archivo))
                c.drawString(70, y, f"• {nombre}")
                y -= 15
            y -= 10


        # Gráfica de pastel
        if total_codigos > 0:
            if y < 350:
                y = nueva_pagina("REPORTE DE ESTADÍSTICAS (gráfica)")
            etiquetas = ["Códigos Cumple", "Códigos No Cumple"]
            valores = [codigos_cumple, codigos_no_cumple]
            colores = ["#ECD925", "#282828"]
            porcentajes = [porcentaje_cumple, porcentaje_no_cumple]

            plt.figure(figsize=(8, 6))
            wedges, texts, autotexts = plt.pie(
                valores, labels=etiquetas, colors=colores,
                autopct='%1.1f%%', startangle=90, textprops={'fontsize': 12, 'color': '#282828'}
            )
            for autotext in autotexts:
                autotext.set_color('white')
                autotext.set_fontweight('bold')
                autotext.set_fontsize(12)
            for text in texts:
                text.set_fontsize(12)
                text.set_fontweight('bold')
            plt.title("Distribución de Códigos", fontsize=16, fontweight='bold', color="#282828", pad=20)
            plt.axis('equal')
            leyenda_labels = [f'{etiqueta}: {valor} ({porcentaje:.1f}%)'
                              for etiqueta, valor, porcentaje in zip(etiquetas, valores, porcentajes)]
            plt.legend(wedges, leyenda_labels, title="Estadísticas", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
            plt.tight_layout()

            buf = BytesIO()
            plt.savefig(buf, format="PNG", dpi=150, bbox_inches='tight')
            plt.close()
            buf.seek(0)
            imagen_grafica = ImageReader(buf)
            c.drawImage(imagen_grafica, 50, y - 280, width=500, height=280)

        # Footer de la última página
        dibujar_footer(pagina_actual)

        c.save()
        messagebox.showinfo("Éxito", f"PDF generado correctamente en:\n{ruta}")

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo generar el PDF:\n{e}")
        print(f"Error detallado: {e}")



def on_closing():
    # Detener monitoreo
    monitor.detener_monitoreo()
    
    # Guardar archivos procesados antes de salir
    try:
        with open(ARCHIVOS_PROCESADOS_FILE, 'w', encoding='utf-8') as f:
            json.dump(archivos_procesados, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Error guardando archivos al cerrar: {e}")
    root.destroy()



# ---------------- Ventana principal ---------------- #

def main():
    global lbl_total_valor, lbl_cumple_valor, lbl_cumple_porcentaje, lbl_no_cumple_valor, lbl_no_cumple_porcentaje
    global canvas_grafica, lst_archivos, lbl_totales, root

    root = tk.Tk()
    root.title("Dashboard de Códigos - V&C")
    root.geometry("1000x600")
    root.configure(bg=COL_BG)

    # Iniciar monitoreo de cambios
    monitor.iniciar_monitoreo()

    # Cargar archivos procesados al iniciar
    cargar_archivos_procesados()

    # Frame principal
    main_container = tk.Frame(root, bg=COL_BG)
    main_container.pack(fill="both", expand=True, padx=15, pady=15)

    # Header
    header_frame = tk.Frame(main_container, bg=COL_BG)
    header_frame.pack(fill="x", pady=(0, 10))

    lbl_titulo = tk.Label(header_frame, text="📊 Dashboard de Análisis de Códigos",
                          bg=COL_BG, fg=COL_TEXT, font=("INTER", 16, "bold"))
    lbl_titulo.pack(side="left")

    lbl_subtitulo = tk.Label(header_frame, text="Reporte de Mercancía - V&C",
                             bg=COL_BG, fg=COL_TEXT_LIGHT, font=("INTER", 10))
    lbl_subtitulo.pack(side="left", padx=(10, 0))

    # Tarjetas
    stats_frame = tk.Frame(main_container, bg=COL_BG)
    stats_frame.pack(fill="x", pady=(0, 10))

    tarjeta_total, lbl_total_valor, _ = crear_tarjeta(stats_frame, "TOTAL DE CÓDIGOS", "0", color=COL_BAR)
    tarjeta_total.pack(side="left", padx=(0, 10), fill="both", expand=True)

    tarjeta_cumple, lbl_cumple_valor, lbl_cumple_porcentaje = crear_tarjeta(stats_frame, "CÓDIGOS CUMPLEN", "0", "0%", color=COL_SUCCESS)
    tarjeta_cumple.pack(side="left", padx=(0, 10), fill="both", expand=True)

    tarjeta_no_cumple, lbl_no_cumple_valor, lbl_no_cumple_porcentaje = crear_tarjeta(stats_frame, "CÓDIGOS NO CUMPLEN", "0", "0%", color=COL_DANGER)
    tarjeta_no_cumple.pack(side="left", fill="both", expand=True)

    # Contenido principal
    content_frame = tk.Frame(main_container, bg=COL_BG)
    content_frame.pack(fill="both", expand=True, pady=10)

    # Frame izquierdo para la gráfica
    left_frame = tk.Frame(content_frame, bg=COL_BG)
    left_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))

    # Gráfica
    graph_card = tk.Frame(left_frame, bg=COL_CARD_BG, relief="flat", bd=1,
                          highlightbackground=COL_BORDER, highlightthickness=1)
    graph_card.pack(fill="both", expand=True)

    lbl_graph_title = tk.Label(graph_card, text="DISTRIBUCIÓN DE CÓDIGOS",
                               bg=COL_CARD_BG, fg=COL_TEXT, font=("INTER", 11, "bold"))
    lbl_graph_title.pack(pady=(10, 5))

    canvas_grafica = tk.Canvas(graph_card, width=400, height=250,
                               bg=COL_CARD_BG, highlightthickness=0)
    canvas_grafica.pack(pady=(0, 5), padx=10, fill="both", expand=True)

    lbl_totales = tk.Label(graph_card, text="", bg=COL_CARD_BG,
                           fg=COL_TEXT_LIGHT, font=("INTER", 9))
    lbl_totales.pack(pady=(0, 10))

    # Frame derecho para archivos

    right_frame = tk.Frame(content_frame, bg=COL_BG, width=450)
    right_frame.pack(side="right", fill="y")
    right_frame.pack_propagate(False)

    files_card = tk.Frame(right_frame, bg=COL_CARD_BG, relief="flat", bd=1,
                        highlightbackground=COL_BORDER, highlightthickness=1)
    files_card.pack(fill="both", expand=True)

    lbl_archivos = tk.Label(files_card, text="📁 ARCHIVOS PROCESADOS",
                            bg=COL_CARD_BG, fg=COL_TEXT, font=("INTER", 11, "bold"))
    lbl_archivos.pack(pady=(10, 5))

    # Frame principal para la lista
    list_frame = tk.Frame(files_card, bg=COL_CARD_BG)
    list_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    # Frame para encabezados
    header_frame = tk.Frame(list_frame, bg=COL_CARD_BG)
    header_frame.pack(fill="x", pady=(0, 5))

    # Encabezado para nombre de archivo (izquierda)
    lbl_nombre = tk.Label(header_frame, text="ARCHIVO", 
                        bg=COL_CARD_BG, fg=COL_TEXT, 
                        font=("INTER", 9, "bold"), anchor="w")
    lbl_nombre.pack(side="left", fill="x", expand=True)

    # Encabezado para fecha (derecha)
    lbl_fecha = tk.Label(header_frame, text="FECHA", 
                        bg=COL_CARD_BG, fg=COL_TEXT, 
                        font=("INTER", 9, "bold"), anchor="e", width=10)
    lbl_fecha.pack(side="right")

    # Frame para lista y scrollbar
    list_content_frame = tk.Frame(list_frame, bg=COL_CARD_BG)
    list_content_frame.pack(fill="both", expand=True)

    scrollbar = tk.Scrollbar(list_content_frame)
    scrollbar.pack(side="right", fill="y")

    # Listbox con formato mejorado
    lst_archivos = tk.Listbox(list_content_frame, 
                            bg=COL_LIST_BG, 
                            fg=COL_TEXT, 
                            font=("INTER", 9),
                            yscrollcommand=scrollbar.set, 
                            relief="flat", 
                            bd=0,
                            highlightthickness=0,
                            justify="left")
    lst_archivos.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=lst_archivos.yview)

    # Función para agregar archivos con formato de dos columnas
    def agregar_archivo_con_fecha(nombre_archivo, fecha):
        # Formatear la línea: nombre a la izquierda, fecha a la derecha
        nombre_truncado = nombre_archivo[:35] + "..." if len(nombre_archivo) > 38 else nombre_archivo
        espacio_disponible = 50 - len(nombre_truncado)  # Ajustar según el ancho del Listbox
        espacios = " " * max(espacio_disponible, 5)  # Mínimo 5 espacios
        linea = f"{nombre_truncado}{espacios}{fecha}"
        lst_archivos.insert(tk.END, linea)

    # Footer con botones
    footer_frame = tk.Frame(main_container, bg=COL_BG)
    footer_frame.pack(fill="x", pady=(10, 0))

    # Botón para eliminar archivo seleccionado
    btn_eliminar = tk.Button(
        footer_frame,
        text="🗑️ Eliminar Archivo",
        command=lambda: eliminar_archivo_seleccionado(lst_archivos),
        bg="#D9534F",  # color rojo corporativo
        fg="white",
        font=("INTER", 9, "bold"),
        relief="flat",
        padx=15,
        pady=6,
        cursor="hand2"
    )
    btn_eliminar.pack(side="left", padx=(0, 5))

    btn_limpiar = tk.Button(footer_frame, text="🗑️ Limpiar Lista",
                            command=lambda: limpiar_lista(lst_archivos),
                            bg=COL_TEXT_LIGHT, fg="white", font=("INTER", 9, "bold"),
                            relief="flat", padx=15, pady=6, cursor="hand2")
    btn_limpiar.pack(side="left", padx=(0, 5))

    btn_exportar = tk.Button(footer_frame, text="📊 Exportar PDF",
                             command=exportar_pdf_simple,
                             bg=COL_BTN, fg=COL_TEXT, font=("INTER", 9, "bold"),
                             relief="flat", padx=15, pady=6, cursor="hand2")
    btn_exportar.pack(side="left", padx=(0, 5))

    btn_cerrar = tk.Button(footer_frame, text="❌ Cerrar",
                           command=lambda: [monitor.detener_monitoreo(), root.destroy()],
                           bg=COL_BTN_CERRAR, fg="white", font=("INTER", 9, "bold"),
                           relief="flat", padx=15, pady=6, cursor="hand2")
    btn_cerrar.pack(side="right")

    # Dibujar gráfica inicial
    dibujar_grafica(canvas_grafica, lbl_totales, lst_archivos)


    # Centrar ventana
    root.eval('tk::PlaceWindow . center')
    root.mainloop()

if __name__ == "__main__":
    main()
