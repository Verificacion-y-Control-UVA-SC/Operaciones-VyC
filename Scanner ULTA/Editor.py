import os
import json
import pandas as pd
import customtkinter as ctk
from tkinter import messagebox
import traceback
import time
from datetime import datetime

# Importar configuración
from Configuracion import cargar_configuracion, guardar_configuracion

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
FONT_TEXT = ("Inter", 10, "bold")

DATA_DIR = "data"
JSON_FILE = os.path.join(DATA_DIR, "base_etiquetado.json")

# ---------------- FUNCIONES DE EDICIÓN ---------------- #
def normalizar_columna(nombre):
    nombre = str(nombre).strip().upper()
    nombre = nombre.replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")
    return nombre

def buscar_archivo_excel():
    """Busca archivos Excel en la carpeta data"""
    data_dir = "data"
    if not os.path.exists(data_dir):
        return None
    
    excel_files = [f for f in os.listdir(data_dir) if f.lower().endswith(('.xlsx', '.xls'))]
    
    if excel_files:
        # Ordenar por fecha de modificación (más reciente primero)
        excel_files.sort(key=lambda x: os.path.getmtime(os.path.join(data_dir, x)), reverse=True)
        return os.path.join(data_dir, excel_files[0])
    
    return None

def cargar_json_seguro():
    """Carga el JSON de manera segura, manejando archivos corruptos"""
    try:
        # Verificar si el archivo existe
        if not os.path.exists(JSON_FILE):
            print(f"Archivo JSON no encontrado: {JSON_FILE}")
            return []
        
        # Verificar tamaño del archivo
        file_size = os.path.getsize(JSON_FILE)
        if file_size == 0:
            print("Archivo JSON vacío")
            return []
        
        # Intentar cargar el JSON normalmente
        try:
            with open(JSON_FILE, 'r', encoding='utf-8') as f:
                datos = json.load(f)
            
            if not isinstance(datos, list):
                print("ADVERTENCIA: JSON no es una lista, convirtiendo...")
                if isinstance(datos, dict):
                    datos = [datos]
                else:
                    datos = []
            return datos
            
        except json.JSONDecodeError as e:
            print(f"JSON corrupto detectado: {e}")
            
            # Intentar reparar el JSON
            with open(JSON_FILE, 'r', encoding='utf-8') as f:
                contenido = f.read()
            
            # Método simple de reparación
            lineas = contenido.split('\n')
            lineas_buenas = []
            
            for linea in lineas:
                linea = linea.strip()
                if linea:
                    # Verificar si la línea parece un objeto JSON válido
                    if (linea.startswith('{') and linea.endswith('},')) or \
                       (linea.startswith('{') and linea.endswith('}')):
                        lineas_buenas.append(linea)
                    elif ':' in linea and '"' in linea and '{' in linea:
                        lineas_buenas.append(linea)
            
            # Reconstruir el JSON
            if lineas_buenas:
                contenido_reparado = '[\n' + ',\n'.join(lineas_buenas) + '\n]'
                
                try:
                    datos = json.loads(contenido_reparado)
                    print(f"JSON reparado, {len(datos)} registros recuperados")
                    
                    # Guardar el JSON reparado
                    with open(JSON_FILE, 'w', encoding='utf-8') as f:
                        json.dump(datos, f, ensure_ascii=False, indent=4)
                    
                    return datos
                except:
                    print("No se pudo reparar el JSON, creando uno nuevo")
            
            # Si no se pudo reparar, crear uno nuevo
            datos = []
            with open(JSON_FILE, 'w', encoding='utf-8') as f:
                json.dump(datos, f, ensure_ascii=False, indent=4)
            return datos
                
    except Exception as e:
        print(f"Error inesperado cargando JSON: {e}")
        traceback.print_exc()
        return []

def guardar_json_seguro(datos):
    """Guarda datos en JSON de manera segura y simple"""
    try:
        # Crear directorio si no existe
        os.makedirs(DATA_DIR, exist_ok=True)
        
        # Serializar datos de forma segura
        datos_serializables = []
        for item in datos:
            if isinstance(item, dict):
                item_serializable = {}
                for key, value in item.items():
                    if pd.isna(value) or value is None:
                        item_serializable[key] = ""
                    elif isinstance(value, pd.Timestamp):
                        item_serializable[key] = str(value)
                    elif isinstance(value, (int, float, str, bool)):
                        item_serializable[key] = value
                    else:
                        try:
                            item_serializable[key] = str(value)
                        except:
                            item_serializable[key] = ""
                datos_serializables.append(item_serializable)
        
        # Intentar guardar directamente
        try:
            with open(JSON_FILE, 'w', encoding='utf-8') as f:
                json.dump(datos_serializables, f, ensure_ascii=False, indent=4)
            print("JSON guardado exitosamente")
            return True
        except Exception as e:
            print(f"Error al guardar JSON: {e}")
            return False
                
    except Exception as e:
        print(f"Error crítico guardando JSON: {e}")
        traceback.print_exc()
        return False

def actualizar_json_original(producto_actualizado):
    """Actualiza el archivo JSON original con persistencia completa"""
    try:
        print(f"Actualizando JSON para UPC: {producto_actualizado.get('UPC', 'N/A')}")
        
        # Cargar datos existentes
        productos = cargar_json_seguro()
        
        # Buscar el producto por UPC
        upc = str(producto_actualizado.get("UPC", "")).strip()
        if not upc:
            print("Error: Producto no tiene UPC")
            return False
        
        encontrado = False
        
        for i, producto in enumerate(productos):
            if isinstance(producto, dict):
                producto_upc = str(producto.get("UPC", "")).strip()
                if producto_upc == upc:
                    # Actualizar el producto existente
                    for key, value in producto_actualizado.items():
                        productos[i][key] = value
                    encontrado = True
                    print(f"Producto actualizado en JSON: {upc}")
                    break
        
        # Si no se encontró, agregar como nuevo producto
        if not encontrado:
            productos.append(producto_actualizado)
            print(f"Nuevo producto agregado al JSON: {upc}")
        
        # Guardar el JSON
        if guardar_json_seguro(productos):
            print(f"JSON actualizado exitosamente para UPC {upc}")
            return True
        else:
            print(f"Error guardando JSON para UPC {upc}")
            return False
        
    except Exception as e:
        print(f"Error actualizando JSON original: {e}")
        traceback.print_exc()
        return False

def actualizar_excel_original(producto_actualizado):
    """Actualiza el Excel original con los datos modificados con persistencia"""
    try:
        # Verificar que producto_actualizado sea un diccionario
        if isinstance(producto_actualizado, list):
            if len(producto_actualizado) > 0:
                print("ADVERTENCIA: Se recibió una lista en lugar de un diccionario. Tomando el primer elemento.")
                producto_actualizado = producto_actualizado[0]
            else:
                print("ERROR: Se recibió una lista vacía en actualizar_excel_original")
                return False
        
        if not isinstance(producto_actualizado, dict):
            print(f"ERROR: Se esperaba un diccionario pero se recibió {type(producto_actualizado)}")
            return False
        
        # PRIMERO: Actualizar el JSON original
        json_actualizado = actualizar_json_original(producto_actualizado)
        
        if not json_actualizado:
            print("Error al actualizar JSON original")
            return False
        
        # SEGUNDO: Intentar actualizar el Excel
        config = cargar_configuracion()
        excel_path = config.get("excel_path")
        
        # Si no hay ruta guardada, buscar el archivo más reciente
        if not excel_path or not os.path.exists(excel_path):
            print("Buscando archivo Excel más reciente...")
            excel_path = buscar_archivo_excel()
            
            if not excel_path:
                print("No se encontró archivo Excel en la carpeta data")
                # Ya actualizamos el JSON, así que al menos eso está guardado
                return True
            
            # Actualizar la información del Excel
            config["excel_path"] = excel_path
            config["fecha_actualizacion"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            guardar_configuracion(config)
        
        print(f"Actualizando Excel: {excel_path}")
        
        try:
            # Leer el Excel original manteniendo todos los datos
            df = pd.read_excel(excel_path, dtype=str, keep_default_na=False)
        except Exception as e:
            print(f"Error leyendo Excel: {e}")
            # El JSON ya está actualizado, así que retornamos True
            return True
        
        # Limpiar nombres de columnas pero mantener originales para referencia
        df.columns = [str(col).strip() for col in df.columns]
        
        # Buscar el UPC en el Excel
        upc = str(producto_actualizado.get("UPC", "")).strip()
        
        if not upc:
            print("No se encontró UPC en el producto")
            return True  # JSON ya actualizado
        
        # Crear diccionario de mapeo de columnas normalizadas
        columnas_normalizadas = {}
        for col in df.columns:
            col_normalized = str(col).strip().upper().replace(" ", "_").replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")
            columnas_normalizadas[col_normalized] = col
        
        # Buscar columna que contenga UPC
        upc_columna = None
        posibles_upc = ["UPC", "CODIGOUPC", "CODIGO_UPC", "CODIGO", "EAN", "SKU"]
        
        for posible in posibles_upc:
            if posible in columnas_normalizadas:
                upc_columna = columnas_normalizadas[posible]
                break
        
        if not upc_columna and df.shape[1] > 0:
            # Asumir que la primera columna es el identificador
            upc_columna = df.columns[0]
            print(f"Usando primera columna como UPC: {upc_columna}")
        
        if not upc_columna:
            print("No se pudo encontrar columna para UPC")
            return True  # JSON ya actualizado
        
        # Buscar la fila que coincida con el UPC
        mask = df[upc_columna].astype(str).str.strip() == upc
        
        if mask.any():
            idx = df[mask].index[0]
            print(f"Encontrada fila {idx} para UPC {upc}")
            
            # Actualizar cada campo del producto
            actualizaciones_realizadas = 0
            for campo, valor in producto_actualizado.items():
                if campo == "UPC":  # No actualizar el UPC
                    continue
                    
                valor_str = str(valor).strip() if valor is not None else ""
                
                # Normalizar nombre del campo para búsqueda
                campo_normalized = str(campo).strip().upper().replace(" ", "_").replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")
                
                # Buscar columna que coincida
                columna_encontrada = None
                for col_normalized, col_original in columnas_normalizadas.items():
                    if col_normalized == campo_normalized:
                        columna_encontrada = col_original
                        break
                
                if columna_encontrada:
                    valor_original = str(df.at[idx, columna_encontrada])
                    if valor_str != valor_original:
                        df.at[idx, columna_encontrada] = valor_str
                        print(f"Actualizado {columna_encontrada}: '{valor_original}' -> '{valor_str}'")
                        actualizaciones_realizadas += 1
                else:
                    # Si no se encuentra la columna, agregarla al final
                    print(f"Añadiendo nueva columna: {campo}")
                    df[campo] = ""
                    df.at[idx, campo] = valor_str
                    actualizaciones_realizadas += 1
            
            # Guardar el Excel actualizado
            try:
                df.to_excel(excel_path, index=False)
                print(f"Excel actualizado exitosamente. {actualizaciones_realizadas} campos modificados.")
                return True
            except Exception as e:
                print(f"Error al guardar Excel: {e}")
                # El JSON ya está actualizado
                return True
        else:
            print(f"No se encontró el UPC {upc} en el Excel")
            # El JSON ya está actualizado
            return True
            
    except Exception as e:
        print(f"Error al actualizar Excel original: {e}")
        traceback.print_exc()
        # Aún así retornamos True porque el JSON se actualizó
        return True

# ---------------- CLASE EDITOR ---------------- #
class EditorInformacionComercial(ctk.CTkToplevel):
    def __init__(self, parent, producto, callback_guardar):
        super().__init__(parent)
        self.parent = parent
        self.producto_original = producto.copy() if isinstance(producto, dict) else {}
        self.producto = producto if isinstance(producto, dict) else {}
        self.callback_guardar = callback_guardar
        
        # Añadir flag para controlar estado
        self.is_destroying = False
        
        self.title("✏️ EDITAR INFORMACIÓN COMERCIAL")
        self.geometry("480x640")
        self.configure(fg_color=STYLE["fondo"])
        self.resizable(True, True)
        
        # Configurar icono
        try:
            base_path = os.path.dirname(os.path.abspath(__file__))
            icon_path = os.path.join(base_path, "img", "icon.ico")
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except Exception as e:
            print(f"No se pudo cargar el icono: {e}")
        
        # Forzar enfoque en esta ventana
        self.transient(parent)
        self.grab_set()
        self.focus_set()
        
        # Configurar manejo de cierre
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.crear_interfaz()

    def crear_interfaz(self):
        # Frame principal con scroll
        main_frame = ctk.CTkScrollableFrame(self, fg_color=STYLE["fondo"])
        main_frame.pack(fill="both", expand=True, padx=12, pady=12)
        
        # Título
        titulo = ctk.CTkLabel(
            main_frame, 
            text="✏️ EDITAR INFORMACIÓN COMERCIAL", 
            font=FONT_SUB, 
            text_color=STYLE["header_texto"]
        )
        titulo.pack(anchor="w", pady=(0, 10))
        
        # Campos editables
        campos_config = [
            ("OBSERVACIONES_REVISION", "Observaciones de Revisión"),
            ("CATEGORIA", "Categoría"),
            ("UPC", "Código UPC"),
            ("DENOMINACION", "Denominación"),
            ("DENOMINACION_AXO", "Denominación AXO"),
            ("MARCA", "Marca"),
            ("CONTENIDO", "Contenido"),
            ("TAMAÑO_DE_LA_DECLARACION_DE_CONTENIDO", "Declaración de Tamaño (NOM)"),
            ("PAIS_ORIGEN", "País Origen"),
            ("INSTRUCCIONES_DE_USO", "Instrucciones de uso"),
            ("LEYENDAS_PRECAUTORIAS", "Leyendas precautorias"),
            ("OBSERVACIONES", "Observaciones"),
            ("INGREDIENTES_Y_LOTE", "Ingredientes y Lote"),
            ("MEDIDAS", "Medidas"),
            ("TIPO_DE_ETIQUETA", "Tipo de etiqueta"),
            ("NORMA", "Norma")
        ]
        
        self.widgets = {}
        
        for campo_key, etiqueta in campos_config:
            # Frame para cada campo
            campo_frame = ctk.CTkFrame(main_frame, fg_color=STYLE["surface"], corner_radius=6)
            campo_frame.pack(fill="x", padx=5, pady=5)
            
            # Etiqueta
            label = ctk.CTkLabel(
                campo_frame,
                text=etiqueta,
                font=("Inter", 11, "bold"),
                text_color=STYLE["texto_oscuro"],
                anchor="w"
            )
            label.pack(fill="x", padx=10, pady=(10, 5))
            
            # Widget de entrada
            valor = str(self.producto.get(campo_key, "")).strip() if isinstance(self.producto, dict) else ""
            
            if campo_key in ["INSTRUCCIONES_DE_USO", "LEYENDAS_PRECAUTORIAS", "OBSERVACIONES", "INGREDIENTES_Y_LOTE"]:
                # Para campos largos
                widget = ctk.CTkTextbox(campo_frame, height=80, font=("Inter", 10))
                widget.pack(fill="x", padx=10, pady=(0, 10))
                widget.insert("1.0", valor)
            else:
                # Para campos cortos
                widget = ctk.CTkEntry(
                    campo_frame, 
                    height=35,
                    font=("Inter", 10),
                    placeholder_text=f"Ingrese {etiqueta.lower()}..."
                )
                widget.pack(fill="x", padx=10, pady=(0, 10))
                widget.insert(0, valor)
            
            self.widgets[campo_key] = widget
        
        # Información de estado
        estado_frame = ctk.CTkFrame(main_frame, fg_color=STYLE["surface"], corner_radius=6)
        estado_frame.pack(fill="x", padx=5, pady=10)
        
        self.estado_label = ctk.CTkLabel(
            estado_frame,
            text="✅ Listo para guardar cambios",
            font=("Inter", 10),
            text_color=STYLE["exito"]
        )
        self.estado_label.pack(padx=10, pady=10)
        
        # Botones
        btn_frame = ctk.CTkFrame(main_frame, fg_color=STYLE["fondo"])
        btn_frame.pack(fill="x", pady=20)
        
        # Botón Cancelar
        btn_cancelar = ctk.CTkButton(
            btn_frame,
            text="✖ Cancelar",
            command=self.confirmar_salida,
            width=120,
            height=35,
            fg_color=STYLE["borde"],
            hover_color="#A0A7B0",
            font=("Inter", 11, "bold")
        )
        btn_cancelar.pack(side="left", padx=10)
        
        # Botón Guardar
        btn_guardar = ctk.CTkButton(
            btn_frame,
            text="💾 Guardar Cambios",
            command=self.guardar_cambios,
            width=120,
            height=35,
            fg_color=STYLE["primario"],
            hover_color="#D9C421",
            text_color=STYLE["header_texto"],
            font=("Inter", 11, "bold")
        )
        btn_guardar.pack(side="right", padx=10)
        
        # Enfocar la primera entrada
        if self.widgets:
            self.after(100, lambda: list(self.widgets.values())[0].focus_set())

    def obtener_datos(self):
        """Obtiene los datos actuales de los widgets"""
        datos = {}
        for campo_key, widget in self.widgets.items():
            try:
                if isinstance(widget, ctk.CTkTextbox):
                    valor = widget.get("1.0", "end-1c").strip()
                else:
                    valor = widget.get().strip()
            except Exception:
                valor = ""
            datos[campo_key] = valor
        return datos

    def guardar_cambios(self):
        if self.is_destroying or not self.winfo_exists():
            return
        try:
            # Recopilar datos de todos los widgets
            producto_actualizado = self.obtener_datos()
            
            # Mantener campos que no estaban en los widgets pero sí en el original
            if isinstance(self.producto_original, dict):
                for key, value in self.producto_original.items():
                    if key not in producto_actualizado:
                        producto_actualizado[key] = str(value).strip() if value is not None else ""
            
            # Verificar si hay cambios
            cambios = False
            if isinstance(self.producto_original, dict):
                for key in producto_actualizado:
                    valor_original = str(self.producto_original.get(key, "")).strip()
                    valor_nuevo = str(producto_actualizado.get(key, "")).strip()
                    
                    if valor_original != valor_nuevo:
                        cambios = True
                        print(f"Cambio detectado en {key}: '{valor_original}' -> '{valor_nuevo}'")
                        break
            
            if not cambios:
                response = messagebox.askyesno(
                    "Sin cambios", 
                    "No se detectaron cambios. ¿Desea cerrar el editor?"
                )
                if response:
                    self.is_destroying = True
                    self.destroy()
                return
            
            # Actualizar estado
            self.estado_label.configure(
                text="⏳ Guardando cambios...",
                text_color=STYLE["advertencia"]
            )
            self.update()
            
            # Llamar al callback para guardar los cambios
            if self.callback_guardar:
                self.callback_guardar(producto_actualizado)
            
            # Actualizar estado
            self.estado_label.configure(
                text="✅ Cambios guardados exitosamente",
                text_color=STYLE["exito"]
            )
            
            # Esperar un momento y cerrar
            self.is_destroying = True
            self.after(1500, self.destroy)
            
        except Exception as e:
            if not self.is_destroying and self.winfo_exists():
                messagebox.showerror("Error", f"No se pudieron guardar los cambios:\n{str(e)}")
                print(f"Error en guardar_cambios: {e}")
                traceback.print_exc()
                
                # Restaurar estado
                if hasattr(self, 'estado_label'):
                    self.estado_label.configure(
                        text="❌ Error al guardar",
                        text_color=STYLE["peligro"]
                    )

    def confirmar_salida(self):
        """Confirma si el usuario desea salir sin guardar"""
        # Marcar que estamos destruyendo
        self.is_destroying = True
        
        # Verificar si hay cambios no guardados
        cambios = False
        try:
            producto_actual = self.obtener_datos()
            if isinstance(self.producto_original, dict):
                for key in producto_actual:
                    valor_original = str(self.producto_original.get(key, "")).strip()
                    valor_actual = str(producto_actual.get(key, "")).strip()
                    
                    if valor_original != valor_actual:
                        cambios = True
                        break
        except:
            cambios = False
        
        if cambios:
            response = messagebox.askyesnocancel(
                "Cambios sin guardar",
                "Tienes cambios sin guardar. ¿Deseas guardarlos antes de salir?\n\n"
                "Sí: Guardar y salir\n"
                "No: Salir sin guardar\n"
                "Cancelar: Continuar editando"
            )
            
            if response is True:  # Sí - Guardar y salir
                self.is_destroying = False
                self.guardar_cambios()
            elif response is False:  # No - Salir sin guardar
                self.destroy()
            else:  # Cancelar - Continuar editando
                self.is_destroying = False
        else:
            self.destroy()

    def on_closing(self):
        """Manejador para el cierre de la ventana"""
        self.confirmar_salida()

