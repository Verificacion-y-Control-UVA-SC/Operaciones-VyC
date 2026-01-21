import os
import json
import shutil
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
from Bernier import dibujar_regla_bernier

# Importar módulos de edición
from Editor import EditorInformacionComercial
from Normas import obtener_puntos_normativos 
from Configuracion import (cargar_configuracion, guardar_configuracion, 
                          seleccionar_carpeta_imagenes, cargar_base_excel, 
                          buscar_producto_por_upc, mostrar_imagen, extraer_valor_numerico, 
                          subir_imagen, obtener_campo)

# Importar módulos de edición
from Editor import EditorInformacionComercial
from Normas import obtener_puntos_normativos 
from Configuracion import (cargar_configuracion, guardar_configuracion, 
                          seleccionar_carpeta_imagenes, cargar_base_excel, 
                          buscar_producto_por_upc, mostrar_imagen, extraer_valor_numerico, 
                          subir_imagen, obtener_campo)
# Agregar importación para editor de facturación
try:
    from EditorFacturacion import EditorFacturacion
except ImportError:
    pass  # Se manejará más adelante

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

# Mapeo para compatibilidad con código existente
STYLE.update({
    "bg": STYLE["fondo"],
    "label": STYLE["primario"],
    "label2": STYLE["hover_primario"],
    "text": STYLE["texto_oscuro"],
    "button": STYLE["secundario"],
    "button_hover": "#3a3a3a",
    "entry": STYLE["surface"],
    "secondary": STYLE["fondo"],
    "border": STYLE["borde"]
})

DATA_DIR = "data"
JSON_FILE = os.path.join(DATA_DIR, "base_etiquetado.json")
CONFIG_FILE = os.path.join(DATA_DIR, "config.json")
FACTURA_FILE = os.path.join(DATA_DIR, "factura_actual.json")
LAYOUT_FILE = os.path.join(DATA_DIR, "layout_actual.json")  
os.makedirs(DATA_DIR, exist_ok=True)

# ---------------- CLASE PRINCIPAL ---------------- #
class EscanerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("📷 Escáner ULTA")
        self.geometry("1200x600")
        self.minsize(1200, 600)
        self.configure(fg_color=STYLE["fondo"])
        self.bind("<Configure>", self.on_resize)

        # Establecer icono
        try:
            base_path = os.path.dirname(os.path.abspath(__file__))
            icon_path = os.path.join(base_path, "img", "icon.ico")
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except Exception as e:
            print(f"No se pudo cargar el icono: {e}")

        # Inicializar variables
        self.productos = []
        self.codigo_actual = ctk.StringVar()
        self.producto_actual = None
        # Contador de escaneos por UPC
        self.contador_escaneos = {}
        self.excel_path = None
        self.factura_data = {}    # Datos de la factura cargada como diccionario
        self.layout_data = {}     # Datos del layout cargado

        # Cargar configuración
        self.config = cargar_configuracion()
        
        # Cargar productos desde JSON
        self.cargar_productos()
        
        # Cargar factura si existe
        self.cargar_factura()
        
        # Cargar layout si existe
        self.cargar_layout()
        
        self.cambios_factura = []
        self.cambios_layout = []



        # Obtener ruta del Excel desde configuración
        self.obtener_ruta_excel()
        
        # Inicializar entrada
        self.codigo_actual.set("")

        # Configurar grid
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Crear interfaz
        self.header = self.crear_header()
        self.header.grid(row=0, column=0, sticky="ew")
        # Actualizar contador de pendientes en la UI principal (si aplica)
        try:
            self.actualizar_pendientes_main()
        except Exception:
            pass
        self.body = self.crear_body()
        self.body.grid(row=1, column=0, sticky="nsew")
        self.footer = self.crear_footer()
        self.footer.grid(row=2, column=0, sticky="ew")

        # Centrar ventana
        self.center_window()

    def obtener_ruta_excel(self):
        """Obtiene la ruta del Excel desde config.json"""
        try:
            # Cargar configuraciones
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
                    # Buscar la ruta del Excel en config.json
                    self.excel_path = config_data.get('excel_path')
                    
                    # Si no está en config.json, buscar en data/
                    if not self.excel_path or not os.path.exists(self.excel_path):
                        self.buscar_excel_en_data()
            else:
                self.buscar_excel_en_data()
                
        except Exception as e:
            print(f"Error al obtener ruta del Excel: {e}")
            self.buscar_excel_en_data()

    def buscar_excel_en_data(self):
        """Busca archivos Excel en la carpeta data"""
        excel_files = []
        for file in os.listdir(DATA_DIR):
            if file.lower().endswith(('.xlsx', '.xls')):
                excel_files.append(os.path.join(DATA_DIR, file))
        
        if excel_files:
            # Tomar el más reciente
            excel_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            self.excel_path = excel_files[0]
            
            # Actualizar config.json
            try:
                config_data = {}
                if os.path.exists(CONFIG_FILE):
                    with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                        config_data = json.load(f)
                
                config_data['excel_path'] = self.excel_path
                with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                    json.dump(config_data, f, indent=4, ensure_ascii=False)
                    
                print(f"Ruta del Excel actualizada: {self.excel_path}")
            except Exception as e:
                print(f"Error actualizando config.json: {e}")

    def cargar_productos(self):
        """Carga productos desde el JSON"""
        try:
            if os.path.exists(JSON_FILE):
                with open(JSON_FILE, 'r', encoding='utf-8') as f:
                    contenido = f.read().strip()
                    
                    if contenido:
                        self.productos = json.loads(contenido)
                        print(f"✅ {len(self.productos)} productos cargados desde JSON")
                    else:
                        self.productos = []
                        print("JSON vacío, iniciando lista vacía")
            else:
                self.productos = []
                print("JSON no encontrado, iniciando lista vacía")
                
        except json.JSONDecodeError as e:
            print(f"Error cargando JSON: {e}")
            messagebox.showwarning("Error", "El archivo JSON está corrupto. Se iniciará con lista vacía.")
            self.productos = []
        except Exception as e:
            print(f"Error inesperado: {e}")
            self.productos = []

    def cargar_factura(self):
        """Carga datos de factura desde JSON si existe"""
        try:
            if os.path.exists(FACTURA_FILE):
                with open(FACTURA_FILE, 'r', encoding='utf-8') as f:
                    contenido = f.read().strip()
                    
                    if contenido:
                        self.factura_data = json.loads(contenido)
                        print(f"✅ Factura cargada: {len(self.factura_data.get('items', []))} registros")
                        # Normalizar claves de cantidad para compatibilidad
                        items = self.factura_data.get('items', [])
                        for registro in items:
                            # Normalizar valores 'nan' y crear clave CANTIDAD_FACTURA
                            for k, v in list(registro.items()):
                                if v is None:
                                    registro[k] = ''
                                else:
                                    sv = str(v).strip()
                                    if sv.lower() == 'nan':
                                        registro[k] = ''
                                    else:
                                        registro[k] = sv

                            # Crear clave estandarizada para la cantidad facturada
                            registro['CANTIDAD_FACTURA'] = (
                                registro.get('CANT. FACT.') or
                                registro.get('CANT. FACT') or
                                registro.get('CANTIDAD_FACTURA') or
                                registro.get('CANTIDAD EN VU') or
                                ''
                            )
                            # Intentar extraer UPC real si la columna 'UPC' no contiene el código
                            try:
                                import re
                                upc_raw = str(registro.get('UPC','')).strip()
                                upc_candidate = ''
                                if not upc_raw or not re.fullmatch(r"\d{6,13}", upc_raw):
                                    # Revisar campos comunes que contienen '1 - 8101376...'
                                    for key in ['# ORDEN - ITEM', 'ORDEN - ITEM', '# ORDEN - ITEM', 'ORDEN - ITEM', 'UPC']:
                                        val = str(registro.get(key, '')).strip()
                                        if ' - ' in val:
                                            parts = val.split(' - ')
                                            last = parts[-1].strip()
                                            if re.fullmatch(r"\d{6,13}", last):
                                                upc_candidate = last
                                                break
                                    # Si no se encontró, buscar en todos los valores cualquier grupo de dígitos razonable
                                    if not upc_candidate:
                                        for v in registro.values():
                                            s = str(v).strip()
                                            m = re.search(r"(\d{6,13})", s)
                                            if m:
                                                upc_candidate = m.group(1)
                                                break
                                if upc_candidate:
                                    registro['UPC'] = upc_candidate
                            except Exception:
                                pass
                        # Guardar cambios normalizados para evitar re-procesos posteriores
                        try:
                            self.guardar_factura()
                        except Exception:
                            pass
                        # Deduplicar items por UPC (preservando el primer registro)
                        try:
                            seen = set()
                            unique_items = []
                            for reg in items:
                                upc = str(reg.get('UPC','')).strip()
                                if not upc:
                                    # skip entries without UPC — they don't have a usable code
                                    continue
                                if upc in seen:
                                    continue
                                seen.add(upc)
                                unique_items.append(reg)
                            # Reassign only if dedup changed
                            if len(unique_items) != len(items):
                                self.factura_data['items'] = unique_items
                                try:
                                    self.guardar_factura()
                                except Exception:
                                    pass
                        except Exception:
                            pass
                        # Ensure every item has a FACTURA field (prefer per-item, else file name)
                        try:
                            nombre_archivo = self.factura_data.get('nombre_sin_ext') or self.factura_data.get('nombre_archivo','')
                            for reg in self.factura_data.get('items', []):
                                if not reg.get('FACTURA'):
                                    reg['FACTURA'] = nombre_archivo or ''
                            # Persist final normalized list (only items with UPC)
                            try:
                                self.guardar_factura()
                            except Exception:
                                pass
                        except Exception:
                            pass
                        
                        # Si la UI ya fue creada, actualizar contadores y tarjetas
                        try:
                            if hasattr(self, 'label_pendientes_main'):
                                self.actualizar_pendientes_main()
                        except Exception:
                            pass
                        try:
                            if hasattr(self, 'footer_label'):
                                self.actualizar_footer()
                        except Exception:
                            pass
                        try:
                            if hasattr(self, 'label_cantidad_facturada'):
                                self.actualizar_tarjeta_informacion()
                        except Exception:
                            pass
                    else:
                        self.factura_data = {"items": [], "nombre_archivo": ""}
                        print("Factura JSON vacío")
            else:
                self.factura_data = {"items": [], "nombre_archivo": ""}
                print("No hay factura cargada")

            # 🔑 Inicializar contador de escaneos desde factura
            self.contador_escaneos = {}

            for item in self.factura_data.get("items", []):
                upc = self.normalizar_upc(item.get("UPC"))
                if upc:
                    self.contador_escaneos[upc] = 0

            # 🔄 Refrescar contador visual
            self.actualizar_pendientes_main()

                
        except json.JSONDecodeError as e:
            print(f"Error cargando factura JSON: {e}")
            self.factura_data = {"items": [], "nombre_archivo": ""}
        except Exception as e:
            print(f"Error inesperado al cargar factura: {e}")
            self.factura_data = {"items": [], "nombre_archivo": ""}

    def cargar_layout(self):
        """Carga datos de layout desde JSON si existe"""
        try:
            if os.path.exists(LAYOUT_FILE):
                with open(LAYOUT_FILE, 'r', encoding='utf-8') as f:
                    contenido = f.read().strip()
                    
                    if contenido:
                        self.layout_data = json.loads(contenido)
                        datos_count = len(self.layout_data.get('datos', []))
                        print(f"✅ Layout cargado: {self.layout_data.get('nombre_archivo', 'Sin nombre')}")
                        print(f"   • {datos_count} registros cargados")
                        print(f"   • Total etiquetas: {self.layout_data.get('total_etiquetas', 0)}")
                    else:
                        self.layout_data = {}
                        print("Layout JSON vacío")
            else:
                self.layout_data = {}
                print("No hay layout cargado")
                
        except json.JSONDecodeError as e:
            print(f"Error cargando layout JSON: {e}")
            self.layout_data = {}
        except Exception as e:
            print(f"Error inesperado al cargar layout: {e}")
            self.layout_data = {}

    def guardar_factura(self):
        """Guarda los datos de factura en JSON"""
        try:
            with open(FACTURA_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.factura_data, f, ensure_ascii=False, indent=4)
            print(f"✅ Factura guardada: {len(self.factura_data.get('items', []))} registros")
        except Exception as e:
            print(f"Error guardando factura: {e}")
            messagebox.showerror("Error", f"No se pudo guardar la factura:\n{e}")

    def guardar_layout(self):
        try:
            with open(LAYOUT_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.layout_data, f, ensure_ascii=False, indent=4)
            datos_count = len(self.layout_data.get('datos', []))
            columnas_encontradas = len(self.layout_data.get('columnas_encontradas', []))
            print(f"✅ Layout guardado: {datos_count} registros, {columnas_encontradas} columnas")
        except Exception as e:
            print(f"Error guardando layout: {e}")
            messagebox.showerror("Error", f"No se pudo guardar el layout:\n{e}")

    def cargar_factura_archivo(self):
        """Carga un archivo de factura y extrae únicamente las columnas requeridas"""
        try:
            file_path = filedialog.askopenfilename(
                title="Seleccionar archivo de factura",
                filetypes=[
                    ("Excel files", "*.xlsx *.xls"),
                    ("CSV files", "*.csv"),
                    ("All files", "*.*")
                ]
            )

            if not file_path:
                return

            # Leer archivo
            if file_path.lower().endswith((".xlsx", ".xls")):
                df = pd.read_excel(file_path, dtype=str)
            elif file_path.lower().endswith(".csv"):
                df = pd.read_csv(file_path, dtype=str, encoding="utf-8")
            else:
                messagebox.showerror("Error", "Formato de archivo no soportado")
                return

            # Limpiar encabezados
            df.columns = [c.strip() for c in df.columns]

            # Columnas requeridas EXACTAS
            columnas_requeridas = [
                "# PARTIDA - FRACCION", "FACTURA", "# ORDEN - ITEM", "UPC", "MARCA", "PAIS",
                "DESC. FACTURA", "DESC. PEDIMENTO", "UNIDAD VU", "CANTIDAD EN VU",
                "UNI. TARIFA", "CANT. TARIFA", "UNI. FACT.", "CANT. FACT.",
                "PRECIO UNIT.", "TOTAL", "FRACCION", "NICO", "FRACCION CORRELACION",
                "NICO CORRELACION", "CORRELACION", "MET. VAL.", "VINCULACION",
                "DESCUENTO", "MODELO", "SERIE", "ESPECIAL",
                "IMPUESTO", "TASA", "IMPORTE", "TT", "P/I",
                "C1", "C2", "C3", "NUM", "FIRMA"
            ]

            # Detectar columnas existentes (case-insensitive)
            columnas_encontradas = {}
            for col_req in columnas_requeridas:
                for col_df in df.columns:
                    if col_req.lower() == col_df.lower():
                        columnas_encontradas[col_req] = col_df
                        break

            factura_items = []

            for _, row in df.iterrows():
                registro = {}

                for col_req, col_real in columnas_encontradas.items():
                    valor = row.get(col_real, "")
                    if pd.isna(valor):
                        valor = ""
                    registro[col_req] = str(valor).strip()

                # Ignorar filas sin UPC
                if not registro.get("UPC"):
                    continue

                factura_items.append(registro)

            # Guardar datos
            nombre_archivo = os.path.basename(file_path)
            nombre_sin_ext = os.path.splitext(nombre_archivo)[0]

            self.factura_data = {
                "items": factura_items,
                "nombre_archivo": nombre_archivo,
                "nombre_sin_ext": nombre_sin_ext,
                "ruta_archivo": file_path,
                "fecha_carga": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
            }

            self.guardar_factura()

            # Actualizar UI
            self.actualizar_footer()
            self.actualizar_tarjeta_informacion()

            messagebox.showinfo(
                "Factura cargada",
                f"Archivo: {nombre_archivo}\n"
                f"Registros procesados: {len(factura_items)}\n"
                f"Columnas extraídas: {len(columnas_encontradas)}/{len(columnas_requeridas)}"
            )

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar la factura:\n{e}")
            print(f"❌ Error cargando factura: {e}")

    def cargar_layout_archivo(self):
        """Carga un archivo de layout y extrae los datos relevantes desde la hoja 'Layout 1'"""
        try:
            # Abrir diálogo para seleccionar archivo
            file_path = filedialog.askopenfilename(
                title="Seleccionar archivo de Layout",
                filetypes=[
                    ("Excel files", "*.xlsx *.xls"),
                    ("All files", "*.*")
                ]
            )
            
            if not file_path:
                return
            
            # Leer archivo Excel
            try:
                # Usar ExcelFile para inspeccionar hojas
                xls = pd.ExcelFile(file_path)
                print(f"✅ Hojas disponibles en el archivo: {xls.sheet_names}")
                
                # Buscar EXACTAMENTE la hoja "Layout 1"
                sheet_name = None
                for sheet in xls.sheet_names:
                    if sheet.strip() == "Layout 1":
                        sheet_name = sheet
                        print(f"✅ Hoja 'Layout 1' encontrada: {sheet}")
                        break
                
                if not sheet_name:
                    messagebox.showerror(
                        "Error", 
                        f"No se encontró la hoja 'Layout 1' en el archivo.\n\n"
                        f"Hojas disponibles:\n{', '.join(xls.sheet_names)}"
                    )
                    return
                
                # Leer SOLO la hoja "Layout 1" con header=2 (tercera fila como encabezados)
                df = pd.read_excel(
                    xls, 
                    sheet_name=sheet_name, 
                    header=2,  # Fila 3 (0-indexed) contiene los encabezados
                    dtype=str
                )
                
                
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo leer la hoja 'Layout 1':\n{e}")
                print(f"❌ Error leyendo hoja 'Layout 1': {e}")
                import traceback
                traceback.print_exc()
                return
            
            # LISTA DE COLUMNAS ESPECÍFICAS DEL LAYOUT QUE NECESITAMOS
            columnas_requeridas = [
                "Folio de Solicitud",
                "NOM",
                "Número de Acreditación", 
                "RFC",
                "Denominación social o nombre",
                "Tipo de persona",
                "Marca del producto",
                "Descripción del producto",
                "Fracción arancelaria",
                "Fecha de envío de la solicitud",
                "Vigencia de la Solicitud",
                "Modalidad de etiquetado",
                "Modelo",
                "UMC",
                "Cantidad",
                "Número de etiquetas a verificar",
                "Parte",
                "Partida",
                "Pais Origen",
                "Pais Comprador"
            ]
            
            print(f"\n🔍 BUSCANDO COLUMNAS ESPECÍFICAS DEL LAYOUT...")
            
            # Mapear nombres normalizados a nombres originales
            mapeo_columnas = {}
            columnas_encontradas = []
            columnas_faltantes = []
            
            for col_requerida in columnas_requeridas:
                encontrada = False
                # Normalizar nombre requerido para comparación
                col_requerida_normalizada = col_requerida.lower().strip()
                
                # Buscar en las columnas del DataFrame
                for col_df in df.columns:
                    col_df_normalizada = str(col_df).lower().strip()
                    
                    # Comparar diferentes formas del nombre
                    if (col_requerida_normalizada == col_df_normalizada or
                        col_requerida_normalizada in col_df_normalizada or
                        col_df_normalizada in col_requerida_normalizada):
                        
                        mapeo_columnas[col_requerida] = col_df
                        columnas_encontradas.append(col_requerida)
                        print(f"✅ '{col_requerida}' -> '{col_df}'")
                        encontrada = True
                        break
                
                if not encontrada:
                    columnas_faltantes.append(col_requerida)
                    print(f"❌ '{col_requerida}' NO ENCONTRADA")
            
            # Mostrar resumen de búsqueda
            print(f"\n📊 RESUMEN DE BÚSQUEDA:")
            print(f"   • Columnas requeridas: {len(columnas_requeridas)}")
            print(f"   • Columnas encontradas: {len(columnas_encontradas)}")
            print(f"   • Columnas faltantes: {len(columnas_faltantes)}")
            
            if columnas_faltantes:
                print(f"   • Faltantes: {columnas_faltantes}")
            
            # Verificar que tenemos las columnas esenciales
            columnas_esenciales = ["Número de etiquetas a verificar", "Cantidad", "Folio de Solicitud", "NOM"]
            columnas_esenciales_faltantes = []
            
            for col_esencial in columnas_esenciales:
                if col_esencial not in columnas_encontradas:
                    columnas_esenciales_faltantes.append(col_esencial)
            
            if columnas_esenciales_faltantes:
                messagebox.showwarning(
                    "⚠️ Columnas esenciales faltantes",
                    f"Las siguientes columnas esenciales no se encontraron:\n\n" +
                    "\n".join(columnas_esenciales_faltantes) +
                    f"\n\nColumnas disponibles:\n" +
                    "\n".join([f"- {col}" for col in df.columns[:20]]) +
                    ("\n..." if len(df.columns) > 20 else "")
                )
            
            # BUSCAR COLUMNA PARA CÁLCULO DE TOTAL
            columna_a_usar = None
            columna_nombre = ""
            
            # Prioridad 1: Número de etiquetas a verificar
            if "Número de etiquetas a verificar" in mapeo_columnas:
                columna_a_usar = mapeo_columnas["Número de etiquetas a verificar"]
                columna_nombre = "Número de etiquetas a verificar"
            # Prioridad 2: Cantidad
            elif "Cantidad" in mapeo_columnas:
                columna_a_usar = mapeo_columnas["Cantidad"]
                columna_nombre = "Cantidad"
            else:
                # Buscar cualquier columna que tenga "etiquetas" o "cantidad"
                for col_df in df.columns:
                    col_df_str = str(col_df).lower()
                    if 'etiqueta' in col_df_str or 'cantidad' in col_df_str:
                        columna_a_usar = col_df
                        columna_nombre = str(col_df)
                        break
            
            if not columna_a_usar:
                messagebox.showerror(
                    "Error",
                    "No se encontró ninguna columna para calcular el total de etiquetas.\n\n"
                    "Columnas disponibles:\n" +
                    "\n".join([f"- {col}" for col in df.columns])
                )
                return
            
            print(f"\n📊 COLUMNA PARA CÁLCULO DE TOTAL:")
            print(f"   • Usando: '{columna_nombre}' -> '{columna_a_usar}'")
            
            # Obtener el total de etiquetas/cantidad
            total_etiquetas = 0
            try:
                # Mostrar algunos valores de la columna para depuración
                print(f"\n🔢 VALORES EN LA COLUMNA '{columna_a_usar}':")
                valores_columna = df[columna_a_usar].astype(str).str.strip()
                
                # Mostrar los primeros 10 valores no vacíos
                valores_no_vacios = valores_columna[valores_columna != '']
                valores_no_vacios = valores_no_vacios[valores_no_vacios != 'nan']
                valores_no_vacios = valores_no_vacios[valores_no_vacios != 'None']
                
                for i, valor in enumerate(valores_no_vacios.head(10)):
                    print(f"   Fila {i}: '{valor}'")
                
                # Convertir a numérico
                valores_numericos = pd.to_numeric(valores_columna, errors='coerce')
                
                print(f"\n🧮 CONVERSIÓN NUMÉRICA:")
                print(f"   • Valores convertidos exitosamente: {valores_numericos.notna().sum()}")
                print(f"   • Valores que fallaron: {valores_numericos.isna().sum()}")
                
                # Calcular suma
                total_etiquetas = valores_numericos.sum()
                
                if pd.isna(total_etiquetas):
                    total_etiquetas = 0
                else:
                    total_etiquetas = int(total_etiquetas)
                    
                print(f"\n✅ TOTAL CALCULADO: {total_etiquetas}")
                
            except Exception as e:
                print(f"⚠️ Error en conversión numérica: {e}")
                total_etiquetas = 0
            
            # Obtener nombre del archivo
            nombre_archivo = os.path.basename(file_path)
            nombre_sin_ext = os.path.splitext(nombre_archivo)[0]
            
            # PREPARAR DATOS DEL LAYOUT PARA JSON
            datos_layout = []
            
            # Procesar cada fila del DataFrame
            for index, row in df.iterrows():
                fila_dict = {}
                tiene_datos = False
                
                # Extraer solo las columnas que encontramos
                for col_requerida in columnas_encontradas:
                    if col_requerida in mapeo_columnas:
                        col_df = mapeo_columnas[col_requerida]
                        valor = row[col_df]
                        
                        # Convertir a string y limpiar
                        if pd.isna(valor):
                            valor_str = None
                        else:
                            valor_str = str(valor).strip()
                            if valor_str and valor_str.lower() != 'nan':
                                tiene_datos = True
                        
                        fila_dict[col_requerida] = valor_str
                    else:
                        fila_dict[col_requerida] = None
                
                # Solo agregar filas que tengan datos en al menos una columna
                if tiene_datos:
                    # Agregar información adicional
                    fila_dict['_row_index'] = index
                    datos_layout.append(fila_dict)
            
            print(f"\n💾 DATOS PROCESADOS:")
            print(f"   • Filas totales en DataFrame: {len(df)}")
            print(f"   • Filas con datos guardadas: {len(datos_layout)}")
            
            # Mostrar ejemplo de datos procesados
            if datos_layout:
                print(f"\n📄 EJEMPLO DE DATOS PROCESADOS (primera fila):")
                for key, value in datos_layout[0].items():
                    if value and key != '_row_index':
                        print(f"   • {key}: {value[:50]}{'...' if len(str(value)) > 50 else ''}")
            
            # GUARDAR DATOS DEL LAYOUT
            self.layout_data = {
                'nombre_archivo': nombre_archivo,
                'nombre_sin_ext': nombre_sin_ext,
                'total_etiquetas': total_etiquetas,
                'columna_utilizada': columna_nombre,
                'columna_nombre_real': str(columna_a_usar),
                'ruta_archivo': file_path,
                'fecha_carga': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S'),
                'hoja': sheet_name,
                'datos': datos_layout,
                'columnas_encontradas': columnas_encontradas,
                'columnas_faltantes': columnas_faltantes,
                'mapeo_columnas': mapeo_columnas,
                'estadisticas': {
                    'total_filas': len(df),
                    'filas_con_datos': len(datos_layout),
                    'columnas_requeridas': len(columnas_requeridas),
                    'columnas_encontradas': len(columnas_encontradas)
                }
            }
            
            # Guardar layout en JSON
            self.guardar_layout()
            
            # Mostrar resumen final
            print(f"\n📋 RESUMEN FINAL:")
            print(f"   • Archivo: {nombre_archivo}")
            print(f"   • Hoja: {sheet_name}")
            print(f"   • Total etiquetas: {total_etiquetas}")
            print(f"   • Registros guardados: {len(datos_layout)}")
            
            # Mostrar mensaje de éxito o advertencia
            if total_etiquetas > 0:
                messagebox.showinfo(
                    "✅ Layout Cargado Exitosamente",
                    f"Archivo: {nombre_archivo}\n"
                    f"Hoja: Layout 1\n"
                    f"Total de etiquetas: {total_etiquetas}\n"
                    f"Columnas encontradas: {len(columnas_encontradas)}/{len(columnas_requeridas)}"
                )
            else:
                messagebox.showwarning(
                    "⚠️ Layout cargado con observación (puede continuar)",
                    f"Archivo: {nombre_archivo}\n"
                    f"Hoja: Layout 1\n\n"
                    f"El sistema detectó que la información requerida no se encontraba "
                    f"en la columna esperada.\n"
                    f"Sin embargo, fue localizada correctamente en otra columna válida, "
                    f"por lo que el proceso continuará sin afectar el funcionamiento.\n\n"
                    f"📦 Total de etiquetas procesadas: {total_etiquetas}\n"
                    f"📄 Registros válidos: {len(datos_layout)}\n"
                    f"📊 Columnas detectadas: {len(columnas_encontradas)}/{len(columnas_requeridas)}\n\n"
                    f"Columna utilizada para el cálculo: '{columna_nombre}'.\n\n"
                    f"⚠️ Recomendación: Verifique el layout para mantener la estructura estándar."
                )

            
            # Actualizar estado
            self.label_estado.configure(
                text=f"✅ Layout cargado: {total_etiquetas} etiquetas ({len(columnas_encontradas)} columnas)",
                text_color=STYLE["exito"] if total_etiquetas > 0 else STYLE["advertencia"]
            )
            self.after(5000, lambda: self.label_estado.configure(text=""))
            
            # ✅ IMPORTANTE: Actualizar contador de pendientes
            self.actualizar_pendientes_main()
            
            # Actualizar footer
            self.actualizar_footer()
            
            # Actualizar la tarjeta de información
            self.actualizar_tarjeta_informacion()
            
            # Actualizar vista si hay producto actual
            if self.producto_actual:
                self.actualizar_vista_producto(self.producto_actual)
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el layout:\n{str(e)[:200]}")
            print(f"❌ Error general cargando layout: {e}")
            import traceback
            traceback.print_exc()

    def limpiar_factura(self):
        """Limpia los datos de la factura cargada"""
        if not self.factura_data.get('items'):
            messagebox.showinfo("Información", "No hay factura cargada para limpiar.")
            return
        
        if messagebox.askyesno("Confirmar", "¿Está seguro de que desea limpiar los datos de la factura?"):
            self.factura_data = {"items": [], "nombre_archivo": ""}
            self.guardar_factura()
            self.actualizar_footer()
            
            # ✅ AGREGADO: Actualizar contador de pendientes
            self.actualizar_pendientes_main()
            
            # Actualizar la tarjeta de información
            self.actualizar_tarjeta_informacion()
            
            if self.producto_actual:
                self.actualizar_vista_producto(self.producto_actual)
            
            self.label_estado.configure(
                text="✅ Factura eliminada",
                text_color=STYLE["exito"]
            )
            self.after(3000, lambda: self.label_estado.configure(text=""))
            messagebox.showinfo("Éxito", "Datos de factura eliminados correctamente.")

    def limpiar_factura(self):
        """Limpia los datos de la factura cargada"""
        if not self.factura_data.get('items'):
            messagebox.showinfo("Información", "No hay factura cargada para limpiar.")
            return
        
        if messagebox.askyesno("Confirmar", "¿Está seguro de que desea limpiar los datos de la factura?"):
            self.factura_data = {"items": [], "nombre_archivo": ""}
            self.guardar_factura()
            self.actualizar_footer()
            
            # ✅ ACTUALIZAR: Llamar a actualizar_pendientes_main()
            self.actualizar_pendientes_main()
            
            # Actualizar la tarjeta de información
            self.actualizar_tarjeta_informacion()
            
            if self.producto_actual:
                self.actualizar_vista_producto(self.producto_actual)
            
            self.label_estado.configure(
                text="✅ Factura eliminada",
                text_color=STYLE["exito"]
            )
            self.after(3000, lambda: self.label_estado.configure(text=""))
            messagebox.showinfo("Éxito", "Datos de factura eliminados correctamente.")

    def limpiar_layout(self):
        """Limpia los datos del layout cargado"""
        if not self.layout_data:
            messagebox.showinfo("Información", "No hay layout cargado para limpiar.")
            return
        
        if messagebox.askyesno("Confirmar", "¿Está seguro de que desea limpiar los datos del layout?"):
            self.layout_data = {}
            self.guardar_layout()
            self.actualizar_footer()
            
            # ✅ ACTUALIZAR: Llamar a actualizar_pendientes_main()
            self.actualizar_pendientes_main()
            
            # Actualizar la tarjeta de información
            self.actualizar_tarjeta_informacion()
            
            if self.producto_actual:
                self.actualizar_vista_producto(self.producto_actual)
            
            self.label_estado.configure(
                text="✅ Layout eliminado",
                text_color=STYLE["exito"]
            )
            self.after(3000, lambda: self.label_estado.configure(text=""))
            messagebox.showinfo("Éxito", "Datos de layout eliminados correctamente.")

    def limpiar_todo(self):
        """Limpia tanto la factura como el layout de una vez"""
        # Verificar si hay algo para limpiar
        hay_factura = self.factura_data.get('items', [])
        hay_layout = self.layout_data
        
        if not hay_factura and not hay_layout:
            messagebox.showinfo("Información", "No hay factura ni layout cargados para limpiar.")
            return
        
        # Mostrar confirmación
        mensaje_confirmacion = "¿Está seguro de que desea limpiar "
        if hay_factura and hay_layout:
            mensaje_confirmacion += "tanto la factura como el layout?"
        elif hay_factura:
            mensaje_confirmacion += "la factura?"
        else:
            mensaje_confirmacion += "el layout?"

        if messagebox.askyesno("Confirmar", mensaje_confirmacion):
            # Limpiar factura si existe
            if hay_factura:
                self.factura_data = {"items": [], "nombre_archivo": ""}
                self.guardar_factura()
                # Limpiar cambios persistentes relacionados con factura
                try:
                    self.cambios_factura = []
                    import glob
                    for f in glob.glob(os.path.join('data', 'cambios_factura_*.json')):
                        try:
                            os.remove(f)
                        except Exception:
                            pass
                except Exception:
                    pass
            
            # Limpiar layout si existe
            if hay_layout:
                self.layout_data = {}
                self.guardar_layout()
                # Limpiar cambios persistentes relacionados con layout
                try:
                    self.cambios_layout = []
                    import glob
                    for f in glob.glob(os.path.join('data', 'cambios_layout_*.json')):
                        try:
                            os.remove(f)
                        except Exception:
                            pass
                except Exception:
                    pass
            
            # Actualizar footer
            self.actualizar_footer()
            
            # ✅ ACTUALIZAR: Resetear contador de escaneos y llamar a actualizar_pendientes_main()
            try:
                self.contador_escaneos = {}
            except Exception:
                pass
            self.actualizar_pendientes_main()
            
            # Actualizar la tarjeta de información
            self.actualizar_tarjeta_informacion()
            
            # Actualizar vista si hay producto actual
            if self.producto_actual:
                self.actualizar_vista_producto(self.producto_actual)
            
            # Mostrar mensaje de éxito
            mensaje_exito = "✅ "
            if hay_factura and hay_layout:
                mensaje_exito += "Factura y layout eliminados"
            elif hay_factura:
                mensaje_exito += "Factura eliminada"
            else:
                mensaje_exito += "Layout eliminado"
            
            self.label_estado.configure(
                text=mensaje_exito,
                text_color=STYLE["exito"]
            )
            self.after(3000, lambda: self.label_estado.configure(text=""))
            
            messagebox.showinfo("Éxito", "Datos eliminados correctamente.") 

    def normalizar_upc(self, valor):
        import re
        if not valor:
            return ""
        match = re.search(r"\d{6,14}", str(valor))
        return match.group() if match else ""

    def mostrar_cantidad_factura(self, item_factura):
        """Muestra la cantidad de factura para el producto actual"""
        try:
            # Soportar varias claves posibles que puedan venir del Excel
            cantidad = (
                item_factura.get('CANT. FACT.') or
                item_factura.get('CANTIDAD_FACTURA') or
                item_factura.get('CANTIDAD EN VU') or
                item_factura.get('CANT. TARIFA') or
                'N/A'
            )

            # Actualizar etiqueta principal de la tarjeta INFORMACIÓN FACTURA
            if hasattr(self, 'label_cantidad_facturada'):
                self.label_cantidad_facturada.configure(text=f"Cantidad Facturada: {cantidad}")

            # También actualizar la información comercial (label_info) si aplica
            if self.producto_actual and hasattr(self, 'label_info'):
                info_actual = self.label_info.cget("text")
                if "Cantidad Factura:" not in info_actual:
                    nueva_info = info_actual + f"\n\n📦 Cantidad Factura: {cantidad}"
                    self.label_info.configure(text=nueva_info)

            # No modificar pedimento/factura aquí; la tarjeta principal se encarga
                
        except Exception as e:
            print(f"Error mostrando cantidad de factura: {e}")

    def buscar_en_factura(self, upc):
        """Busca un UPC en los datos de factura cargados"""
        factura_items = self.factura_data.get('items', [])
        if not factura_items:
            return None
        import re
        def extract_digits_candidate(s):
            s = str(s or "")
            matches = re.findall(r"\d{6,13}", s)
            if matches:
                return matches[-1]
            return ''

        upc_buscar = str(upc).strip()
        upc_digits = extract_digits_candidate(upc_buscar) or re.sub(r"\D", "", upc_buscar)

        for item in factura_items:
            # Generar candidatos de UPC desde varias columnas posibles
            candidatos = []
            candidatos.append(item.get('UPC', ''))
            candidatos.append(item.get('# ORDEN - ITEM', ''))
            candidatos.append(item.get('ORDEN - ITEM', ''))
            candidatos.append(item.get('# ORDEN - ITEM', ''))
            candidatos.append(item.get('ORDEN - ITEM', ''))

            for cand in candidatos:
                cand_digits = extract_digits_candidate(cand) or re.sub(r"\D", "", str(cand or ""))
                if not cand_digits:
                    continue
                # Coincidencia exacta en dígitos
                if cand_digits == upc_digits and upc_digits:
                    return item

            # Fallback: intentar las reglas anteriores de contener ' - '
            try:
                item_upc = str(item.get('UPC','')).strip()
                if (item_upc == upc_buscar or f" - {item_upc}" in upc_buscar or f" - {upc_buscar}" in item_upc):
                    return item
            except Exception:
                pass

        return None

    def center_window(self):
        """Centra la ventana en la pantalla"""
        self.update_idletasks()
        width = 1200
        height = 600
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def crear_header(self):
        header = ctk.CTkFrame(self, fg_color=STYLE["fondo"], height=60)
        titulo = ctk.CTkLabel(header, text="📦 SISTEMA DE ESCANEO", font=FONT_TITLE, text_color=STYLE["header_texto"])
        titulo.place(relx=0.5, rely=0.5, anchor="center")
        # Contador de pendientes visible en la ventana principal
        self.label_pendientes_main = ctk.CTkLabel(
            header,
            text="Codigos por escanear: 0/0",
            font=("Inter", 12, "bold"),
            text_color=STYLE["texto_claro"]
        )
        self.label_pendientes_main.place(relx=0.02, rely=0.5, anchor="w")
        return header

    def actualizar_pendientes_main(self):
        try:
            factura_items = self.factura_data.get("items", [])
            if not factura_items:
                self.label_pendientes_main.configure(
                    text="Total de códigos por lote: 0/0"
                )
                return

            total = 0
            escaneados = 0

            for item in factura_items:
                upc = self.normalizar_upc(item.get("UPC"))
                if not upc:
                    continue

                total += 1
                if self.contador_escaneos.get(upc, 0) > 0:
                    escaneados += 1

            self.label_pendientes_main.configure(
                text=f"Total de códigos por lote: {escaneados}/{total}"
            )

        except Exception as e:
            print(f"❌ Error actualizando pendientes: {e}")
            self.label_pendientes_main.configure(
                text="Total de códigos por lote: 0/0"
            )

    def crear_body(self):
        body = ctk.CTkFrame(self, fg_color=STYLE["fondo"])
        
        body.grid_rowconfigure(0, weight=0)
        body.grid_rowconfigure(1, weight=0)
        body.grid_rowconfigure(2, weight=1)
        
        body.grid_columnconfigure(0, weight=0)
        body.grid_columnconfigure(1, weight=1)
        body.grid_columnconfigure(2, weight=0)
        
        # Barra de escaneo
        scan_bar = ctk.CTkFrame(body, fg_color=STYLE["fondo"], height=45)
        scan_bar.grid(row=0, column=0, columnspan=3, sticky="ew", padx=10, pady=(10, 5))
        scan_bar.grid_propagate(False)
        
        scan_content = ctk.CTkFrame(scan_bar, fg_color=STYLE["fondo"])
        scan_content.pack(fill="x", padx=20, pady=5)
        
        # Campo de entrada UPC
        entry_upc = ctk.CTkEntry(
            scan_content, 
            textvariable=self.codigo_actual,
            placeholder_text="Pase el código de barras aquí",
            width=150,
            height=32,
            font=FONT_TEXT
        )
        entry_upc.pack(side="left", padx=(0, 10))
        entry_upc.bind("<Return>", self.on_scan)
        entry_upc.focus_set()
        
        # Botón Limpiar
        btn_limpiar = ctk.CTkButton(
            scan_content,
            text="Limpiar",
            width=70,
            height=32,
            command=self.clear_search,
            fg_color=STYLE["borde"],
            hover_color=STYLE["button_hover"],
            text_color=STYLE["texto_oscuro"],
            font=("Inter", 10, "bold")
        )
        btn_limpiar.pack(side="left", padx=(0, 20))
        
        # Botones de configuración - AGREGADO BOTÓN FACTURA y LAYOUT
        btn_factura = ctk.CTkButton(
            scan_content,
            text="📋 Subir Factura",
            command=self.cargar_factura_archivo,
            width=70,
            height=32,
            fg_color=STYLE["primario"],
            hover_color=STYLE["hover_primario"],
            text_color=STYLE["header_texto"],
            font=("Inter", 10, "bold")
        )
        btn_factura.pack(side="left", padx=(0, 5))
        
        btn_layout = ctk.CTkButton(
            scan_content,
            text="📄 Subir Layout",
            command=self.cargar_layout_archivo,
            width=70,
            height=32,
            fg_color=STYLE["primario"],
            hover_color=STYLE["hover_primario"],
            text_color=STYLE["header_texto"],
            font=("Inter", 10, "bold")
        )
        btn_layout.pack(side="left", padx=(0, 5))
        
        # Botón unificado para limpiar factura y layout
        btn_limpiar_todo = ctk.CTkButton(
            scan_content,
            text="🗑️ Limpiar Archvios",
            command=self.limpiar_todo,
            width=80,
            height=32,
            fg_color=STYLE["peligro"],
            hover_color="#e74c3c",
            text_color="white",
            font=("Inter", 10, "bold")
        )
        btn_limpiar_todo.pack(side="left", padx=(0, 5))
        
        btn_carpeta = ctk.CTkButton(
            scan_content,
            text="Subir Carpeta de Imágenes",
            command=self.seleccionar_carpeta,
            width=70,
            height=32,
            fg_color=STYLE["secundario"],
            hover_color=STYLE["button_hover"],
            text_color="white",
            font=("Inter", 10, "bold")
        )
        btn_carpeta.pack(side="left", padx=(0, 5))
        
        btn_subir = ctk.CTkButton(
            scan_content,
            text="Subir Imagen",
            command=lambda: subir_imagen(self),
            width=70,
            height=32,
            fg_color=STYLE["primario"],
            hover_color=STYLE["hover_primario"],
            text_color=STYLE["header_texto"],
            font=("Inter", 10, "bold")
        )
        btn_subir.pack(side="left", padx=(0, 5))
        
        btn_excel = ctk.CTkButton(
            scan_content,
            text="Subir Base",
            command=self.cargar_excel,
            width=70,
            height=32,
            fg_color=STYLE["primario"],
            hover_color=STYLE["hover_primario"],
            text_color=STYLE["header_texto"],
            font=("Inter", 10, "bold")
        )
        btn_excel.pack(side="left")
        
        # Estado del sistema
        self.label_estado = ctk.CTkLabel(
            scan_content,
            text="",
            font=("Inter", 10),
            text_color=STYLE["exito"]
        )
        self.label_estado.pack(side="right", padx=(10, 0))
        
        # Control de medición
        self.medicion_bar = ctk.CTkFrame(body, fg_color=STYLE["surface"], corner_radius=6, height=45)
        self.medicion_bar.grid(row=1, column=0, columnspan=3, sticky="ew", padx=10, pady=(0, 5))
        self.medicion_bar.grid_propagate(False)

        titulo_container = ctk.CTkFrame(self.medicion_bar, fg_color=STYLE["surface"])
        titulo_container.pack(side="left", padx=(20, 0), pady=10)

        titulo_medicion = ctk.CTkLabel(
            titulo_container,
            text="CONTENIDO DECLARADO",
            font=("Inter", 14, "bold"),
            text_color=STYLE["texto_oscuro"]
        )
        titulo_medicion.pack()

        valor_container = ctk.CTkFrame(self.medicion_bar, fg_color=STYLE["surface"])
        valor_container.pack(side="right", padx=(0, 20), pady=10)

        self.label_contenido_grande = ctk.CTkLabel(
            valor_container, 
            text="—", 
            font=("Inter", 20, "bold"),
            text_color=STYLE["texto_oscuro"],
            justify="center"
        )
        self.label_contenido_grande.pack()
        
        # Contenido principal
        # Regla Calibrador Vernier con tarjeta de información
        regla_frame = ctk.CTkFrame(body, fg_color=STYLE["surface"], corner_radius=8)
        regla_frame.grid(row=2, column=0, sticky="nsew", padx=(10, 5), pady=(0, 10))
        regla_frame.grid_propagate(False)

        # 🔽 Disminuye ligeramente el tamaño total del contenedor (antes height=80)
        regla_frame.configure(width=220, height=70)   # ← Ajuste principal

        # Configurar grid para regla_frame (2 filas: regla y tarjeta)
        # 🔽 Reducimos la proporción ocupada por la regla
        regla_frame.grid_rowconfigure(0, weight=2)    # ← Antes: weight=3
        regla_frame.grid_rowconfigure(1, weight=1)    # Tarjeta igual
        regla_frame.grid_columnconfigure(0, weight=1)
        
        # Parte superior: Título y regla
        regla_top_frame = ctk.CTkFrame(regla_frame, fg_color=STYLE["surface"])
        regla_top_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        
        titulo_regla = ctk.CTkLabel(
            regla_top_frame, 
            text="📐 CALIBRADOR VERNIER", 
            font=("Inter", 14, "bold"),
            text_color=STYLE["texto_oscuro"]
        )
        titulo_regla.pack(pady=(15, 10))
        
        canvas_container = ctk.CTkFrame(regla_top_frame, fg_color=STYLE["surface"])
        canvas_container.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.canvas_regla = ctk.CTkCanvas(
            canvas_container, 
            bg=STYLE["surface"], 
            highlightthickness=0,
            relief="flat"
        )
        self.canvas_regla.pack(fill="both", expand=True)
        
        self.label_estado_regla = ctk.CTkLabel(
            regla_top_frame, 
            text="Esperando medición...", 
            font=("Inter", 10),
            text_color=STYLE["texto_claro"],
            justify="center"
        )
        self.label_estado_regla.pack(side="bottom", pady=(0, 10))
        
        # Parte inferior: Tarjeta de información
        self.tarjeta_info_frame = ctk.CTkFrame(regla_frame, fg_color=STYLE["surface"], corner_radius=6, height=120)
        self.tarjeta_info_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.tarjeta_info_frame.grid_propagate(False)
        
        # Título de la tarjeta
        titulo_tarjeta = ctk.CTkLabel(
            self.tarjeta_info_frame,
            text="📋 INFORMACIÓN FACTURA",
            font=("Inter", 12, "bold"),
            text_color=STYLE["texto_oscuro"]
        )
        titulo_tarjeta.pack(pady=(8, 5))

        # Botón para editar la facturación (ubicado en la tarjeta de información)
        try:
            btn_editar_facturacion = ctk.CTkButton(
                self.tarjeta_info_frame,
                text="✏️ Editar Datos",
                command=self.abrir_editor_facturacion,
                width=160,
                height=30,
                fg_color=STYLE["secundario"],
                hover_color=STYLE.get("button_hover", "#4b4b4b"),
                text_color="white",
                font=("Inter", 10, "bold")
            )
            btn_editar_facturacion.pack(anchor="ne", padx=10, pady=(0, 6))
        except Exception:
            # En caso de que customtkinter no soporte algún argumento, evitar romper la UI
            pass
        
        # Contenido de la tarjeta
        contenido_tarjeta = ctk.CTkFrame(self.tarjeta_info_frame, fg_color=STYLE["surface"])
        contenido_tarjeta.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Etiquetas para la información
        self.label_cantidad_facturada = ctk.CTkLabel(
            contenido_tarjeta,
            text="Cantidad Facturada: N/A",
            font=("Inter", 10),
            text_color=STYLE["texto_claro"],
            justify="left",
            anchor="w"
        )
        self.label_cantidad_facturada.pack(anchor="w", pady=(0, 3))
        
        # Mostrar la FACTURA separada del pedimento para mejor legibilidad
        self.label_factura = ctk.CTkLabel(
            contenido_tarjeta,
            text="Factura: N/A",
            font=("Inter", 10),
            text_color=STYLE["texto_claro"],
            justify="left",
            anchor="w"
        )
        self.label_factura.pack(anchor="w", pady=(0, 2))

        self.label_pedimento = ctk.CTkLabel(
            contenido_tarjeta,
            text="Pedimento: N/A",
            font=("Inter", 10),
            text_color=STYLE["texto_claro"],
            justify="left",
            anchor="w"
        )
        self.label_pedimento.pack(anchor="w", pady=(0, 3))
        
        self.label_archivo_activo = ctk.CTkLabel(
            contenido_tarjeta,
            text="Archivo Activo: Ninguno",
            font=("Inter", 10),
            text_color=STYLE["texto_claro"],
            justify="left",
            anchor="w"
        )
        self.label_archivo_activo.pack(anchor="w", pady=(0, 3))
        
        # Imagen
        img_frame = ctk.CTkFrame(body, fg_color=STYLE["surface"], corner_radius=8)
        img_frame.grid(row=2, column=1, sticky="nsew", padx=5, pady=(0, 10))
        
        self.label_imagen = ctk.CTkLabel(
            img_frame, 
            text="🖼️ La imagen aparecerá aquí", 
            fg_color=STYLE["fondo"], 
            corner_radius=6,
            font=FONT_LABEL, 
            text_color=STYLE["texto_oscuro"]
        )
        self.label_imagen.pack(padx=10, pady=10, fill="both", expand=True)
        
        # Información crítica y comercial
        info_column = ctk.CTkFrame(body, fg_color=STYLE["fondo"])
        info_column.grid(row=2, column=2, sticky="nsew", padx=(5, 10), pady=(0, 10))
        info_column.grid_propagate(False)
        info_column.configure(width=380)
        
        info_column.grid_rowconfigure(0, weight=1)
        info_column.grid_rowconfigure(1, weight=1)
        info_column.grid_columnconfigure(0, weight=1)
        
        # Información crítica
        critica_frame = ctk.CTkFrame(info_column, fg_color=STYLE["surface"], corner_radius=8)
        critica_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=(0, 5))
        
        titulo_critica = ctk.CTkLabel(
            critica_frame, 
            text="📑 INFORMACIÓN CRÍTICA NOM", 
            font=("Inter", 14, "bold"), 
            text_color=STYLE["texto_oscuro"]
        )
        titulo_critica.pack(pady=(12, 8), padx=15, anchor="w")
        
        critica_scroll = ctk.CTkScrollableFrame(
            critica_frame, 
            fg_color=STYLE["surface"], 
            corner_radius=6,
            scrollbar_button_color=STYLE["secundario"],
            scrollbar_button_hover_color=STYLE["button_hover"]
        )
        critica_scroll.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        self.label_puntos = ctk.CTkLabel(
            critica_scroll, 
            text="\nSeleccione un producto...", 
            font=FONT_TEXT, 
            justify="left", 
            anchor="nw", 
            text_color=STYLE["texto_claro"],
            wraplength=380
        )
        self.label_puntos.pack(padx=12, pady=8, fill="both", expand=True)

        
        # Información comercial
        comercial_frame = ctk.CTkFrame(info_column, fg_color=STYLE["surface"], corner_radius=8)
        comercial_frame.grid(row=1, column=0, sticky="nsew", padx=0, pady=(5, 0))
        
        comercial_header = ctk.CTkFrame(comercial_frame, fg_color=STYLE["surface"])
        comercial_header.pack(fill="x", padx=15, pady=(12, 8))
        
        comercial_title = ctk.CTkLabel(
            comercial_header, 
            text="📋 INFORMACIÓN COMERCIAL", 
            font=("Inter", 14, "bold"), 
            text_color=STYLE["texto_oscuro"]
        )
        comercial_title.pack(side="left", anchor="w")
        
        self.btn_editar_info = ctk.CTkButton(
            comercial_header,
            text="✏️ Editar",
            command=self.abrir_editor_informacion,
            width=80,
            height=28,
            fg_color=STYLE["secundario"],
            hover_color=STYLE["hover_boton"],
            text_color="white",
            font=("Inter", 10, "bold")
        )
        self.btn_editar_info.pack(side="right", padx=(5, 0))
        self.btn_editar_info.configure(state="disabled")
        
        comercial_scroll = ctk.CTkScrollableFrame(
            comercial_frame, 
            fg_color=STYLE["surface"], 
            corner_radius=6,
            scrollbar_button_color=STYLE["secundario"],
            scrollbar_button_hover_color=STYLE["button_hover"]
        )
        comercial_scroll.pack(fill="both", expand=True, padx=10, pady=(0, 12))
        
        self.label_info = ctk.CTkLabel(
            comercial_scroll, 
            text="La información aparecerá aquí después del escaneo", 
            font=FONT_TEXT, 
            justify="left", 
            anchor="nw", 
            text_color=STYLE["texto_claro"],
            wraplength=380
        )
        self.label_info.pack(padx=12, pady=8, fill="both", expand=True)
        
        # Inicializar regla Calibrador Vernier
        dibujar_regla_bernier(self.canvas_regla, "")
        self.after(180, lambda: dibujar_regla_bernier(self.canvas_regla, ""))
        
        # Inicializar tarjeta de información
        self.actualizar_tarjeta_informacion()
        
        return body

    def actualizar_tarjeta_informacion(self, item_factura=None):
        """Actualiza la tarjeta INFORMACIÓN FACTURA mostrando la factura correcta por UPC"""
        try:
            cantidad_facturada = "N/A"
            factura = "N/A"
            pedimento = "N/A"
            archivo_activo = "Ninguno"

            # ─────────────────────────────────────────────
            # 1) OBTENER DATOS DE FACTURA SEGÚN UPC
            # ─────────────────────────────────────────────
            if item_factura and self.producto_actual:
                cantidad_facturada = (
                    item_factura.get('CANT. FACT.')
                    or item_factura.get('CANTIDAD_FACTURA')
                    or "N/A"
                )
                factura = item_factura.get('FACTURA', 'N/A')

            elif self.producto_actual:
                upc_actual = str(self.producto_actual.get("UPC", "")).strip()
                item_factura = self.buscar_en_factura(upc_actual)

                if item_factura:
                    cantidad_facturada = (
                        item_factura.get('CANT. FACT.')
                        or item_factura.get('CANTIDAD_FACTURA')
                        or "N/A"
                    )
                    factura = item_factura.get('FACTURA', 'N/A')

            # ─────────────────────────────────────────────
            # 2) PEDIMENTO = NOMBRE DEL ARCHIVO DE FACTURA
            # ─────────────────────────────────────────────
            if self.factura_data:
                nombre_archivo = self.factura_data.get("nombre_archivo", "")
                if nombre_archivo:
                    pedimento = os.path.splitext(nombre_archivo)[0]

            # ─────────────────────────────────────────────
            # 3) DETERMINAR ARCHIVO ACTIVO
            # ─────────────────────────────────────────────
            factura_items = self.factura_data.get('items', [])
            layout_items = self.layout_data

            if factura_items and layout_items:
                archivo_activo = "Factura y Layout"
            elif factura_items:
                archivo_activo = "Factura"
            elif layout_items:
                archivo_activo = "Layout"

            # ─────────────────────────────────────────────
            # 4) ACTUALIZAR UI (SOLO INFORMACIÓN FACTURA)
            # ─────────────────────────────────────────────
            self.label_cantidad_facturada.configure(
                text=f"Cantidad Facturada: {cantidad_facturada}"
            )

            # Mostrar FACTURA y PEDIMENTO en líneas separadas (Factura arriba, Pedimento abajo)
            if factura and factura != 'N/A':
                try:
                    self.label_factura.configure(text=f"Factura: {factura}")
                except Exception:
                    # fallback if label missing
                    self.label_factura = ctk.CTkLabel(
                        self.tarjeta_info_frame, text=f"Factura: {factura}", font=("Inter",10), text_color=STYLE["texto_claro"]
                    )
                    self.label_factura.pack(anchor="w", pady=(0,2))
            else:
                try:
                    self.label_factura.configure(text="Factura: N/A")
                except Exception:
                    pass

            if pedimento:
                self.label_pedimento.configure(text=f"Pedimento: {pedimento}")
            else:
                self.label_pedimento.configure(text="Pedimento: N/A")

            # Mostrar solo el estado de archivo activo (Factura/Layout)
            self.label_archivo_activo.configure(
                text=f"Activo: {archivo_activo}"
            )

            # ⚠️ Pedimento (archivo) se muestra SOLO aquí si tienes el label
            if hasattr(self, "label_pedimento_archivo"):
                self.label_pedimento_archivo.configure(
                    text=f"Pedimento: {pedimento}"
                )

        except Exception as e:
            print(f"❌ Error actualizando INFORMACIÓN FACTURA: {e}")

    def actualizar_footer(self):
        """Actualiza el texto del footer con la información actual"""
        footer_text = f"Sistema De V&C para ULTA AXO"
        if self.excel_path:
            excel_name = os.path.basename(self.excel_path)
            footer_text += f" | Excel: {excel_name}"
        
        # Mostrar si hay factura cargada
        factura_items = self.factura_data.get('items', [])
        if factura_items:
            nombre_archivo = self.factura_data.get('nombre_archivo', 'Sin nombre')
            nombre_sin_ext = os.path.splitext(nombre_archivo)[0] if nombre_archivo else 'Sin nombre'
            footer_text += f" | Factura: {nombre_sin_ext} ({len(factura_items)} productos)"
        
        # Mostrar si hay layout cargado
        if self.layout_data and 'total_etiquetas' in self.layout_data:
            nombre_layout = self.layout_data.get('nombre_archivo', 'Sin nombre')
            nombre_sin_ext_layout = os.path.splitext(nombre_layout)[0] if nombre_layout else 'Sin nombre'
            footer_text += f" | Layout: {nombre_sin_ext_layout} ({self.layout_data['total_etiquetas']} etiquetas)"
        
        self.footer_label.configure(text=footer_text)

    def crear_footer(self):
        footer = ctk.CTkFrame(self, fg_color=STYLE["fondo"], height=20)
        
        self.footer_label = ctk.CTkLabel(
            footer, 
            text="",
            text_color=STYLE["header_texto"],
            font=FONT_TEXT
        )
        self.footer_label.place(relx=0.5, rely=0.5, anchor="center")
        
        # Llamar a actualizar_footer para establecer el texto inicial
        self.actualizar_footer()
        
        return footer

    def on_resize(self, event):
        """Redibuja la regla cuando cambia el tamaño de la ventana"""
        try:
            if self.producto_actual:
                tamano_declaracion = obtener_campo(self.producto_actual, "TAMAÑO DE LA DECLARACION DE CONTENIDO") or ""
                tamano_declaracion = str(tamano_declaracion).strip()
                dibujar_regla_bernier(self.canvas_regla, tamano_declaracion)
        except:
            pass

    def seleccionar_carpeta(self):
        """Permite al usuario seleccionar la carpeta de imágenes"""
        carpeta = seleccionar_carpeta_imagenes()
        if carpeta:
            self.config = cargar_configuracion()
            messagebox.showinfo("Carpeta seleccionada", f"Carpeta de imágenes configurada:\n{carpeta}")
            
            image_dir = self.config.get("image_dir", "No configurada")
            if image_dir and len(image_dir) > 40:
                image_dir_display = "..." + image_dir[-37:]
            else:
                image_dir_display = image_dir
                
            self.footer_label.configure(
                text=f"Sistema De V&C para ULTA AXO | Carpeta: {image_dir_display}"
            )
            
            if self.producto_actual:
                categoria = str(self.producto_actual.get("CATEGORIA"))
                mostrar_imagen(self.label_imagen, categoria)

    def abrir_editor_informacion(self):
        """Abre el editor de información comercial"""
        if self.producto_actual:
            EditorInformacionComercial(self, self.producto_actual, self.guardar_cambios_producto)
        else:
            messagebox.showwarning("Sin producto", "No hay ningún producto escaneado para editar.")

    def abrir_editor_facturacion(self):
        """Abre el editor de información facturada o layout"""
        # Importar aquí para evitar dependencia circular
        from EditorFacturacion import EditorFacturacion
        
        # Determinar qué datos pasar
        factura_items = self.factura_data.get('items', [])
        layout_items = self.layout_data.get('datos', [])
        
        if factura_items:
            # Si hay factura, abrir editor en modo factura
            editor = EditorFacturacion(
                self, 
                factura_data=self.factura_data, 
                layout_data=None,
                contador_escaneos=self.contador_escaneos
            )
            # Guardar referencia para poder notificar actualizaciones de escaneo
            try:
                self.editor_facturacion = editor
                # Limpiar la referencia cuando la ventana se destruya
                editor.bind('<Destroy>', lambda e: setattr(self, 'editor_facturacion', None))
            except Exception:
                pass
        elif layout_items:
            # Si no hay factura pero sí layout, abrir editor en modo layout
            editor = EditorFacturacion(
                self, 
                factura_data=None,
                layout_data=self.layout_data,
                contador_escaneos=self.contador_escaneos
            )
            try:
                self.editor_facturacion = editor
                editor.bind('<Destroy>', lambda e: setattr(self, 'editor_facturacion', None))
            except Exception:
                pass
        else:
            # Si no hay nada cargado
            messagebox.showwarning(
                "Sin datos", 
                "No hay factura ni layout cargados para editar.\n\n"
                "Por favor, suba primero una factura o un layout."
            )
            return
        
        editor.focus()

    def clear_search(self):
        """Limpia la búsqueda actual"""
        try:
            self.codigo_actual.set("")
            self.producto_actual = None
            
            try:
                self.label_imagen.configure(image="")
                self.label_imagen.image = None
                self.label_imagen.configure(text="🖼️ La imagen aparecerá aquí")
            except:
                pass
            
            try:
                self.label_info.configure(text="La información aparecerá aquí después del escaneo")
                self.label_puntos.configure(text="\nSeleccione un producto...")
                self.label_contenido_grande.configure(text="—")
            except:
                pass
            
            dibujar_regla_bernier(self.canvas_regla, "")
            
            try:
                self.btn_editar_info.configure(state="disabled")
            except:
                pass
                
            self.label_estado.configure(text="")
            
            # Actualizar tarjeta de información
            self.actualizar_tarjeta_informacion()
            
        except Exception as e:
            print(f"Error limpiando búsqueda: {e}")

    def guardar_cambios_producto(self, producto_actualizado):
        """Guarda los cambios del producto en JSON y Excel"""
        try:
            # Actualizar el producto actual
            self.producto_actual = producto_actualizado.copy()
            
            # Actualizar la lista de productos
            upc = str(producto_actualizado.get("UPC", "")).strip()
            encontrado = False
            
            for i, producto in enumerate(self.productos):
                if isinstance(producto, dict):
                    producto_upc = str(producto.get("UPC", "")).strip()
                    if producto_upc == upc:
                        # Fusionar cambios manteniendo todos los campos
                        for key, value in producto_actualizado.items():
                            self.productos[i][key] = value
                        encontrado = True
                        break
            
            if not encontrado:
                self.productos.append(producto_actualizado.copy())
            
            # Guardar en JSON
            try:
                with open(JSON_FILE, 'w', encoding='utf-8') as f:
                    json.dump(self.productos, f, ensure_ascii=False, indent=4)
                print(f"✅ JSON guardado: {upc}")
            except Exception as e:
                print(f"❌ Error guardando JSON: {e}")
                messagebox.showerror("Error", f"No se pudo guardar el JSON:\n{e}")
                return
            
            # Actualizar Excel si existe
            if self.excel_path and os.path.exists(self.excel_path):
                try:
                    self.actualizar_excel(producto_actualizado)
                except Exception as e:
                    print(f"❌ Error actualizando Excel: {e}")
                    # Aún mostramos éxito porque el JSON se guardó
                    messagebox.showinfo("Información", 
                                      "✅ Cambios guardados en JSON. " +
                                      f"Error al actualizar Excel: {str(e)[:100]}")
            else:
                messagebox.showinfo("Información", 
                                  "✅ Cambios guardados en JSON. " +
                                  "No se encontró archivo Excel para actualizar.")
            
            # Actualizar vista
            self.actualizar_vista_producto(self.producto_actual)
            
            # Actualizar imagen si es necesario
            categoria = str(self.producto_actual.get("CATEGORIA", ""))
            if categoria:
                mostrar_imagen(self.label_imagen, categoria)
            
            # Verificar si el producto está en la factura cargada
            if self.factura_data.get('items'):
                item_factura = self.buscar_en_factura(upc)
                if item_factura:
                    self.actualizar_tarjeta_informacion(item_factura)
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron guardar los cambios:\n{e}")

    def actualizar_excel(self, producto_actualizado):
        """Actualiza el archivo Excel con los cambios"""
        try:
            # Leer Excel
            df = pd.read_excel(self.excel_path, dtype=str, keep_default_na=False)
            
            # Buscar columna UPC
            upc_col = None
            for col in df.columns:
                col_norm = str(col).upper().replace(" ", "_").replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")
                if col_norm in ["UPC", "CODIGOUPC", "CODIGO_UPC"]:
                    upc_col = col
                    break
            
            if not upc_col and len(df.columns) > 0:
                upc_col = df.columns[0]
            
            # Buscar fila con el UPC
            upc = str(producto_actualizado.get("UPC", "")).strip()
            mask = df[upc_col].astype(str).str.strip() == upc
            
            if mask.any():
                idx = df[mask].index[0]
                
                # Actualizar cada campo
                for campo, valor in producto_actualizado.items():
                    if campo == "UPC":
                        continue
                    
                    # Buscar columna correspondiente
                    col_encontrada = None
                    for col in df.columns:
                        col_norm = str(col).upper().replace(" ", "_").replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")
                        campo_norm = str(campo).upper().replace(" ", "_").replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")
                        
                        if col_norm == campo_norm:
                            col_encontrada = col
                            break
                    
                    if col_encontrada:
                        valor_str = str(valor).strip() if valor is not None else ""
                        df.at[idx, col_encontrada] = valor_str
                    else:
                        # Si no existe la columna, agregarla
                        df[campo] = ""
                        df.at[idx, campo] = str(valor).strip() if valor is not None else ""
                
                # Guardar Excel
                df.to_excel(self.excel_path, index=False)
                print(f"✅ Excel actualizado: {upc}")
                
            else:
                print(f"⚠️ UPC {upc} no encontrado en Excel")
                
        except Exception as e:
            print(f"❌ Error actualizando Excel: {e}")
            raise

    def actualizar_vista_producto(self, producto):
        """Actualiza la vista con la información del producto"""
        try:
            # Excluir campos que ya se muestran en otras partes
            campos_excluidos = ["NORMA"]
            campos_prioritarios = [
                "CATEGORIA", "UPC", "DENOMINACION", "DENOMINACION_AXO",
                "MARCA", "OBSERVACIONES_REVISION", "CONTENIDO", "TAMAÑO_DE_LA_DECLARACION_DE_CONTENIDO",
                "PAIS_ORIGEN", "IMPORTADOR", "INSTRUCCIONES_DE_USO",
                "LEYENDAS_PRECAUTORIAS", "OBSERVACIONES", "INGREDIENTES_Y_LOTE",
                "MEDIDAS", "TIPO_DE_ETIQUETA"
            ]
            
            # Formatear información
            info_text = ""
            
            # ⚠️ VERIFICAR OBSERVACIONES_REVISION PRIMERO - mostrar con alerta si no es "OK"
            observaciones_revision = str(producto.get("OBSERVACIONES_REVISION", "OK")).strip()
            if observaciones_revision and observaciones_revision.lower() != "nan" and observaciones_revision.upper() != "OK":
                # Mostrar alerta prominente para observaciones que no son "OK"
                info_text += f"⚠️  OBSERVACIONES DE REVISIÓN:\n"
                info_text += f"🔴 {observaciones_revision}\n"
                info_text += f"\n{'─'*40}\n\n"
                # También mostrar en el estado con color de advertencia
                self.label_estado.configure(
                    text=f"⚠️  {observaciones_revision[:70]}",
                    text_color=STYLE["advertencia"]
                )
            
            for key in campos_prioritarios:
                if key not in campos_excluidos:
                    # Saltar OBSERVACIONES_REVISION pues ya se mostró arriba en alerta
                    if key == "OBSERVACIONES_REVISION":
                        continue
                    
                    valor = producto.get(key, "")
                    if valor and str(valor).strip() and str(valor).strip().lower() != "nan":
                        nombre_campo = key.replace("_", " ").title()
                        if nombre_campo == "Upc":
                            nombre_campo = "Código UPC"
                        elif nombre_campo == "Denominacion Axo":
                            nombre_campo = "Denominación AXO"
                        elif nombre_campo == "Tamaño De La Declaracion De Contenido":
                            nombre_campo = "Declaración de Tamaño (NOM)"
                        
                        info_text += f"📌 {nombre_campo}:\n{valor}\n\n"
            
            # Resto de campos
            for key in producto.keys():
                if key not in campos_prioritarios and key not in campos_excluidos:
                    valor = producto.get(key, "")
                    if valor and str(valor).strip() and str(valor).strip().lower() != "nan":
                        nombre_campo = key.replace("_", " ").title()
                        info_text += f"• {nombre_campo}: {valor}\n\n"
            
            # Agregar información de factura si existe para este producto
            factura_items = self.factura_data.get('items', [])
            if factura_items:
                upc = str(producto.get("UPC", "")).strip()
                item_factura = self.buscar_en_factura(upc)
                if item_factura:
                    cantidad = item_factura.get('CANTIDAD_FACTURA', 'N/A')
                    # info_text += f"\n📦 Cantidad Factura: {cantidad}\n"
                    
                    # Mostrar nombre del archivo de factura
                    nombre_archivo_factura = self.factura_data.get('nombre_archivo', '')
                    if nombre_archivo_factura:
                        nombre_sin_ext_factura = os.path.splitext(nombre_archivo_factura)[0]
                        # Nota: el nombre del archivo (pedimento) se mostrará en la tarjeta INFORMACIÓN FACTURA
                        # para evitar duplicados en la sección de Información Comercial.
                        pass
            
            # Agregar información de layout si está cargado
            if self.layout_data and 'total_etiquetas' in self.layout_data:
                info_text += f"\n📄 Información de Layout:\n"
                info_text += f"   • Total de etiquetas a verificar: {self.layout_data['total_etiquetas']}\n"
                info_text += f"   • Columna utilizada: {self.layout_data.get('columna_utilizada', 'N/A')}\n"
                
                # Mostrar nombre del archivo según el formato especificado
                nombre_archivo = self.layout_data.get('nombre_archivo', '')
                nombre_sin_ext = self.layout_data.get('nombre_sin_ext', '')
                
                if nombre_archivo:
                    # Para nombres como "25 24 3622 5009437.xlsx"
                    info_text += f"   • Archivo Layout: {nombre_sin_ext}\n"
                else:
                    info_text += f"   • Archivo Layout: No disponible\n"
            
            
            if not info_text.strip():
                info_text = "La información aparecerá aquí después del escaneo"
            
            self.label_info.configure(text=info_text.strip())
            
            # Mostrar contenido y tamaño
            contenido_real = str(producto.get("CONTENIDO", "")).strip()
            tamano_declaracion = str(producto.get("TAMAÑO_DE_LA_DECLARACION_DE_CONTENIDO", "")).strip()
            
            contenido_display = contenido_real if contenido_real else "—"
            self.label_contenido_grande.configure(text=contenido_display)
            
            # Actualizar regla Bernier
            if tamano_declaracion and tamano_declaracion.lower() != "nan":
                dibujar_regla_bernier(self.canvas_regla, tamano_declaracion)
                self.label_estado_regla.configure(
                    text=f"Medición: {tamano_declaracion}",
                    text_color=STYLE["exito"]
                )
            else:
                dibujar_regla_bernier(self.canvas_regla, "")
                self.label_estado_regla.configure(
                    text="Esperando medición...",
                    text_color=STYLE["texto_claro"]
                )
            
            # Puntos normativos
            norma = producto.get("NORMA", "").upper()
            puntos = obtener_puntos_normativos(norma)
            puntos_text = f"📋 Norma: {norma}\n\n"
            
            for i, punto in enumerate(puntos, 1):
                punto_text = str(punto)
                if len(punto_text) > 120:
                    punto_text = punto_text[:117] + "..."
                puntos_text += f"{i}. {punto_text}\n\n"
            
            self.label_puntos.configure(text=puntos_text.strip())
            
            # Habilitar edición
            self.btn_editar_info.configure(state="normal")
            
            # Actualizar tarjeta de información
            self.actualizar_tarjeta_informacion()
            
        except Exception as e:
            print(f"Error actualizando vista: {e}")

    def on_scan(self, event=None):
        upc = self.codigo_actual.get().strip()
        if not upc:
            return

        try:
            producto = buscar_producto_por_upc(upc, self.productos)
            if not producto:
                self.clear_search()
                self.label_estado.configure(
                    text="❌ Producto no encontrado",
                    text_color=STYLE["peligro"]
                )
                return

            # Mostrar imagen
            categoria = str(producto.get("CATEGORIA", ""))
            if categoria and categoria.lower() != "nan":
                mostrar_imagen(self.label_imagen, categoria)
            else:
                self.label_imagen.configure(image="", text="🖼️ Sin imagen disponible")
                self.label_imagen.image = None

            # Actualizar producto actual
            self.producto_actual = producto
            self.btn_editar_info.configure(state="normal")

            # Actualizar vista
            self.actualizar_vista_producto(producto)

            # Actualizar tarjeta de información de factura (si existe)
            try:
                self.actualizar_tarjeta_informacion()
            except Exception:
                pass

            # Incrementar contador de escaneos usando el UPC normalizado
            upc_norm = self.normalizar_upc(upc)
            if upc_norm:
                self.contador_escaneos[upc_norm] = self.contador_escaneos.get(upc_norm, 0) + 1

            # Mostrar estado usando la clave normalizada cuando esté disponible
            display_key = upc_norm or upc
            display_count = self.contador_escaneos.get(display_key, 0)
            self.label_estado.configure(
                text=f"✅ Producto encontrado: {upc} (Escaneos: {display_count})",
                text_color=STYLE["exito"]
            )

            # Actualizar contador de pendientes en la ventana principal
            try:
                self.actualizar_pendientes_main()
            except Exception:
                pass

            # Si el editor de facturación está abierto, notificarle el escaneo (usar UPC normalizado)
            try:
                editor = getattr(self, 'editor_facturacion', None)
                if editor is not None:
                    try:
                        notify_key = upc_norm or upc
                        if hasattr(editor, 'actualizar_escaneo'):
                            editor.actualizar_escaneo(notify_key)
                        # Siempre intentar recalcular el contador visible en el editor
                        if hasattr(editor, 'actualizar_contador_pendientes'):
                            editor.actualizar_contador_pendientes()
                    except Exception:
                        pass
            except Exception:
                pass

        except Exception as e:
            messagebox.showerror("Error", f"Error al procesar el escaneo:\n{e}")
            print(f"Error en on_scan: {e}")

    def cargar_excel(self):
        """Carga un archivo Excel y actualiza el sistema"""
        try:
            path_excel = filedialog.askopenfilename(
                title="Seleccionar archivo Excel",
                filetypes=[("Excel files", "*.xlsx *.xls")]
            )
            
            if not path_excel:
                return
            
            # Cargar base desde Excel
            resultado = cargar_base_excel(path_excel)
            if resultado:
                # Actualizar ruta del Excel
                self.excel_path = path_excel
                
                # Actualizar config.json
                try:
                    config_data = {}
                    if os.path.exists(CONFIG_FILE):
                        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                            config_data = json.load(f)
                    
                    config_data['excel_path'] = path_excel
                    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                        json.dump(config_data, f, indent=4, ensure_ascii=False)
                except Exception as e:
                    print(f"Error actualizando config.json: {e}")
                
                # Recargar productos
                self.cargar_productos()
                
                # Actualizar footer
                self.actualizar_footer()
                
                # Mostrar mensaje
                self.label_estado.configure(
                    text="✅ Base actualizada con éxito",
                    text_color=STYLE["exito"]
                )
                self.after(3000, lambda: self.label_estado.configure(text=""))
                
                # Actualizar vista si hay producto actual
                if self.codigo_actual.get().strip():
                    self.on_scan()
                    
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el Excel:\n{e}")

# ---------------- EJECUCIÓN ---------------- #
if __name__ == "__main__":
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")
    app = EscanerApp()
    app.mainloop()

  