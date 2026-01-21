import os, sys
import customtkinter as ctk
import pandas as pd
import json
import re
from tkinter import filedialog, messagebox
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, numbers
from tkinter import simpledialog
from datetime import datetime
import numpy as np


# -----------------------------
# RUTAS Y ARCHIVOS DE PERSISTENCIA
# -----------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RESOURCES_DIR = os.path.join(BASE_DIR, "resources")
USER_DATA_DIR = os.path.join(os.path.expanduser("~"), "ULTA_APP")
FOLIOS_FILE = os.path.join(USER_DATA_DIR, "folios.json")
os.makedirs(USER_DATA_DIR, exist_ok=True)



def resource_path(relative_path):
    """ Devuelve la ruta absoluta al recurso, compatible con PyInstaller """
    try:
        base_path = sys._MEIPASS  # Carpeta temporal en exe
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# -----------------------------
# CONFIGURACIÓN DE ESTILO
# -----------------------------
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("dark-blue")

COLORES = {
    "amarillo": "#ecd925",
    "negro": "#282828",
    "gris_oscuro": "#4d4d4d",
    "gris_claro": "#d8d8d8",
    "blanco": "#FFFFFF"
}

# Configurar tipografía INTER
FUENTE_PRINCIPAL = "Inter"
FUENTE_SECUNDARIA = "Inter"

# -----------------------------
# FUNCIONES MEJORADAS PARA COMPATIBILIDAD .EXE
# -----------------------------
def leer_json(ruta):
    """Lee un archivo JSON con manejo de errores para .exe"""
    try:
        # Primero intenta con la ruta normal
        if os.path.exists(ruta):
            return pd.read_json(ruta, orient="records")
        
        # Si no existe, intenta con resource_path
        ruta_alternativa = resource_path(ruta)
        if os.path.exists(ruta_alternativa):
            return pd.read_json(ruta_alternativa, orient="records")
        
        # Si no existe en ninguna ruta, crear archivo vacío
        print(f"⚠️ Archivo {ruta} no encontrado, creando DataFrame vacío")
        return pd.DataFrame()
        
    except Exception as e:
        print(f"❌ Error leyendo {ruta}: {e}")
        return pd.DataFrame()

def cargar_paises():
    """Carga el archivo de países con manejo de errores"""
    try:
        # Intentar diferentes rutas
        rutas_posibles = [
            os.path.join("resources", "Paises.json"),
            resource_path(os.path.join("resources", "Paises.json")),
            "Paises.json",
            resource_path("Paises.json")
        ]
        
        for ruta in rutas_posibles:
            if os.path.exists(ruta):
                with open(ruta, "r", encoding="utf-8") as f:
                    paises_data = json.load(f)
                if isinstance(paises_data, list) and len(paises_data) > 0:
                    return {k.upper(): v for k, v in paises_data[0].items()}
                else:
                    return {k.upper(): v for k, v in paises_data.items()}
        
        # Si no se encuentra, retornar diccionario vacío
        print("⚠️ No se pudo cargar Paises.json, usando diccionario vacío")
        return {}
        
    except Exception as e:
        print(f"❌ Error cargando países: {e}")
        return {}

def cargar_machote():
    """Carga el archivo machote.json con manejo de errores"""
    try:
        # Intentar diferentes rutas
        rutas_posibles = [
            os.path.join("theme", "machote.json"),
            resource_path(os.path.join("theme", "machote.json")),
            "machote.json",
            resource_path("machote.json"),
            os.path.join("resources", "machote.json"),
            resource_path(os.path.join("resources", "machote.json"))
        ]
        
        for ruta in rutas_posibles:
            if os.path.exists(ruta):
                with open(ruta, "r", encoding="utf-8") as f:
                    return json.load(f)
        
        # Si no se encuentra, retornar diccionario vacío
        print("⚠️ No se pudo cargar machote.json, usando valores por defecto")
        return {}
        
    except Exception as e:
        print(f"❌ Error cargando machote: {e}")
        return {}

# -----------------------------
# CLASE VENTANA PRINCIPAL (SIN MODIFICAR LÓGICA)
# -----------------------------
class VentanaULTA(ctk.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        
        self.folios_disponibles = []
        self.cargar_folios_json()
        self.crear_interfaz()

        # Esto asegura que el contador muestre folios al iniciar
        self.actualizar_contador_folios()

        self.title("Generador Tablas de Relación ULTA")
        self.geometry("800x520")
        self.minsize(800, 520)
        self.configure(fg_color=COLORES["blanco"])
        
        # Centrar ventana en la pantalla
        self.center_window()

        # Archivos cargados
        self.layout = None
        self.emmanuel = None

        # Carpeta destino - crear si no existe
        self.carpeta_resources = "resources"
        os.makedirs(self.carpeta_resources, exist_ok=True)

    def center_window(self):
        """Centra la ventana en la pantalla"""
        self.update_idletasks()
        width = 800
        height = 520
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')

    def crear_interfaz(self):
        """Crea la interfaz gráfica compacta y visible con botones funcionales y tamaño uniforme"""
        BTN_WIDTH = 160
        BTN_HEIGHT = 30
        BTN_FONT = (FUENTE_SECUNDARIA, 11)

        # ------------------ CONTENEDOR PRINCIPAL ------------------
        main_container = ctk.CTkFrame(self, fg_color=COLORES["blanco"], corner_radius=0)
        main_container.pack(fill="both", expand=True, padx=10, pady=10)

        # ------------------ HEADER ------------------
        header_frame = ctk.CTkFrame(main_container, fg_color=COLORES["blanco"], corner_radius=0)
        header_frame.pack(fill="x", pady=(0, 10))

        titulo_frame = ctk.CTkFrame(header_frame, fg_color=COLORES["amarillo"], corner_radius=20)
        titulo_frame.pack(fill="x", padx=20)

        ctk.CTkLabel(titulo_frame, text="📊 TABLAS DE RELACIÓN ULTA",
                    font=(FUENTE_PRINCIPAL, 20, "bold"),
                    text_color=COLORES["negro"]).pack(pady=10)

        # ------------------ DATOS MANUALES ------------------
        datos_frame = ctk.CTkFrame(main_container, fg_color=COLORES["gris_claro"], corner_radius=15, border_width=1, border_color=COLORES["gris_oscuro"])
        datos_frame.pack(fill="x", pady=5, padx=5)

        ctk.CTkLabel(datos_frame, text="DATOS MANUALES",
                    font=(FUENTE_PRINCIPAL, 14, "bold"),
                    text_color=COLORES["negro"]).grid(row=0, column=0, columnspan=6, pady=5)

        form_frame = ctk.CTkFrame(datos_frame, fg_color=COLORES["gris_claro"])
        form_frame.grid(row=1, column=0, columnspan=6, pady=3, padx=3, sticky="ew")
        form_frame.columnconfigure([0,1,2,3,4,5], weight=1)

        def crear_campo(nombre, fila, columna):
            ctk.CTkLabel(form_frame, text=nombre,
                        font=(FUENTE_PRINCIPAL, 11, "bold"),
                        text_color=COLORES["negro"],
                        anchor="w").grid(row=fila, column=columna*2, padx=3, pady=3, sticky="e")
            entry = ctk.CTkEntry(form_frame, width=120, fg_color=COLORES["blanco"], text_color=COLORES["negro"])
            entry.grid(row=fila, column=columna*2+1, padx=3, pady=3, sticky="w")
            return entry

        self.entry_solicitud = crear_campo("📄 SOLICITUD:", 0, 0)
        self.entry_pedimento = crear_campo("📄 PEDIMENTO:", 0, 1)
        self.entry_fecha_entrada = crear_campo("📅 FECHA ENTRADA:", 0, 2)
        self.entry_fecha_verificacion = crear_campo("📅 FECHA VERIFICACIÓN:", 1, 0)
        self.entry_firma = crear_campo("✍️ FIRMA:", 1, 1)
        self.entry_fecha_emision = crear_campo("📅 FECHA EMISIÓN:", 1, 2)

        # ------------------ ARCHIVOS ------------------
        archivos_frame = ctk.CTkFrame(main_container, fg_color=COLORES["gris_claro"], corner_radius=15, border_width=1, border_color=COLORES["gris_oscuro"])
        archivos_frame.pack(fill="x", pady=5, padx=5)

        ctk.CTkLabel(archivos_frame, text="CARGAR ARCHIVOS REQUERIDOS",
                    font=(FUENTE_PRINCIPAL, 14, "bold"),
                    text_color=COLORES["negro"]).pack(pady=(5, 5))

        # --- Layout ---
        layout_frame = ctk.CTkFrame(archivos_frame, fg_color=COLORES["gris_claro"])
        layout_frame.pack(fill="x", pady=3, padx=10)

        ctk.CTkLabel(layout_frame, text="📋 ARCHIVO LAYOUT:",
                    font=(FUENTE_PRINCIPAL, 12, "bold"),
                    text_color=COLORES["negro"]).pack(side="left")

        self.lbl_layout_status = ctk.CTkLabel(layout_frame, text="No cargado",
                                            font=BTN_FONT,
                                            text_color=COLORES["gris_oscuro"])
        self.lbl_layout_status.pack(side="left", padx=5)

        btn_layout = ctk.CTkButton(layout_frame, text="📂 Seleccionar Layout",
                                command=self.cargar_layout,
                                fg_color=COLORES["negro"], hover_color=COLORES["gris_oscuro"],
                                font=BTN_FONT, width=BTN_WIDTH, height=BTN_HEIGHT, corner_radius=15)
        btn_layout.pack(side="right")

        # --- Emmanuel ---
        emmanuel_frame = ctk.CTkFrame(archivos_frame, fg_color=COLORES["gris_claro"])
        emmanuel_frame.pack(fill="x", pady=3, padx=10)

        ctk.CTkLabel(emmanuel_frame, text="📄 ARCHIVO EMMANUEL:",
                    font=(FUENTE_PRINCIPAL, 12, "bold"),
                    text_color=COLORES["negro"]).pack(side="left")

        self.lbl_emmanuel_status = ctk.CTkLabel(emmanuel_frame, text="No cargado",
                                                font=BTN_FONT,
                                                text_color=COLORES["gris_oscuro"])
        self.lbl_emmanuel_status.pack(side="left", padx=5)

        btn_emmanuel = ctk.CTkButton(emmanuel_frame, text="📂 Seleccionar Emmanuel",
                                    command=self.cargar_emmanuel,
                                    fg_color=COLORES["negro"], hover_color=COLORES["gris_oscuro"],
                                    font=BTN_FONT, width=BTN_WIDTH, height=BTN_HEIGHT, corner_radius=15)
        btn_emmanuel.pack(side="right")

        # ------------------ PANEL DE ACCIÓN ------------------
        accion_frame = ctk.CTkFrame(main_container, fg_color=COLORES["blanco"], corner_radius=15, border_width=1, border_color=COLORES["gris_claro"])
        accion_frame.pack(fill="x", pady=10, padx=5)

        self.info_label = ctk.CTkLabel(accion_frame,
                                    text="Seleccione los archivos requeridos para generar la tabla",
                                    font=BTN_FONT,
                                    text_color=COLORES["gris_oscuro"])
        self.info_label.pack(pady=3)

        self.lbl_contador_folios = ctk.CTkLabel(accion_frame,
                                                text="Folios disponibles: 0 | Siguiente folio: N/A",
                                                font=BTN_FONT,
                                                text_color=COLORES["negro"])
        self.lbl_contador_folios.pack(pady=3)

        btn_cargar_folios = ctk.CTkButton(accion_frame, text="📂 Cargar folios",
                                        command=self.cargar_folios,
                                        fg_color=COLORES["negro"], hover_color=COLORES["gris_oscuro"],
                                        font=BTN_FONT, width=BTN_WIDTH, height=BTN_HEIGHT, corner_radius=15)
        btn_cargar_folios.pack(pady=3)

        self.btn_generar = ctk.CTkButton(accion_frame, text="🚀 GENERAR TABLA DE RELACIÓN",
                                        command=self.generar_tabla,
                                        fg_color=COLORES["negro"], hover_color=COLORES["gris_oscuro"],
                                        text_color=COLORES["blanco"],
                                        font=(FUENTE_PRINCIPAL, 14, "bold"),
                                        width=BTN_WIDTH, height=BTN_HEIGHT, corner_radius=15)
        self.btn_generar.pack(pady=5)
        self.btn_generar.configure(state="disabled")

        # ------------------ FOOTER ------------------
        footer_frame = ctk.CTkFrame(main_container, fg_color=COLORES["blanco"])
        footer_frame.pack(fill="x", pady=(5, 0))

        ctk.CTkLabel(footer_frame,
                    text="Sistema de Generación de Tablas de Relación ULTA © 2025",
                    font=(FUENTE_SECUNDARIA, 9),
                    text_color=COLORES["gris_oscuro"]).pack()

    def verificar_archivos_cargados(self):
        if self.layout is not None:
            self.btn_generar.configure(state="normal")
            self.info_label.configure(text="✅ Layout cargado listo para generar tabla de relación")
        else:
            self.btn_generar.configure(state="disabled")
            self.info_label.configure(text="⏳ Esperando Layout:")

    def convertir_a_json(self, archivo_excel, hoja, nombre_json):
        try:
            if "layout" in nombre_json.lower():
                COLUMNAS_REQUERIDAS = [
                    "Folio de Solicitud", "NOM", "Número de Acreditación", "RFC",
                    "Denominación social o nombre", "Tipo de persona", "Marca del producto",
                    "Descripción del producto", "Fracción arancelaria", "Fecha de envío de la solicitud",
                    "Vigencia de la Solicitud", "Modalidad de etiquetado", "Modelo", "UMC",
                    "Cantidad", "Número de etiquetas a verificar", "Parte", "Partida",
                    "Pais Origen", "Pais Comprador"
                ]
                header_row = 2

            elif "emmanuel" in nombre_json.lower():
                COLUMNAS_REQUERIDAS = [
                    "# PARTIDA - FRACCION", "FACTURA", "# ORDEN - ITEM", "UPC", "MARCA", "PAIS",
                    "DESC. FACTURA", "DESC. PEDIMENTO", "UNIDAD VU", "CANTIDAD EN VU",
                    "UNI. TARIFA", "CANT. TARIFA", "UNI. FACT.", "CANT. FACT.", "PRECIO UNIT.",
                    "TOTAL", "FRACCION", "NICO", "FRACCION CORRELACION", "NICO CORRELACION",
                    "CORRELACION", "MET. VAL.", "VINCULACION", "DESCUENTO", "MARCA", "MODELO",
                    "SERIE", "ESPECIAL", "ORDEN - ITEM", "IMPUESTO", "TASA", "IMPORTE TT",
                    "P/I", "C1", "C2", "C3", "NUM", "FIRMA"
                ]
                header_row = 0

            else:
                raise ValueError("Archivo no reconocido")

            df = pd.read_excel(archivo_excel, sheet_name=hoja, header=header_row, dtype=str)
            df.columns = df.columns.str.strip()
            columnas_presentes = [col for col in COLUMNAS_REQUERIDAS if col in df.columns]
            df_filtrado = df[columnas_presentes].dropna(how="all")
            data = df_filtrado.to_dict(orient="records")

            destino = os.path.join(self.carpeta_resources, nombre_json)
            with open(destino, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)

            print(f"✅ Archivo convertido a JSON: {destino}")
            return data
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo convertir {archivo_excel}:\n{e}")
            return None

    def cargar_layout(self):
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo LAYOUT",
            filetypes=[("Excel files", "*.xls *.xlsx")]
        )
        if archivo:
            try:
                self.layout = self.convertir_a_json(archivo, "Layout 1", "layout.json")
                nombre_archivo = os.path.basename(archivo)

                if self.layout and len(self.layout) > 0:
                    self.lbl_layout_status.configure(text=f"✅ {nombre_archivo}", text_color="#2ecc71")
                    self.info_label.configure(text="✅ Layout cargado correctamente")
                else:
                    self.lbl_layout_status.configure(text=f"❌ {nombre_archivo} inválido", text_color="#e74c3c")
                    self.info_label.configure(text="❌ Layout no contiene datos válidos")

            except Exception as e:
                self.layout = None
                self.lbl_layout_status.configure(text=f"❌ Error al cargar", text_color="#e74c3c")
                self.info_label.configure(text=f"❌ Error al cargar layout")
                messagebox.showerror("Error", f"No se pudo cargar el layout:\n{e}")

            # Siempre verificar si se puede habilitar el botón
            self.verificar_archivos_cargados()

    def cargar_emmanuel(self):
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo EMMANUEL",
            filetypes=[("Excel files", "*.xls *.xlsx")]
        )
        if archivo:
            self.emmanuel = self.convertir_a_json(archivo, 0, "EMMANUEL.json")
            if self.emmanuel:
                nombre_archivo = os.path.basename(archivo)
                self.lbl_emmanuel_status.configure(text=f"✅ {nombre_archivo}", text_color="#2ecc71")
                self.verificar_archivos_cargados()

    def cargar_folios_json(self):
        if os.path.exists(FOLIOS_FILE):
            with open(FOLIOS_FILE, "r", encoding="utf-8") as f:
                self.folios_disponibles = json.load(f)
        else:
            self.folios_disponibles = []

        with open(FOLIOS_FILE, "w", encoding="utf-8") as f:
            json.dump(self.folios_disponibles, f, indent=4)

    def cargar_folios(self):
        """Permite al usuario cargar un archivo Excel de folios (reemplaza o agrega)"""
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo de folios",
            filetypes=[("Excel files", "*.xls *.xlsx")]
        )
        if archivo:
            try:
                df = pd.read_excel(archivo, sheet_name=0, header=0, dtype=str)
                folios_nuevos = df.iloc[:, 1].dropna().astype(int).tolist()
                # Asegurar 6 dígitos
                folios_nuevos = [str(f).zfill(6) for f in folios_nuevos]

                # Preguntar si quiere reemplazar o agregar
                if self.folios_disponibles:
                    respuesta = messagebox.askyesno(
                        "Agregar o reemplazar",
                        "Ya existen folios cargados.\n"
                        "¿Desea reemplazarlos? (Sí = reemplazar, No = agregar al final)"
                    )
                    if respuesta:  # Reemplazar
                        self.folios_disponibles = folios_nuevos
                    else:  # Agregar
                        self.folios_disponibles.extend(folios_nuevos)
                else:
                    self.folios_disponibles = folios_nuevos

                # Guardar JSON persistente
                os.makedirs(RESOURCES_DIR, exist_ok=True)
                with open(FOLIOS_FILE, "w", encoding="utf-8") as f:
                    json.dump(self.folios_disponibles, f, indent=4)

                self.actualizar_contador_folios()
                messagebox.showinfo("Éxito", f"{len(folios_nuevos)} folios cargados correctamente")

            except Exception as e:
                messagebox.showerror("Error", f"No se pudo cargar el archivo:\n{e}")

    def actualizar_contador_folios(self):
        """Actualiza el label que muestra folios disponibles"""
        if hasattr(self, "lbl_contador_folios"):
            total = len(self.folios_disponibles)
            siguiente = self.folios_disponibles[0] if total > 0 else "N/A"
            self.lbl_contador_folios.configure(
                text=f"Folios disponibles: {total} | Siguiente folio: {siguiente}"
            )

    def generar_tabla(self):
        try:
            self.info_label.configure(text="🔄 Procesando archivos...", text_color=COLORES["negro"])
            self.update()

            # LEER ARCHIVOS CON FUNCIONES MEJORADAS
            df_base = leer_json(resource_path("resources/base_general.json"))
            df_layout = leer_json(os.path.join("resources", "layout.json"))
            
            # Cargar Emmanuel solo si existe, sino DataFrame vacío
            df_emmanuel = pd.DataFrame()
            if self.emmanuel is not None:
                try:
                    df_emmanuel = leer_json(os.path.join("resources", "EMMANUEL.json"))
                except:
                    df_emmanuel = pd.DataFrame()

            # CARGAR PAISES CON FUNCIÓN MEJORADA
            paises_dict = cargar_paises()

            # CREAR TABLA FINAL (MANTENIENDO TU LÓGICA EXACTA)
            COLUMNAS_RELACION = [
                "SOLICITUD", "LISTA", "PEDIMENTO", "FECHA ENTRADA", "FECHA DE VERIFICACION",
                "MARCA", "CODIGO", "FACTURA", "CANTIDAD", "PAIS DE ORIGEN",
                "DESCRIPCION", "CONTENIDO", "INSUMO", "FORRO", "CLASF UVA", "NORMA UVA",
                "ESTATUS", "FIRMA", "OBSERVACIONES", "OBSERVACIONES DE DICTAMEN",
                "TIPO DE DOCUMENTO", "FOLIO", "MEDIDAS", "PAUS DE PROCEDENCIA",
                "TIPO DE LISTA", "FECHA DE EMISION DE SOLICITUD", "PUNTO DPNS",
                "NO DE INVENTARIO DE MEDICION", "ASIGNACION"
            ]
            tabla_relacion = pd.DataFrame(columns=COLUMNAS_RELACION)

            # MAPEO DE DATOS (TU LÓGICA ORIGINAL)
            if "Denominación social o nombre" in df_layout.columns:
                tabla_relacion["MARCA"] = df_layout["Denominación social o nombre"]
            else:
                tabla_relacion["MARCA"] = ""

            if "Parte" in df_layout.columns:
                tabla_relacion["CODIGO"] = df_layout["Parte"].apply(
                    lambda x: str(int(float(x))) if pd.notnull(x) and str(x).replace(".", "").isdigit() else str(x)
                )
            else:
                tabla_relacion["CODIGO"] = ""

            if "Folio de Solicitud" in df_layout.columns:
                tabla_relacion["FACTURA"] = df_layout["Folio de Solicitud"]
            else:
                tabla_relacion["FACTURA"] = ""

            if "Cantidad" in df_layout.columns:
                tabla_relacion["CANTIDAD"] = df_layout["Cantidad"]
            else:
                tabla_relacion["CANTIDAD"] = ""

            if "Pais Origen" in df_layout.columns:
                tabla_relacion["PAIS DE ORIGEN"] = df_layout["Pais Origen"].astype(str).str.strip().str.upper().map(
                    lambda x: paises_dict.get(x, x)
                )
            else:
                tabla_relacion["PAIS DE ORIGEN"] = ""

            def extraer_numero_nom(texto):
                if pd.isna(texto):
                    return ""
                match = re.search(r"\d+", str(texto))
                return match.group(0) if match else ""

            if "NOM" in df_layout.columns:
                tabla_relacion["CLASF UVA"] = df_layout["NOM"].astype(str).str.extract(r'(\d+)')[0].fillna(0).astype(int)
                tabla_relacion["NORMA UVA"] = df_layout["NOM"].apply(extraer_numero_nom)
            else:
                tabla_relacion["CLASF UVA"] = ""
                tabla_relacion["NORMA UVA"] = ""

            tabla_relacion["CLASF UVA"] = pd.to_numeric(tabla_relacion["CLASF UVA"], errors="coerce").fillna(0).astype(int)
            tabla_relacion["NORMA UVA"] = pd.to_numeric(tabla_relacion["NORMA UVA"], errors="coerce").fillna(0).astype(int)

            # BUSCAR DESCRIPCION Y CONTENIDO EN BASE_GENERAL
            if "Parte" not in df_layout.columns:
                raise ValueError("❌ La columna 'Parte' no existe en layout.json")
            if "UPC" not in df_base.columns:
                raise ValueError("❌ La columna 'UPC' no existe en base_general.json")

            columnas_base_necesarias = ["Denominación genérica", "CONTENIDO", "CATEGORIA"]
            for col in columnas_base_necesarias:
                if col not in df_base.columns:
                    raise ValueError(f"❌ La columna '{col}' no existe en base_general.json")

            def clean_codigo(x):
                if pd.isnull(x):
                    return ""
                try:
                    return str(int(float(x)))
                except:
                    return str(x).strip()

            df_layout["PARTE_CLEAN"] = df_layout["Parte"].apply(clean_codigo)
            df_base["UPC_CLEAN"] = df_base["UPC"].apply(clean_codigo)

            dict_descripcion = dict(zip(df_base["UPC_CLEAN"], df_base["Denominación genérica"]))
            dict_contenido = dict(zip(df_base["UPC_CLEAN"], df_base["CONTENIDO"]))
            dict_asignacion = dict(zip(df_base["UPC_CLEAN"], df_base["CATEGORIA"]))

            tabla_relacion["DESCRIPCION"] = df_layout["PARTE_CLEAN"].map(dict_descripcion).fillna("NO ENCONTRADO")
            tabla_relacion["CONTENIDO"] = df_layout["PARTE_CLEAN"].map(dict_contenido).fillna("NO ENCONTRADO")
            tabla_relacion["ASIGNACION"] = df_layout["PARTE_CLEAN"].map(dict_asignacion).fillna("NO ENCONTRADO")

            df_layout.drop("PARTE_CLEAN", axis=1, inplace=True)

            # REASIGNAR LISTA SEGÚN ASIGNACION
            if "ASIGNACION" in tabla_relacion.columns:
                tabla_relacion["ASIGNACION_NUM"] = pd.to_numeric(tabla_relacion["ASIGNACION"], errors="coerce").fillna(0).astype(int)
                tabla_relacion.sort_values("ASIGNACION_NUM", inplace=True, ignore_index=True)
                
                lista = []
                counter = 1
                for asign_val, group in tabla_relacion.groupby("ASIGNACION_NUM"):
                    lista.extend([counter]*len(group))
                    counter += 1
                
                tabla_relacion["LISTA"] = lista
                tabla_relacion.drop(columns=["ASIGNACION_NUM"], inplace=True)

            # DATOS MANUALES
            solicitud = self.entry_solicitud.get()
            firma = self.entry_firma.get()
            pedimento = self.entry_pedimento.get().strip()

            def parsear_fecha(fecha_str):
                try:
                    return datetime.strptime(fecha_str, "%d/%m/%y")
                except:
                    return None

            fecha_entrada_dt = parsear_fecha(self.entry_fecha_entrada.get())
            fecha_verificacion_dt = parsear_fecha(self.entry_fecha_verificacion.get())
            fecha_emision_dt = parsear_fecha(self.entry_fecha_emision.get())

            tabla_relacion["SOLICITUD"] = solicitud
            tabla_relacion["PEDIMENTO"] = pedimento
            tabla_relacion["FECHA ENTRADA"] = fecha_entrada_dt.strftime("%d/%m/%Y") if fecha_entrada_dt else ""
            tabla_relacion["FECHA DE VERIFICACION"] = fecha_verificacion_dt.strftime("%d/%m/%Y") if fecha_verificacion_dt else ""
            tabla_relacion["FIRMA"] = firma
            tabla_relacion["FECHA DE EMISION DE SOLICITUD"] = fecha_emision_dt.strftime("%d/%m/%Y") if fecha_emision_dt else ""



            # Asignacion de folios
            # -----------------------------
            # ASIGNACIÓN DE FOLIOS SEGÚN LISTA
            # -----------------------------
            if hasattr(self, "folios_disponibles") and self.folios_disponibles:
                folios_asignados = []
                lista_max = tabla_relacion["LISTA"].max()
                for i in range(1, lista_max + 1):
                    if self.folios_disponibles:
                        folio = self.folios_disponibles.pop(0)
                        with open(FOLIOS_FILE, "w", encoding="utf-8") as f:
                            json.dump(self.folios_disponibles, f, indent=4)
                        # Asegurar que el folio tenga siempre 6 dígitos
                        folio = str(folio).zfill(6)
                    else:
                        folio = "SIN FOLIO"
                    tabla_relacion.loc[tabla_relacion["LISTA"] == i, "FOLIO"] = folio
                    folios_asignados.append(folio)
                
                # Guardar JSON actualizado
                with open(FOLIOS_FILE, "w", encoding="utf-8") as f:
                    json.dump(self.folios_disponibles, f, indent=4)
                self.actualizar_contador_folios()
            else:
                tabla_relacion["FOLIO"] = "SIN FOLIO"

        
            
            # RESTO DE COLUMNAS FIJAS
            tabla_relacion["ESTATUS"] = "N/A"
            tabla_relacion["OBSERVACIONES"] = "N/A"
            tabla_relacion["OBSERVACIONES DE DICTAMEN"] = "N/A"
            tabla_relacion["TIPO DE DOCUMENTO"] = "D"
            tabla_relacion["MEDIDAS"] = "N/A"
            tabla_relacion["PAUS DE PROCEDENCIA"] = "E.U.A."
            tabla_relacion["TIPO DE LISTA"] = "N/A"
            tabla_relacion["PUNTO DPNS"] = "N/A"
            tabla_relacion["NO DE INVENTARIO DE MEDICION"] = "N/A"
            tabla_relacion["INSUMO"] = "N/A"
            tabla_relacion["FORRO"] = "N/A"

            # ASIGNAR FACTURAS DESDE EMMANUEL (SOLO SI EXISTE)
            if not df_emmanuel.empty and "FACTURA" in df_emmanuel.columns and "# ORDEN - ITEM" in df_emmanuel.columns:
                tabla_relacion["CODIGO_STR"] = tabla_relacion["CODIGO"].astype(str).str.strip()
                df_emmanuel["CODIGO_STR"] = df_emmanuel["# ORDEN - ITEM"].astype(str).str.split(" - ").str[-1].str.strip()

                tabla_relacion["CODIGO_STR"] = tabla_relacion["CODIGO_STR"].str.replace(r"\s+", "", regex=True)
                df_emmanuel["CODIGO_STR"] = df_emmanuel["CODIGO_STR"].str.replace(r"\s+", "", regex=True)

                mapa_facturas = dict(zip(df_emmanuel["CODIGO_STR"], df_emmanuel["FACTURA"]))
                tabla_relacion["FACTURA"] = tabla_relacion["CODIGO_STR"].map(mapa_facturas).fillna(tabla_relacion["FACTURA"])

                tabla_relacion.drop(columns=["CODIGO_STR"], inplace=True)

            # GUARDAR EXCEL
            if "CODIGO" in tabla_relacion.columns:
                tabla_relacion["CODIGO"] = pd.to_numeric(tabla_relacion["CODIGO"], errors="coerce").fillna(0).astype("int64")

            if "CANTIDAD" in tabla_relacion.columns:
                tabla_relacion["CANTIDAD"] = pd.to_numeric(tabla_relacion["CANTIDAD"], errors="coerce").fillna(0).astype("int64")

            salida = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Guardar tabla de relación"
            )
            
            if salida:
                # APLICAR MACHOTE CON FUNCIÓN MEJORADA
                machote = cargar_machote()

                with pd.ExcelWriter(salida, engine="xlsxwriter") as writer:
                    tabla_relacion.to_excel(writer, index=False, sheet_name="TablaRelacion")
                    workbook = writer.book
                    worksheet = writer.sheets["TablaRelacion"]

                    header_cfg = machote.get("header", {})
                    font_cfg = header_cfg.get("font", {})

                    formato_header = workbook.add_format({
                        "bold": font_cfg.get("bold", True),
                        "font_name": font_cfg.get("name", "Calibri"),
                        "font_size": font_cfg.get("size", 11),
                        "align": font_cfg.get("align", "center"),
                        "valign": font_cfg.get("valign", "vcenter"),
                        "bg_color": header_cfg.get("color", "#FF9900"),
                        "font_color": "#000000"
                    })

                    formato_entero = workbook.add_format({"num_format": "0", "align": "left"})

                    for col_num, value in enumerate(tabla_relacion.columns.values):
                        worksheet.write(0, col_num, value, formato_header)

                    col_cfg = machote.get("columns", {})
                    auto_adjust = col_cfg.get("auto_adjust", True)
                    min_width = col_cfg.get("min_width", 10)
                    max_width = col_cfg.get("max_width", 50)

                    if auto_adjust:
                        for idx, col in enumerate(tabla_relacion.columns):
                            max_len = max(
                                tabla_relacion[col].astype(str).map(len).max(),
                                len(str(col))
                            ) + 2
                            max_len = max(min_width, min(max_len, max_width))
                            worksheet.set_column(idx, idx, max_len)

                    for colname in ["CODIGO", "CANTIDAD"]:
                        if colname in tabla_relacion.columns:
                            col_idx = tabla_relacion.columns.get_loc(colname)
                            worksheet.set_column(col_idx, col_idx, 20, formato_entero)

                self.info_label.configure(text="✅ Tabla generada exitosamente!", text_color="#2ecc71")
                messagebox.showinfo("Éxito", f"Tabla de relación generada:\n{salida}")

        except Exception as e:
            self.info_label.configure(text="❌ Error al generar tabla", text_color="#e74c3c")
            messagebox.showerror("Error", f"Ocurrió un error al generar la tabla:\n{e}")

# EJECUTAR VENTANA
if __name__ == "__main__":
    try:
        app = VentanaULTA()
        app.mainloop()
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo iniciar la aplicación:\n{e}")