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
        if os.path.exists(ruta):
            return pd.read_json(ruta, orient="records")
        ruta_alternativa = resource_path(ruta)
        if os.path.exists(ruta_alternativa):
            return pd.read_json(ruta_alternativa, orient="records")
        print(f"⚠️ Archivo {ruta} no encontrado, creando DataFrame vacío")
        return pd.DataFrame()
    except Exception as e:
        print(f"❌ Error leyendo {ruta}: {e}")
        return pd.DataFrame()

def cargar_paises():
    """Carga el archivo de países con manejo de errores"""
    try:
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
        
        print("⚠️ No se pudo cargar Paises.json, usando diccionario vacío")
        return {}
    except Exception as e:
        print(f"❌ Error cargando países: {e}")
        return {}

def cargar_machote():
    """Carga el archivo machote.json con manejo de errores"""
    try:
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
        
        print("⚠️ No se pudo cargar machote.json, usando valores por defecto")
        return {}
    except Exception as e:
        print(f"❌ Error cargando machote: {e}")
        return {}

# -----------------------------
# CLASE VENTANA PRINCIPAL
# -----------------------------
class VentanaULTA(ctk.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        
        self.folios_disponibles = []
        self.cargar_folios_json()
        self.crear_interfaz()
        self.actualizar_contador_folios()

        self.title("Generador Tablas de Relación ULTA")
        self.geometry("800x580")
        self.minsize(800, 580)
        self.configure(fg_color=COLORES["blanco"])
        self.center_window()

        # Archivos cargados
        self.layout = None
        self.emmanuel = None
        self.base_general = None

        # Carpeta destino - crear si no existe
        self.carpeta_resources = "resources"
        os.makedirs(self.carpeta_resources, exist_ok=True)

        # Inicializar base general al iniciar
        self.inicializar_base_general()

    def center_window(self):
        """Centra la ventana en la pantalla"""
        self.update_idletasks()
        width = 800
        height = 580
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

        self.lbl_emmanuel_status = ctk.CTkLabel(emmanuel_frame, text="No cargado (Opcional)",
                                                font=BTN_FONT,
                                                text_color=COLORES["gris_oscuro"])
        self.lbl_emmanuel_status.pack(side="left", padx=5)

        btn_emmanuel = ctk.CTkButton(emmanuel_frame, text="📂 Seleccionar Emmanuel",
                                    command=self.cargar_emmanuel,
                                    fg_color=COLORES["negro"], hover_color=COLORES["gris_oscuro"],
                                    font=BTN_FONT, width=BTN_WIDTH, height=BTN_HEIGHT, corner_radius=15)
        btn_emmanuel.pack(side="right")

        # --- Base General ---
        base_frame = ctk.CTkFrame(archivos_frame, fg_color=COLORES["gris_claro"])
        base_frame.pack(fill="x", pady=3, padx=10)

        ctk.CTkLabel(base_frame, text="📚 ARCHIVO REL ULTA EMBARQUES:",
                    font=(FUENTE_PRINCIPAL, 12, "bold"),
                    text_color=COLORES["negro"]).pack(side="left")

        self.lbl_base_status = ctk.CTkLabel(base_frame, text="No cargado",
                                            font=BTN_FONT,
                                            text_color=COLORES["gris_oscuro"])
        self.lbl_base_status.pack(side="left", padx=5)

        btn_base = ctk.CTkButton(base_frame, text="📂 Seleccionar Rel ULTA Embarques",
                                command=self.cargar_base_general,
                                fg_color=COLORES["negro"], hover_color=COLORES["gris_oscuro"],
                                font=BTN_FONT, width=BTN_WIDTH, height=BTN_HEIGHT, corner_radius=15)
        btn_base.pack(side="right")

        # ------------------ PANEL DE ACCIÓN ------------------
        accion_frame = ctk.CTkFrame(main_container, fg_color=COLORES["blanco"], corner_radius=15, border_width=1, border_color=COLORES["gris_claro"])
        accion_frame.pack(fill="x", pady=10, padx=5)

        self.info_label = ctk.CTkLabel(accion_frame,
                                    text="Seleccione el archivo Layout para generar la tabla",
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
            # ---------- Layout ----------
            if "layout" in nombre_json.lower():
                COLUMNAS_REQUERIDAS = [
                    "Folio de Solicitud", "NOM", "Número de Acreditación", "RFC",
                    "Denominación social o nombre", "Tipo de persona", "Marca del producto",
                    "Descripción del producto", "Fracción arancelaria", "Fecha de envío de la solicitud",
                    "Vigencia de la Solicitud", "Modalidad de etiquetado", "Modelo", "UMC",
                    "Cantidad", "Número de etiquetas a verificar", "Parte", "Partida",
                    "Pais Origen", "Pais Comprador"
                ]
                df_raw = pd.read_excel(archivo_excel, sheet_name=hoja, header=None, nrows=10, dtype=str)
                header_row = None
                for i, fila in df_raw.iterrows():
                    if "Denominación social o nombre" in fila.values:
                        header_row = i
                        break
                if header_row is None:
                    raise ValueError("❌ No se encontró encabezado válido en el archivo Layout")

                df = pd.read_excel(archivo_excel, sheet_name=hoja, header=header_row, dtype=str)
                df.columns = df.columns.str.strip()
            
            # ---------- Emmanuel ----------
            elif "emmanuel" in nombre_json.lower():
                COLUMNAS_REQUERIDAS = ["FACTURA", "# ORDEN - ITEM"]

                # Leer todo el Excel
                df = pd.read_excel(archivo_excel, sheet_name=hoja, header=0, dtype=str)
                df.columns = df.columns.str.strip().str.upper()

                # Evitar duplicados en columnas automáticamente
                cols = pd.Series(df.columns)
                for dup in cols[cols.duplicated()].unique():
                    cols[cols[cols == dup].index.values.tolist()] = [f"{dup}_{i+1}" for i in range(sum(cols == dup))]
                df.columns = cols

                # Renombrar columnas si es necesario
                if "ORDEN - ITEM" in df.columns and "# ORDEN - ITEM" not in df.columns:
                    df.rename(columns={"ORDEN - ITEM": "# ORDEN - ITEM"}, inplace=True)

                # Filtrar solo las columnas necesarias
                columnas_presentes = [col for col in COLUMNAS_REQUERIDAS if col in df.columns]
                if not columnas_presentes:
                    raise ValueError("No se encontraron las columnas requeridas en Emmanuel")
                
                df_filtrado = df[columnas_presentes].copy()

                # Guardar JSON
                os.makedirs(self.carpeta_resources, exist_ok=True)
                destino = os.path.join(self.carpeta_resources, "emmanuel.json")
                df_filtrado.to_json(destino, orient="records", force_ascii=False, indent=4)

                return df_filtrado.to_dict(orient="records")

            # ---------- Base General ----------
            elif "base_general" in nombre_json.lower():
                COLUMNAS_REQUERIDAS = [
                    "SOLICITUD", "LISTA", "PEDIMENTO", "FECHA DE ENTRADA",
                    "FECHA DE VERIFICACION", "MARCA", "CODIGO", "FACTURA",
                    "CANTIDAD", "PAIS DE ORIGEN", "DESCRIPCION", "CONTENIDO",
                    "ASIGNACIÓN"
                ]
                header_row = 0
                df = pd.read_excel(archivo_excel, sheet_name=hoja, header=header_row, dtype=str)
                df.columns = df.columns.str.strip()

            else:
                raise ValueError("Archivo no reconocido")

            # Para Layout y Base General: filtrar columnas
            if "emmanuel" not in nombre_json.lower():
                faltantes = [col for col in COLUMNAS_REQUERIDAS if col not in df.columns]
                if faltantes:
                    raise ValueError(f"❌ Faltan columnas en {nombre_json}: {', '.join(faltantes)}")

                df_filtrado = df[COLUMNAS_REQUERIDAS].dropna(how="all")

                # Guardar JSON
                os.makedirs(self.carpeta_resources, exist_ok=True)
                destino = os.path.join(self.carpeta_resources, nombre_json)
                with open(destino, "w", encoding="utf-8") as f:
                    json.dump(df_filtrado.to_dict(orient="records"), f, indent=4, ensure_ascii=False)

                return df_filtrado.to_dict(orient="records")

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
            try:
                self.emmanuel = self.convertir_a_json(archivo, 0, "emmanuel.json")
                if self.emmanuel:
                    nombre_archivo = os.path.basename(archivo)
                    self.lbl_emmanuel_status.configure(text=f"✅ {nombre_archivo}", text_color="#2ecc71")
                    # Mostrar cuántas facturas se cargaron
                    facturas = {str(row.get("FACTURA", "")).strip() for row in self.emmanuel if str(row.get("FACTURA", "")).strip() and str(row.get("FACTURA", "")).strip() != "nan"}
                    self.info_label.configure(text=f"✅ Layout y Emmanuel cargados ({len(facturas)} facturas detectadas)")
            except Exception as e:
                self.emmanuel = None
                self.lbl_emmanuel_status.configure(text="❌ Error al cargar", text_color="#e74c3c")
                print(f"⚠️ Emmanuel no cargado: {e}")
        else:
            self.emmanuel = None
            self.lbl_emmanuel_status.configure(text="No cargado (Opcional)", text_color=COLORES["gris_oscuro"])

    def inicializar_base_general(self):
        """Carga la ruta de Base General desde config.json si existe"""
        try:
            config_path = os.path.join("resources", "config.json")
            os.makedirs("resources", exist_ok=True)

            if os.path.exists(config_path):
                with open(config_path, "r", encoding="utf-8") as f:
                    config = json.load(f)

                excel_path = config.get("base_general_excel")
                json_path = config.get("base_general_json")

                if excel_path and os.path.exists(excel_path):
                    self.convertir_a_json(excel_path, hoja=0, nombre_json="base_general.json")
                    self.lbl_base_status.configure(
                        text=f"✅ {os.path.basename(excel_path)}",
                        text_color="#2ecc71"
                    )
                elif json_path and os.path.exists(json_path):
                    destino = os.path.join("resources", "base_general.json")
                    with open(json_path, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    with open(destino, "w", encoding="utf-8") as f:
                        json.dump(data, f, indent=4, ensure_ascii=False)
                    self.lbl_base_status.configure(
                        text=f"✅ {os.path.basename(json_path)}",
                        text_color="#2ecc71"
                    )
                else:
                    self.lbl_base_status.configure(
                        text="No cargado",
                        text_color=COLORES["gris_oscuro"]
                    )
            else:
                self.lbl_base_status.configure(
                    text="No cargado",
                    text_color=COLORES["gris_oscuro"]
                )

        except Exception as e:
            self.lbl_base_status.configure(text="❌ Error al iniciar", text_color="#e74c3c")
            print(f"Error al cargar base_general al iniciar: {e}")

    def cargar_base_general(self):
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo Base General",
            filetypes=[("Archivos Excel", "*.xls *.xlsx"), ("Archivos JSON", "*.json")]
        )

        if archivo:
            try:
                config_path = os.path.join("resources", "config.json")
                os.makedirs("resources", exist_ok=True)

                if archivo.endswith((".xls", ".xlsx")):
                    # Convertir Excel a JSON
                    data = self.convertir_a_json(archivo, hoja=0, nombre_json="base_general.json")
                    if data is None:
                        raise ValueError("Error al convertir Excel a JSON")

                    with open(config_path, "w", encoding="utf-8") as f:
                        json.dump({"base_general_excel": archivo}, f, indent=4, ensure_ascii=False)

                else:  # JSON directo
                    with open(archivo, "r", encoding="utf-8") as f:
                        data = json.load(f)

                    destino = os.path.join("resources", "base_general.json")
                    with open(destino, "w", encoding="utf-8") as f:
                        json.dump(data, f, indent=4, ensure_ascii=False)

                    with open(config_path, "w", encoding="utf-8") as f:
                        json.dump({"base_general_json": archivo}, f, indent=4, ensure_ascii=False)

                self.lbl_base_status.configure(text=f"✅ {os.path.basename(archivo)}", text_color="#2ecc71")
                messagebox.showinfo("Éxito", "Base General cargada correctamente")

            except Exception as e:
                self.lbl_base_status.configure(text="❌ Error", text_color="#e74c3c")
                messagebox.showerror("Error", f"No se pudo cargar la Base General:\n{e}")

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

            # LEER ARCHIVOS - SOLO LAYOUT ES OBLIGATORIO
            df_layout = leer_json(os.path.join("resources", "layout.json"))
            if df_layout.empty:
                raise ValueError("No se pudo cargar el archivo Layout")
            
            # Cargar Emmanuel solo si existe, sino DataFrame vacío
            df_emmanuel = pd.DataFrame()
            if self.emmanuel is not None:
                try:
                    df_emmanuel = leer_json(os.path.join("resources", "emmanuel.json"))
                except Exception as e:
                    print(f"⚠️ No se pudo cargar Emmanuel: {e}")
                    df_emmanuel = pd.DataFrame()

            # Cargar Base General solo si existe
            df_base = pd.DataFrame()
            try:
                df_base = leer_json(os.path.join("resources", "base_general.json"))
            except Exception as e:
                print(f"⚠️ No se pudo cargar Base General: {e}")
                df_base = pd.DataFrame()

            # CARGAR PAISES
            paises_dict = cargar_paises()

            # CREAR TABLA FINAL
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

            # MAPEO DE DATOS DESDE LAYOUT
            if "Marca del producto" in df_layout.columns:
                tabla_relacion["MARCA"] = df_layout["Marca del producto"]
            elif "Denominación social o nombre" in df_layout.columns:
                tabla_relacion["MARCA"] = df_layout["Denominación social o nombre"]
            else:
                tabla_relacion["MARCA"] = ""

            # --- CODIGO ---
            if "Parte" in df_layout.columns:
                def procesar_codigo(x):
                    """Mantiene letras, números y símbolos; convierte solo si es estrictamente numérico"""
                    if pd.isna(x):
                        return ""
                    x = str(x).strip()
                    # Si es numérico puro, quitar decimales .0 pero mantener formato entero
                    if re.fullmatch(r"\d+(\.0+)?", x):
                        return str(int(float(x)))
                    # En cualquier otro caso, mantener texto como viene
                    return x

                tabla_relacion["CODIGO"] = df_layout["Parte"].apply(procesar_codigo)
            else:
                tabla_relacion["CODIGO"] = ""

            # --- Asignar Negaciones o Dictamenes --- #
            if "NOM" in df_layout.columns and "Parte" in df_layout.columns:
                estatus_dict = {}

                # Iterar por cada fila del layout
                for _, fila in df_layout.iterrows():
                    codigo = str(fila.get("Parte", "")).strip()
                    nom_valor = str(fila.get("NOM", "")).upper().strip()

                    if codigo:
                        if codigo not in estatus_dict:
                            estatus_dict[codigo] = ""

                        if "NOM004" in nom_valor:
                            estatus_dict[codigo] = "NEGACIÓN ND"
                        elif "NOM050" in nom_valor:
                            estatus_dict[codigo] = "DICTAMEN"

                # Asignar el estatus correspondiente a cada código
                tabla_relacion["ESTATUS"] = tabla_relacion["CODIGO"].map(estatus_dict).fillna("")
            else:
                tabla_relacion["ESTATUS"] = ""


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


           # --- Aplicar negaciones para NOM004 --- # 
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

            # --- ASIGNAR TIPO DE DOCUMENTO SEGÚN CLASF UVA Y NORMA UVA ---
            def determinar_tipo_documento(clasf, norma):
                try:
                    clasf = int(clasf)
                    norma = int(norma)
                    # Regla especial para NOM 4 y NOM 50
                    if clasf == 4 and norma == 4:
                        return "ND"
                    elif clasf == 50 and norma == 50:
                        return "D"
                    else:
                        # Todas las demás NOM (como 141, etc.)
                        return "D"
                except:
                    return "D"

            tabla_relacion["TIPO DE DOCUMENTO"] = tabla_relacion.apply(
                lambda fila: determinar_tipo_documento(fila["CLASF UVA"], fila["NORMA UVA"]),
                axis=1
            )



            # BUSCAR DESCRIPCION Y CONTENIDO EN BASE_GENERAL (si está disponible)
            if not df_base.empty and "CODIGO" in df_base.columns:
                print("🔍 Usando Base General para descripción y contenido...")
                
                columnas_base_necesarias = ["DESCRIPCION", "CONTENIDO", "ASIGNACIÓN"]
                for col in columnas_base_necesarias:
                    if col not in df_base.columns:
                        print(f"⚠️ Columna '{col}' no encontrada en Base General")
                        continue

                def clean_codigo(x):
                    if pd.isnull(x):
                        return ""
                    try:
                        return str(int(float(x)))
                    except:
                        return str(x).strip()

                df_layout["PARTE_CLEAN"] = df_layout["Parte"].apply(clean_codigo)
                df_base["CODIGO_CLEAN"] = df_base["CODIGO"].apply(clean_codigo)

                # Usar Base General para descripción y contenido
                dict_descripcion = dict(zip(df_base["CODIGO_CLEAN"], df_base["DESCRIPCION"]))
                dict_contenido = dict(zip(df_base["CODIGO_CLEAN"], df_base["CONTENIDO"]))
                dict_asignacion = dict(zip(df_base["CODIGO_CLEAN"], df_base["ASIGNACIÓN"]))

                tabla_relacion["DESCRIPCION"] = df_layout["PARTE_CLEAN"].map(dict_descripcion).fillna("")
                tabla_relacion["CONTENIDO"] = df_layout["PARTE_CLEAN"].map(dict_contenido).fillna("")
                tabla_relacion["ASIGNACION"] = df_layout["PARTE_CLEAN"].map(dict_asignacion).fillna("1")

                df_layout.drop("PARTE_CLEAN", axis=1, inplace=True)
                
                print("✅ Base General aplicada correctamente")
            else:
                # Si no hay base general, usar valores del Layout
                print("⚠️ Base General no disponible, usando datos del Layout")
                if "Descripción del producto" in df_layout.columns:
                    tabla_relacion["DESCRIPCION"] = df_layout["Descripción del producto"]
                else:
                    tabla_relacion["DESCRIPCION"] = ""
                tabla_relacion["CONTENIDO"] = ""
                tabla_relacion["ASIGNACION"] = "1"

            
            # -----------------------------
            # REASIGNAR LISTA SEGÚN ASIGNACION (SIN MODIFICAR ASIGNACION)
            # REASIGNAR LISTA SEGÚN ASIGNACION, repitiendo cada número 11 veces
            # -----------------------------
            if "ASIGNACION" in tabla_relacion.columns:
                REPETICIONES = 11

                # Aseguramos que ASIGNACION sea string limpio
                tabla_relacion["ASIGNACION"] = tabla_relacion["ASIGNACION"].fillna("").astype(str).str.strip()

                # Función para separar en (numero, sufijo)
                def split_asignacion(val):
                    match = re.match(r"^(\d+)(?:-(\w+))?$", val)
                    if match:
                        num = int(match.group(1))
                        sufijo = match.group(2) if match.group(2) else ""
                        return num, sufijo
                    return float("inf"), ""  # valores inválidos al final

                # Generamos columnas auxiliares para ordenar
                tabla_relacion[["_NUM", "_SUFIJO"]] = tabla_relacion["ASIGNACION"].apply(
                    lambda x: pd.Series(split_asignacion(x))
                )

                # Orden: primero los que tienen sufijo (no vacío), luego los demás
                tabla_relacion["_HAS_LETTER"] = tabla_relacion["_SUFIJO"].apply(lambda x: 0 if x else 1)

                # Ordenar con prioridad: con letra primero, luego numérico ascendente, luego sufijo
                tabla_relacion.sort_values(by=["_HAS_LETTER", "_NUM", "_SUFIJO"],
                                        inplace=True, ignore_index=True)

                # Asignar LISTA (misma lógica de bloques de 11)
                lista = []
                counter = 1
                for asign_val, group in tabla_relacion.groupby("ASIGNACION", sort=False):
                    n_filas = len(group)
                    bloques = (n_filas + REPETICIONES - 1) // REPETICIONES
                    for b in range(bloques):
                        inicio = b * REPETICIONES
                        fin = min((b + 1) * REPETICIONES, n_filas)
                        lista.extend([counter] * (fin - inicio))
                        counter += 1

                tabla_relacion["LISTA"] = lista

                # Limpiar columnas auxiliares
                tabla_relacion.drop(columns=["_NUM", "_SUFIJO", "_HAS_LETTER"], inplace=True)


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

            # -- Datos ingresados por el usuario -- #
            tabla_relacion["SOLICITUD"] = solicitud
            tabla_relacion["PEDIMENTO"] = pedimento
            tabla_relacion["FECHA ENTRADA"] = fecha_entrada_dt.strftime("%d/%m/%Y") if fecha_entrada_dt else ""
            tabla_relacion["FECHA DE VERIFICACION"] = fecha_verificacion_dt.strftime("%d/%m/%Y") if fecha_verificacion_dt else ""
            tabla_relacion["FIRMA"] = firma
            tabla_relacion["FECHA DE EMISION DE SOLICITUD"] = fecha_emision_dt.strftime("%d/%m/%Y") if fecha_emision_dt else ""

            # ASIGNACIÓN DE FOLIOS SEGÚN LISTA
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
            tabla_relacion["MEDIDAS"] = "N/A"
            tabla_relacion["PAUS DE PROCEDENCIA"] = "E.U.A."
            tabla_relacion["TIPO DE LISTA"] = "N/A"
            tabla_relacion["PUNTO DPNS"] = "N/A"
            tabla_relacion["NO DE INVENTARIO DE MEDICION"] = "N/A"
            tabla_relacion["INSUMO"] = "N/A"
            tabla_relacion["FORRO"] = "N/A"

            # ASIGNAR FACTURAS DESDE EMMANUEL (SOLO SI EXISTE)
            def normalizar_codigo(x):
                if pd.isnull(x) or str(x).strip().lower() == "nan":
                    return ""
                x_str = str(x).strip()
                x_str = re.sub(r"\s+", "", x_str)  # quitar espacios
                x_str = x_str.lstrip("0")  # quitar ceros a la izquierda
                return x_str

            # Solo procesar Emmanuel si está cargado y tiene las columnas necesarias
            if not df_emmanuel.empty and "# ORDEN - ITEM" in df_emmanuel.columns and "FACTURA" in df_emmanuel.columns:
                print("🔍 Procesando Emmanuel para mapear facturas...")
                
                # Normalizar códigos en tabla_relacion
                tabla_relacion["CODIGO_STR"] = tabla_relacion["CODIGO"].apply(normalizar_codigo)
                
                # Normalizar códigos en Emmanuel
                df_emmanuel["CODIGO_STR"] = df_emmanuel["# ORDEN - ITEM"].apply(
                    lambda x: normalizar_codigo(str(x).split(" - ")[-1]) if pd.notnull(x) else ""
                )
                
                # Crear mapa de facturas
                mapa_facturas = dict(zip(df_emmanuel["CODIGO_STR"], df_emmanuel["FACTURA"]))
                
                # Mapear facturas
                facturas_mapeadas = 0
                for idx, row in tabla_relacion.iterrows():
                    codigo = row["CODIGO_STR"]
                    if codigo in mapa_facturas and mapa_facturas[codigo]:
                        tabla_relacion.at[idx, "FACTURA"] = mapa_facturas[codigo]
                        facturas_mapeadas += 1
                
                print(f"✅ Facturas mapeadas desde Emmanuel: {facturas_mapeadas}")
                
                # Limpiar columna auxiliar
                tabla_relacion.drop(columns=["CODIGO_STR"], inplace=True)
            else:
                print("⚠️ Emmanuel no disponible o no tiene las columnas requeridas")

            # CONVERTIR TIPOS DE DATOS
            # CODIGO se mantiene como texto (puede ser alfanumérico)
            if "CODIGO" in tabla_relacion.columns:
                tabla_relacion["CODIGO"] = tabla_relacion["CODIGO"].astype(str).str.strip()

            # CANTIDAD sí se convierte a entero
            if "CANTIDAD" in tabla_relacion.columns:
                tabla_relacion["CANTIDAD"] = pd.to_numeric(tabla_relacion["CANTIDAD"], errors="coerce").fillna(0).astype("int64")


            # GUARDAR EXCEL
            salida = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Guardar tabla de relación"
            )
            
            if salida:
                # APLICAR MACHOTE
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
            import traceback
            traceback.print_exc()



# EJECUTAR VENTANA
if __name__ == "__main__":
    try:
        app = VentanaULTA()
        app.mainloop()
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo iniciar la aplicación:\n{e}")