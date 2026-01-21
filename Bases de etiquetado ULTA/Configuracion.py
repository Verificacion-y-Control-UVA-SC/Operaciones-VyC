import os
import json
import customtkinter as ctk
from tkinter import filedialog, messagebox
import shutil
import pandas as pd
import sys

COLORES = {
    "primario": "#ECD925",
    "secundario": "#282828",
    "exito": "#008d53",
    "peligro": "#d74a3d",
    "fondo": "#F8F9FA",
    "surface": "#FFFFFF",
    "texto_oscuro": "#282828",
    "texto_claro": "#4b4b4b",
    "borde": "#BDC3C7",
}

FUENTE = "Inter"

def convertir_a_json(archivo_excel: str, nombre_json: str, persist: bool = True):
    """
    Convierte un archivo Excel a JSON y lo guarda en disco.
    """
    xls = pd.ExcelFile(archivo_excel)
    data = {}
    
    for hoja in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=hoja)
        data[hoja] = df.fillna("").to_dict(orient="records")
    
    if persist:
        with open(nombre_json, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
    return data

class ConfiguracionWindow(ctk.CTkToplevel):
    def __init__(self, parent=None, base_general=None, base_ulta=None):
        super().__init__(parent)
        self.title("⚙️ Configuración - Bases ULTA")
        self.geometry("520x600")  # 🔸 Tamaño más manejable con scroll
        self.minsize(520, 400)   # 🔸 Mínimo más pequeño para pantallas pequeñas
        self.configure(fg_color=COLORES["fondo"])
        self.center_window()
        
        # Hacer la ventana modal
        self.transient(parent)
        self.grab_set()

        # 📂 Carpeta destino y archivo de configuración
        self.data_dir = os.path.join(os.getcwd(), "data")
        os.makedirs(self.data_dir, exist_ok=True)
        self.config_file = os.path.join(self.data_dir, "config.json")

        # 📥 Cargar configuración guardada
        self.config_data = self.cargar_configuracion()

        # 📌 Cargar rutas desde configuración si existen
        self.base_general_path = self.config_data.get("base_general_path")
        self.base_ulta_path = self.config_data.get("base_ulta_path")
        self.base_norma_024_path = self.config_data.get("base_norma_024_path")

        self.crear_interfaz()
        self.actualizar_estados()

        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.getcwd()

        self.data_dir = os.path.join(base_dir, "data")
        os.makedirs(self.data_dir, exist_ok=True)
        self.config_file = os.path.join(self.data_dir, "config.json")

    def center_window(self):
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        ww, wh = 520, 600
        x = (sw - ww) // 2
        y = (sh - wh) // 2
        self.geometry(f"{ww}x{wh}+{x}+{y}")

    def crear_interfaz(self):
        # --- HEADER MEJORADO ---
        header = ctk.CTkFrame(self, fg_color=COLORES["fondo"], height=100, corner_radius=0)
        header.pack(fill="x")
        header.pack_propagate(False)

        header_content = ctk.CTkFrame(header, fg_color="transparent")
        header_content.pack(expand=True, fill="both", padx=30, pady=20)

        ctk.CTkLabel(header_content, 
                     text="⚙️ Configuración de Bases", 
                     font=(FUENTE, 22, "bold"), 
                     text_color=COLORES["texto_oscuro"]).pack(anchor="w")
        
        ctk.CTkLabel(header_content,
                     text="Gestiona las bases de datos del sistema",
                     font=(FUENTE, 13),
                     text_color=COLORES["texto_oscuro"]).pack(anchor="w", pady=(2, 0))

        # --- CONTENIDO PRINCIPAL CON SCROLL ---
        # Frame contenedor principal
        main_container = ctk.CTkFrame(self, fg_color=COLORES["fondo"])
        main_container.pack(fill="both", expand=True, padx=0, pady=0)

        # Crear scrollable frame
        self.scrollable_frame = ctk.CTkScrollableFrame(
            main_container, 
            fg_color=COLORES["fondo"],
            scrollbar_button_color=COLORES["primario"],
            scrollbar_button_hover_color=self.ajustar_color(COLORES["primario"], -20)
        )
        self.scrollable_frame.pack(fill="both", expand=True, padx=25, pady=20)

        # =========================
        # 📚 TARJETA BASE GENERAL ULTA
        # =========================
        card_base_general = ctk.CTkFrame(self.scrollable_frame, 
                                       fg_color=COLORES["surface"], 
                                       corner_radius=12,
                                       border_width=1,
                                       border_color=COLORES["borde"])
        card_base_general.pack(fill="x", pady=(0, 15))

        card_header_general = ctk.CTkFrame(card_base_general, fg_color="transparent")
        card_header_general.pack(fill="x", padx=20, pady=(15, 10))

        ctk.CTkLabel(card_header_general,
                     text="📚 Base General ULTA Etiquetado",
                     font=(FUENTE, 16, "bold"),
                     text_color=COLORES["texto_oscuro"]).pack(side="left")

        self.status_general = ctk.CTkLabel(card_header_general,
                                         text="⏳ No cargado",
                                         font=(FUENTE, 11),
                                         text_color=COLORES["texto_claro"])
        self.status_general.pack(side="right")

        card_content_general = ctk.CTkFrame(card_base_general, fg_color="transparent")
        card_content_general.pack(fill="x", padx=20, pady=(0, 10))

        ctk.CTkLabel(card_content_general,
                     text="Archivo principal con información de productos y etiquetado",
                     font=(FUENTE, 12),
                     text_color=COLORES["texto_claro"],
                     wraplength=600).pack(anchor="w", pady=(0, 12))

        self.info_general = ctk.CTkLabel(card_content_general,
                                       text="No se ha cargado ningún archivo",
                                       font=(FUENTE, 10),
                                       text_color=COLORES["texto_claro"],
                                       wraplength=600)
        self.info_general.pack(anchor="w", pady=(0, 15))

        btn_frame_general = ctk.CTkFrame(card_content_general, fg_color="transparent")
        btn_frame_general.pack(fill="x")

        self.btn_subir_general = ctk.CTkButton(btn_frame_general,
                                             text="📂 Seleccionar Archivo",
                                             fg_color=COLORES["secundario"],
                                             text_color="white",
                                             font=(FUENTE, 12, "bold"),
                                             height=38,
                                             corner_radius=8,
                                             hover_color=self.ajustar_color(COLORES["secundario"], -20),
                                             command=self.subir_base_general)
        self.btn_subir_general.pack(side="left")

        self.btn_quitar_general = ctk.CTkButton(btn_frame_general,
                                              text="🗑️ Quitar",
                                              fg_color=COLORES["peligro"],
                                              text_color="white",
                                              font=(FUENTE, 12),
                                              height=38,
                                              corner_radius=8,
                                              hover_color=self.ajustar_color(COLORES["peligro"], -20),
                                              command=self.quitar_base_general,
                                              state="disabled")
        self.btn_quitar_general.pack(side="left", padx=(10, 0))

        # =========================
        # 🚚 TARJETA BASE ULTA
        # =========================
        card_base_ulta = ctk.CTkFrame(self.scrollable_frame, 
                                    fg_color=COLORES["surface"], 
                                    corner_radius=12,
                                    border_width=1,
                                    border_color=COLORES["borde"])
        card_base_ulta.pack(fill="x", pady=(0, 15))

        card_header_ulta = ctk.CTkFrame(card_base_ulta, fg_color="transparent")
        card_header_ulta.pack(fill="x", padx=20, pady=(15, 10))

        ctk.CTkLabel(card_header_ulta,
                     text="🚚 Base ULTA",
                     font=(FUENTE, 16, "bold"),
                     text_color=COLORES["texto_oscuro"]).pack(side="left")

        self.status_ulta = ctk.CTkLabel(card_header_ulta,
                                      text="⏳ No cargado",
                                      font=(FUENTE, 11),
                                      text_color=COLORES["texto_claro"])
        self.status_ulta.pack(side="right")

        card_content_ulta = ctk.CTkFrame(card_base_ulta, fg_color="transparent")
        card_content_ulta.pack(fill="x", padx=20, pady=(0, 15))

        ctk.CTkLabel(card_content_ulta,
                     text="Base de datos con información de embarques y relaciones",
                     font=(FUENTE, 12),
                     text_color=COLORES["texto_claro"],
                     wraplength=600).pack(anchor="w", pady=(0, 12))

        self.info_ulta = ctk.CTkLabel(card_content_ulta,
                                    text="No se ha cargado ningún archivo",
                                    font=(FUENTE, 10),
                                    text_color=COLORES["texto_claro"],
                                    wraplength=600)
        self.info_ulta.pack(anchor="w", pady=(0, 15))

        btn_frame_ulta = ctk.CTkFrame(card_content_ulta, fg_color="transparent")
        btn_frame_ulta.pack(fill="x")

        self.btn_subir_ulta = ctk.CTkButton(btn_frame_ulta,
                                          text="📂 Seleccionar Archivo",
                                          fg_color=COLORES["secundario"],
                                          text_color="white",
                                          font=(FUENTE, 12, "bold"),
                                          height=38,
                                          corner_radius=8,
                                          hover_color=self.ajustar_color(COLORES["secundario"], -20),
                                          command=self.subir_base_ulta)
        self.btn_subir_ulta.pack(side="left")

        self.btn_quitar_ulta = ctk.CTkButton(btn_frame_ulta,
                                           text="🗑️ Quitar",
                                           fg_color=COLORES["peligro"],
                                           text_color="white",
                                           font=(FUENTE, 12),
                                           height=38,
                                           corner_radius=8,
                                           hover_color=self.ajustar_color(COLORES["peligro"], -20),
                                           command=self.quitar_base_ulta,
                                           state="disabled")
        self.btn_quitar_ulta.pack(side="left", padx=(10, 0))

        # =========================
        # 🔩 TARJETA BASE NORMA 024 (ETIQUETAS METÁLICAS)
        # =========================
        card_base_norma_024 = ctk.CTkFrame(self.scrollable_frame, 
                                         fg_color=COLORES["surface"], 
                                         corner_radius=12,
                                         border_width=1,
                                         border_color=COLORES["borde"])
        card_base_norma_024.pack(fill="x", pady=(0, 15))

        card_header_norma_024 = ctk.CTkFrame(card_base_norma_024, fg_color="transparent")
        card_header_norma_024.pack(fill="x", padx=20, pady=(15, 10))

        ctk.CTkLabel(card_header_norma_024,
                     text="🔩 Base Norma 024 - Etiquetas Metálicas",
                     font=(FUENTE, 16, "bold"),
                     text_color=COLORES["texto_oscuro"]).pack(side="left")

        self.status_norma_024 = ctk.CTkLabel(card_header_norma_024,
                                           text="⏳ No cargado",
                                           font=(FUENTE, 11),
                                           text_color=COLORES["texto_claro"])
        self.status_norma_024.pack(side="right")

        card_content_norma_024 = ctk.CTkFrame(card_base_norma_024, fg_color="transparent")
        card_content_norma_024.pack(fill="x", padx=20, pady=(0, 15))

        ctk.CTkLabel(card_content_norma_024,
                     text="Base especializada para etiquetas metálicas según Norma 024",
                     font=(FUENTE, 12),
                     text_color=COLORES["texto_claro"],
                     wraplength=600).pack(anchor="w", pady=(0, 12))

        self.info_norma_024 = ctk.CTkLabel(card_content_norma_024,
                                         text="No se ha cargado ningún archivo",
                                         font=(FUENTE, 10),
                                         text_color=COLORES["texto_claro"],
                                         wraplength=600)
        self.info_norma_024.pack(anchor="w", pady=(0, 15))

        btn_frame_norma_024 = ctk.CTkFrame(card_content_norma_024, fg_color="transparent")
        btn_frame_norma_024.pack(fill="x")

        self.btn_subir_norma_024 = ctk.CTkButton(btn_frame_norma_024,
                                               text="📂 Seleccionar Archivo",
                                               fg_color=COLORES["secundario"],
                                               text_color="white",
                                               font=(FUENTE, 12, "bold"),
                                               height=38,
                                               corner_radius=8,
                                               hover_color=self.ajustar_color(COLORES["secundario"], -20),
                                               command=self.subir_base_norma_024)
        self.btn_subir_norma_024.pack(side="left")

        self.btn_quitar_norma_024 = ctk.CTkButton(btn_frame_norma_024,
                                                text="🗑️ Quitar",
                                                fg_color=COLORES["peligro"],
                                                text_color="white",
                                                font=(FUENTE, 12),
                                                height=38,
                                                corner_radius=8,
                                                hover_color=self.ajustar_color(COLORES["peligro"], -20),
                                                command=self.quitar_base_norma_024,
                                                state="disabled")
        self.btn_quitar_norma_024.pack(side="left", padx=(10, 0))

        # --- FOOTER MEJORADO ---
        footer = ctk.CTkFrame(self, fg_color=COLORES["surface"], height=70, corner_radius=0)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)

        footer_content = ctk.CTkFrame(footer, fg_color="transparent")
        footer_content.pack(expand=True, fill="both", padx=25, pady=15)

        self.estado_config = ctk.CTkLabel(footer_content,
                                        text="⚠️ Configuración incompleta",
                                        font=(FUENTE, 12, "bold"),
                                        text_color=COLORES["peligro"])
        self.estado_config.pack(side="left")

        btn_frame_footer = ctk.CTkFrame(footer_content, fg_color="transparent")
        btn_frame_footer.pack(side="right")

        self.btn_guardar = ctk.CTkButton(btn_frame_footer,
                                       text="✅ Guardar Configuración",
                                       fg_color=COLORES["exito"],
                                       text_color="white",
                                       font=(FUENTE, 12, "bold"),
                                       width=180,
                                       height=35,
                                       corner_radius=8,
                                       hover_color=self.ajustar_color(COLORES["exito"], -20),
                                       command=self.guardar_configuracion,
                                       state="disabled")
        self.btn_guardar.pack(side="right")

    def ajustar_color(self, color, cantidad):
        r = int(color[1:3], 16)
        g = int(color[3:5], 16)
        b = int(color[5:7], 16)
        r = max(0, min(255, r + cantidad))
        g = max(0, min(255, g + cantidad))
        b = max(0, min(255, b + cantidad))
        return f"#{r:02x}{g:02x}{b:02x}"

    def cargar_configuracion(self):
        """Carga el config.json si existe, si no devuelve estructura vacía"""
        if os.path.exists(self.config_file):
            with open(self.config_file, "r", encoding="utf-8") as f:
                return json.load(f)
        return {
            "base_general_path": None, 
            "base_ulta_path": None,
            "base_norma_024_path": None,
            "base_general_excel": None,
            "base_ulta_excel": None,
            "base_norma_024_excel": None
        }
    
    def guardar_en_json(self):
        """Guarda rutas actuales en config.json"""
        with open(self.config_file, "w", encoding="utf-8") as f:
            json.dump({
                "base_general_path": self.base_general_path,
                "base_ulta_path": self.base_ulta_path,
                "base_norma_024_path": self.base_norma_024_path,
                "base_general_excel": getattr(self,"base_general_excel", None),
                "base_ulta_excel": getattr(self, "base_ulta_excel", None),
                "base_norma_024_excel": getattr(self, "base_norma_024_excel", None)
            }, f, indent=4, ensure_ascii=False)
 
    def actualizar_estados(self):
        # Base General
        if self.base_general_path and os.path.exists(self.base_general_path):
            nombre = os.path.basename(self.base_general_path)
            self.status_general.configure(text="✅ Cargado", text_color=COLORES["exito"])
            self.info_general.configure(text=f"Archivo: {nombre}\nUbicación: {self.base_general_path}")
            self.btn_subir_general.configure(text="🔄 Cambiar Archivo")
            self.btn_quitar_general.configure(state="normal")
        else:
            self.status_general.configure(text="⏳ No cargado", text_color=COLORES["texto_claro"])
            self.info_general.configure(text="No se ha cargado ningún archivo")
            self.btn_subir_general.configure(text="📂 Seleccionar Archivo")
            self.btn_quitar_general.configure(state="disabled")

        # Base ULTA
        if self.base_ulta_path and os.path.exists(self.base_ulta_path):
            nombre = os.path.basename(self.base_ulta_path)
            self.status_ulta.configure(text="✅ Cargado", text_color=COLORES["exito"])
            self.info_ulta.configure(text=f"Archivo: {nombre}\nUbicación: {self.base_ulta_path}")
            self.btn_subir_ulta.configure(text="🔄 Cambiar Archivo")
            self.btn_quitar_ulta.configure(state="normal")
        else:
            self.status_ulta.configure(text="⏳ No cargado", text_color=COLORES["texto_claro"])
            self.info_ulta.configure(text="No se ha cargado ningún archivo")
            self.btn_subir_ulta.configure(text="📂 Seleccionar Archivo")
            self.btn_quitar_ulta.configure(state="disabled")

        # Base Norma 024
        if self.base_norma_024_path and os.path.exists(self.base_norma_024_path):
            nombre = os.path.basename(self.base_norma_024_path)
            self.status_norma_024.configure(text="✅ Cargado", text_color=COLORES["exito"])
            self.info_norma_024.configure(text=f"Archivo: {nombre}\nUbicación: {self.base_norma_024_path}")
            self.btn_subir_norma_024.configure(text="🔄 Cambiar Archivo")
            self.btn_quitar_norma_024.configure(state="normal")
        else:
            self.status_norma_024.configure(text="⏳ No cargado", text_color=COLORES["texto_claro"])
            self.info_norma_024.configure(text="No se ha cargado ningún archivo")
            self.btn_subir_norma_024.configure(text="📂 Seleccionar Archivo")
            self.btn_quitar_norma_024.configure(state="disabled")

        # Estado global (solo requiere las bases principales)
        if (self.base_general_path and os.path.exists(self.base_general_path) and 
            self.base_ulta_path and os.path.exists(self.base_ulta_path)):
            self.estado_config.configure(text="✅ Configuración completa", text_color=COLORES["exito"])
            self.btn_guardar.configure(state="normal")
        else:
            self.estado_config.configure(text="⚠️ Configuración incompleta", text_color=COLORES["peligro"])
            self.btn_guardar.configure(state="disabled")

    def subir_base_general(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar BASE GENERAL ULTA ETIQUETADO",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        if ruta:
            try:
                # Convertir a JSON directamente dentro de la carpeta data
                destino_json = os.path.join(self.data_dir, "BASE_GENERAL_ULTA_ETIQUETADO.json")
                convertir_a_json(archivo_excel=ruta, nombre_json=destino_json, persist=True)

                self.base_general_path = destino_json
                self.base_general_excel = ruta
                self.config_data["base_general_excel"] = ruta
                self.guardar_en_json()

                messagebox.showinfo("✅ Éxito", "Se cargó y convirtió correctamente la BASE GENERAL ULTA ETIQUETADO.")
                self.actualizar_estados()
            except Exception as e:
                messagebox.showerror("❌ Error", f"No se pudo cargar el archivo:\n{e}")

    def subir_base_ulta(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar BASE ULTA",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        if ruta:
            try:
                # Convertir a JSON directamente dentro de la carpeta data
                destino_json = os.path.join(self.data_dir, "BASE_ULTA.json")
                convertir_a_json(archivo_excel=ruta, nombre_json=destino_json, persist=True)

                self.base_ulta_path = destino_json
                self.base_ulta_excel = ruta
                self.config_data["base_ulta_excel"] = ruta
                self.guardar_en_json()

                messagebox.showinfo("✅ Éxito", "Se cargó y convirtió correctamente la BASE ULTA.")
                self.actualizar_estados()
            except Exception as e:
                messagebox.showerror("❌ Error", f"No se pudo cargar el archivo:\n{e}")

    def subir_base_norma_024(self):
        """Subir base de Norma 024 para etiquetas metálicas"""
        ruta = filedialog.askopenfilename(
            title="Seleccionar BASE NORMA 024 - Etiquetas Metálicas",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        if ruta:
            try:
                # Convertir a JSON directamente dentro de la carpeta data
                destino_json = os.path.join(self.data_dir, "BASE_NORMA_024.json")
                convertir_a_json(archivo_excel=ruta, nombre_json=destino_json, persist=True)

                self.base_norma_024_path = destino_json
                self.base_norma_024_excel = ruta
                self.config_data["base_norma_024_excel"] = ruta
                self.guardar_en_json()

                messagebox.showinfo("✅ Éxito", "Se cargó y convirtió correctamente la BASE NORMA 024.")
                self.actualizar_estados()
            except Exception as e:
                messagebox.showerror("❌ Error", f"No se pudo cargar el archivo:\n{e}")

    def quitar_base_general(self):
        if messagebox.askyesno("Confirmar", "¿Estás seguro de que quieres quitar la Base General ULTA?"):
            if self.base_general_path and os.path.exists(self.base_general_path):
                os.remove(self.base_general_path)
            self.base_general_path = None
            self.guardar_en_json()
            self.actualizar_estados()

    def quitar_base_ulta(self):
        if messagebox.askyesno("Confirmar", "¿Estás seguro de que quieres quitar la Base ULTA?"):
            if self.base_ulta_path and os.path.exists(self.base_ulta_path):
                os.remove(self.base_ulta_path)
            self.base_ulta_path = None
            self.guardar_en_json()
            self.actualizar_estados()

    def quitar_base_norma_024(self):
        """Quitar base de Norma 024"""
        if messagebox.askyesno("Confirmar", "¿Estás seguro de que quieres quitar la Base Norma 024?"):
            if self.base_norma_024_path and os.path.exists(self.base_norma_024_path):
                os.remove(self.base_norma_024_path)
            self.base_norma_024_path = None
            self.guardar_en_json()
            self.actualizar_estados()

    def guardar_configuracion(self):
        if not self.base_general_path or not self.base_ulta_path:
            messagebox.showwarning("Advertencia", "⚠️ Debes cargar ambas bases principales antes de continuar.")
            return

        self.guardar_en_json()
        messagebox.showinfo("✅ Configuración", "Las bases se guardaron correctamente.")
        self.destroy()
        if not self.base_general_path or not self.base_ulta_path:
            messagebox.showwarning("Advertencia", "⚠️ Debes cargar ambas bases principales antes de continuar.")
            return

        self.guardar_en_json()
        messagebox.showinfo("✅ Configuración", "Las bases se guardaron correctamente.")
        self.destroy()