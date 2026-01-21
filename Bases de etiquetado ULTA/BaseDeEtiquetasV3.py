import customtkinter as ctk
from tkinter import messagebox
from Configuracion import ConfiguracionWindow
from Dashboard import VentanaDashboard
import pandas as pd
import os
import json

# ─── Paleta de colores mejorada con mejor contraste ─────────────────────
COLORES = {
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
    "hover_primario": "#ECD925"   # Hover para el amarillo
}

FUENTE_PRINCIPAL = "Inter"
FUENTE_SECUNDARIA = "Inter"

# ─── Ventana principal ───────────────────────────────────────
class BasePrincipal(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configuración de ventana
        self.title("🏷️ Sistema de Etiquetado")
        self.geometry("800x550")
        self.minsize(800, 550)
        self.configure(fg_color=COLORES["fondo"])
        self.center_window()

        self.crear_interfaz()

        # Atributos que serán actualizados por la ventana de configuración
        self.base_general = None
        self.layout = None

    def center_window(self):
        """Centra la ventana en la pantalla"""
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        ww, wh = 800, 550
        x = (sw - ww) // 2
        y = (sh - wh) // 2
        self.geometry(f"{ww}x{wh}+{x}+{y}")

    def crear_interfaz(self):
        main_container = ctk.CTkFrame(self, fg_color=COLORES["fondo"])
        main_container.pack(fill="both", expand=True)

        # HEADER con gradiente sutil
        header_frame = ctk.CTkFrame(main_container, fg_color=COLORES["fondo"], height=85, corner_radius=0)
        header_frame.pack(fill="x")
        header_frame.pack_propagate(False)

        # Contenido del header
        header_content = ctk.CTkFrame(header_frame, fg_color="transparent")
        header_content.pack(fill="both", expand=True, padx=40, pady=20)

        titulo = ctk.CTkLabel(header_content, 
                             text="🏷️ Sistema de Etiquetado",
                             font=(FUENTE_PRINCIPAL, 24, "bold"), 
                             text_color=COLORES["header_texto"])
        titulo.pack(anchor="w", pady=(0, 5))

        subtitulo = ctk.CTkLabel(header_content,
                                text="Gestión de bases de etiquetado y análisis de datos",
                                font=(FUENTE_SECUNDARIA, 13), 
                                text_color=COLORES["header_texto"])
        subtitulo.pack(anchor="w")

        # CONTENIDO PRINCIPAL
        content_container = ctk.CTkFrame(main_container, fg_color=COLORES["fondo"])
        content_container.pack(fill="both", expand=True, padx=40, pady=25)

        grid_container = ctk.CTkFrame(content_container, fg_color="transparent")
        grid_container.pack(fill="both", expand=True)

        grid_container.grid_columnconfigure((0, 1), weight=1)
        grid_container.grid_rowconfigure((0, 1), weight=1)

        # Funciones con colores diferenciados
        funciones = [
            {
                "icono": "⚙️", 
                "titulo": "Configuración", 
                "descripcion": "Ajustes del sistema",
                "color": COLORES["secundario"], 
                "comando": self.abrir_configuracion
            },
            {
                "icono": "📊", 
                "titulo": "Dashboard", 
                "descripcion": "Análisis de datos",
                "color": COLORES["secundario"], 
                "comando": self.abrir_dashboard
            },
            {
                "icono": "🧾", 
                "titulo": "Generar Base", 
                "descripcion": "Crear base de etiquetado",
                "color": COLORES["secundario"], 
                "comando": self.generar_base
            },
            {
                "icono": "✏️", 
                "titulo": "Editor", 
                "descripcion": "Editar y gestión de datos",
                "color": COLORES["secundario"], 
                "comando": self.abrir_editor
            },
        ]

        # Tarjetas con mejor diseño
        for i, f in enumerate(funciones):
            row, col = divmod(i, 2)
            
            # Frame principal de la tarjeta
            card = ctk.CTkFrame(grid_container, 
                               fg_color=COLORES["surface"],
                               border_color=COLORES["borde"], 
                               border_width=1, 
                               corner_radius=12)
            card.grid(row=row, column=col, padx=15, pady=10, sticky="nsew")
            card.grid_propagate(False)
            card.configure(height=80)

            # Hacer la tarjeta clickeable
            card.bind("<Enter>", lambda e, c=card: self.on_card_hover(c, True))
            card.bind("<Leave>", lambda e, c=card: self.on_card_hover(c, False))
            card.bind("<Button-1>", lambda e, cmd=f["comando"]: cmd())

            card_content = ctk.CTkFrame(card, fg_color="transparent")
            card_content.pack(fill="both", expand=True, padx=20, pady=15)

            # Icono con color específico
            ctk.CTkLabel(card_content, 
                        text=f["icono"], 
                        font=(FUENTE_PRINCIPAL, 24),
                        text_color=f["color"]).pack(anchor="w")
            
            # Título
            ctk.CTkLabel(card_content, 
                        text=f["titulo"], 
                        font=(FUENTE_PRINCIPAL, 17, "bold"),
                        text_color=COLORES["texto_oscuro"]).pack(anchor="w", pady=(4, 0))
            
            # Descripción
            ctk.CTkLabel(card_content, 
                        text=f["descripcion"], 
                        font=(FUENTE_SECUNDARIA, 12),
                        text_color=COLORES["texto_claro"], 
                        wraplength=360).pack(anchor="w", pady=(2, 6))

            # Botón con color específico y hover
            btn = ctk.CTkButton(card_content, 
                               text="Abrir →", 
                               width=100, 
                               height=32,
                               fg_color=f["color"], 
                               text_color="white", 
                               corner_radius=6,
                               font=(FUENTE_PRINCIPAL, 12, "bold"), 
                               hover_color=self.ajustar_color(f["color"], -25),
                               command=f["comando"])
            btn.pack(anchor="e", pady=(4, 0))

        # FOOTER mejorado
        footer = ctk.CTkFrame(main_container, 
                             fg_color=COLORES["surface"], 
                             height=50, 
                             corner_radius=0,
                             border_width=1,
                             border_color=COLORES["borde"])
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)

        footer_content = ctk.CTkFrame(footer, fg_color="transparent")
        footer_content.pack(fill="both", expand=True, padx=40)

        ctk.CTkLabel(footer_content, 
                    text="Sistema V&C v3.0.0 para ULTA-AXO © 2025", 
                    font=(FUENTE_SECUNDARIA, 11),
                    text_color=COLORES["texto_claro"]).pack(side="left")

        ctk.CTkButton(footer_content, 
                     text="🚪 Salir", 
                     fg_color=COLORES["peligro"], 
                     text_color="white",
                     width=90, 
                     height=32, 
                     corner_radius=6,
                     font=(FUENTE_PRINCIPAL, 12, "bold"),
                     hover_color=self.ajustar_color(COLORES["peligro"], -25),
                     command=self.cerrar_aplicacion).pack(side="right")

    def on_card_hover(self, card, is_hover):
        """Efecto hover para las tarjetas"""
        if is_hover:
            card.configure(border_color=self.ajustar_color(COLORES["borde"], -30))
        else:
            card.configure(border_color=COLORES["borde"])

    def ajustar_color(self, color, cantidad):
        """Ajusta el brillo de un color HEX"""
        try:
            r = int(color[1:3], 16)
            g = int(color[3:5], 16)
            b = int(color[5:7], 16)
            r = max(0, min(255, r + cantidad))
            g = max(0, min(255, g + cantidad))
            b = max(0, min(255, b + cantidad))
            return f"#{r:02x}{g:02x}{b:02x}"
        except:
            return color

    # ─── Funciones de los botones ───────────────────────────
    def abrir_configuracion(self):
        if not hasattr(self, 'config_window') or not self.config_window.winfo_exists():
            self.config_window = ConfiguracionWindow(parent=self)
        else:
            self.config_window.lift()

    def generar_base(self):
        from Armado import generar_base as generar_armado
        generar_armado()

    def abrir_dashboard(self):
        from Dashboard import VentanaDashboard
        base_general = os.path.join("data", "BASE_GENERAL_ULTA_ETIQUETADO.json")
        base_ulta = os.path.join("data", "BASE_ULTA.json")

        if not os.path.exists(base_general) or not os.path.exists(base_ulta):
            messagebox.showwarning("Dashboard", "Faltan archivos JSON de base general o ULTA.")
            return

        dashboard = VentanaDashboard(self)
        dashboard.actualizar_dashboard()
        dashboard.focus()

    def abrir_editor(self):
        from Editor import EditorWindow
        base_general = os.path.join("data", "BASE_GENERAL_ULTA_ETIQUETADO.json")

        if not os.path.exists(base_general):
            messagebox.showwarning("Editor", "Falta el archivo JSON de base general.")
            return

        editor = EditorWindow(self)
        editor.focus()

    def cerrar_aplicacion(self):
        if messagebox.askyesno("Confirmar salida", "¿Deseas salir del sistema?", icon="warning"):
            self.destroy()

if __name__ == "__main__":
    ctk.set_appearance_mode("light")
    app = BasePrincipal()
    app.mainloop()
