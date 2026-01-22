import os, sys
import json
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

# Carpeta de configuración persistente
CONFIG_DIR = os.path.join(os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.abspath("."), "datos")
os.makedirs(CONFIG_DIR, exist_ok=True)

CONFIG_FILE = os.path.join(CONFIG_DIR, "config.json")

# -------------------------------
# FUNCIONES DE RUTAS
# -------------------------------
def recurso_path(ruta_relativa):
    """
    Devuelve la ruta correcta de recursos.
    - Para datos (JSON, Excel) => carpeta 'datos' junto al exe.
    - Para recursos estáticos (iconos, etc.) => usar sys._MEIPASS.
    """
    if getattr(sys, "frozen", False):
        # Cuando está en .exe
        base_path = os.path.dirname(sys.executable)  # Carpeta del .exe
    else:
        # Cuando está en .py
        base_path = os.path.abspath(".")
    return os.path.join(base_path, ruta_relativa)

# -------------------------------
# CONFIGURACIÓN
# -------------------------------
def cargar_configuracion():
    """Carga la configuración desde el archivo JSON"""
    config_default = {"rutas": {"base_general": "", "codigos_cumple": ""}}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
            # Si falta la clave 'rutas', restaurar estructura
            if not isinstance(config, dict) or "rutas" not in config:
                config = config_default
                with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                    json.dump(config, f, indent=4)
            # Si existe 'historial' en rutas, eliminarlo
            if "rutas" in config and "historial" in config["rutas"]:
                del config["rutas"]["historial"]
                guardar_configuracion(config)
            return config
        except Exception:
            # Archivo corrupto o vacío, restaurar estructura
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config_default, f, indent=4)
            return config_default
    else:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config_default, f, indent=4)
        return config_default

def guardar_configuracion(config):
    """Guarda la configuración en el archivo JSON"""
    try:
        config["rutas"] = {
            "base_general": config["rutas"].get("base_general", ""),
            "codigos_cumple": config["rutas"].get("codigos_cumple", "")
        }
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4)
        return True
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo guardar configuración:\n{e}")
        return False

# -------------------------------
# INTERFAZ DE CONFIGURACIÓN
# -------------------------------
def configurar_rutas(parent=None):
    try:
        config = cargar_configuracion()
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo cargar configuración:\n{e}")
        return {}

    # Si no se pasa parent, crear un root oculto
    if parent is None:
        root = tk.Tk()
        root.withdraw()
        ventana = tk.Toplevel(root)
    else:
        ventana = tk.Toplevel(parent)

    ventana.title("⚙ Configuración de Rutas")
    ventana.geometry("600x600")
    ventana.configure(bg="#FFFFFF")
    ventana.resizable(False, False)
    ventana.grab_set()

    # --- Funciones internas ---
    import shutil
    def seleccionar_archivo(tipo, label_widget, button_widget):
        archivo = filedialog.askopenfilename(
            title=f"Seleccionar {tipo.replace('_', ' ').title()}",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        if archivo:
            # Nombre destino estándar en carpeta Datos
            if tipo == "codigos_cumple":
                destino = os.path.abspath(os.path.join(CONFIG_DIR, "codigos_cumple.xlsx"))
            elif tipo == "base_general":
                destino = os.path.abspath(os.path.join(CONFIG_DIR, "base_general.xlsx"))
            else:
                destino = os.path.abspath(os.path.join(CONFIG_DIR, os.path.basename(archivo)))

            try:
                shutil.copy2(archivo, destino)
                df = pd.read_excel(destino)
                # Convertir a JSON en carpeta Datos
                nombre_json = os.path.splitext(os.path.basename(destino))[0] + ".json"
                json_path = os.path.abspath(os.path.join(CONFIG_DIR, nombre_json))
                df.to_json(json_path, orient="records", force_ascii=False, indent=4)
                messagebox.showinfo("✅ Conversión exitosa", f"Archivo copiado y convertido a JSON:\n{os.path.basename(json_path)}")
                # Guardar ruta absoluta en config
                config["rutas"][tipo] = destino
            except Exception as e:
                messagebox.showwarning("⚠ Advertencia", f"No se pudo copiar o convertir a JSON:\n{str(e)}")

            # Actualizar interfaz
            nombre_archivo = os.path.basename(destino)
            if len(nombre_archivo) > 30:
                nombre_archivo = nombre_archivo[:27] + "..."
            label_widget.config(text=nombre_archivo, fg="#282828")
            button_widget.config(text="🔄 Cambiar", bg="#4B4B4B", fg="#FFFFFF")
            guardar_configuracion(config)
            actualizar_estado()

    def limpiar_configuracion():
        if messagebox.askyesno("🗑️ Limpiar Configuración", "¿Estás seguro de que quieres limpiar toda la configuración?"):
            config["rutas"] = {"base_general": "", "codigos_cumple": ""}
            lbl_codigos.config(text="No seleccionado", fg="#4B4B4B")
            lbl_base.config(text="No seleccionado", fg="#4B4B4B")
            btn_codigos.config(text="📂 Seleccionar", bg="#ECD925", fg="#282828")
            btn_base.config(text="📂 Seleccionar", bg="#ECD925", fg="#282828")
            guardar_configuracion(config)
            actualizar_estado()
            messagebox.showinfo("✅ Configuración limpiada", "Se han borrado todas las rutas seleccionadas.")

    def actualizar_estado():
        rutas_configuradas = sum(1 for ruta in config["rutas"].values() if ruta)
        if rutas_configuradas == 2:
            lbl_estado.config(text="✅ Configuración completa - Listo para cerrar", fg="#282828")
        elif rutas_configuradas == 1:
            lbl_estado.config(text="⚠️  Falta 1 archivo por configurar", fg="#4B4B4B")
        else:
            lbl_estado.config(text="❌ No hay archivos configurados", fg="#4B4B4B")


    # Frame principal
    main_frame = tk.Frame(ventana, bg="#FFFFFF", padx=40, pady=30)
    main_frame.pack(fill="both", expand=True)

    # Header
    header_frame = tk.Frame(main_frame, bg="#FFFFFF")
    header_frame.pack(fill="x", pady=(0, 30))
    tk.Label(header_frame, text="⚙ CONFIGURACIÓN DE RUTAS", font=("Inter", 20, "bold"), bg="#FFFFFF", fg="#282828").pack()
    tk.Label(header_frame, text="Selecciona los archivos necesarios para iniciar el sistema", font=("Inter", 10), bg="#FFFFFF", fg="#4B4B4B").pack(pady=(5, 0))

    # Sección de Códigos de Cumplimiento
    frame_codigos = tk.Frame(main_frame, bg="#F8F9FA", relief="flat", padx=20, pady=20)
    frame_codigos.pack(fill="x", pady=(0, 15))
    tk.Label(frame_codigos, text="📋 CÓDIGOS DE CUMPLIMIENTO", font=("Inter", 12, "bold"), bg="#F8F9FA", fg="#282828").pack(anchor="w")
    tk.Label(frame_codigos, text="Archivo Excel con los códigos y criterios de evaluación", font=("Inter", 9), bg="#F8F9FA", fg="#4B4B4B").pack(anchor="w", pady=(2, 15))
    file_frame = tk.Frame(frame_codigos, bg="#F8F9FA")
    file_frame.pack(fill="x")
    ruta_actual = config["rutas"].get("codigos_cumple", "")
    texto_inicial = os.path.basename(ruta_actual) if ruta_actual else "No seleccionado"
    color_inicial = "#282828" if ruta_actual else "#4B4B4B"
    lbl_codigos = tk.Label(file_frame, text=texto_inicial, font=("Inter", 10), bg="#F8F9FA", fg=color_inicial, wraplength=400, justify="left")
    lbl_codigos.pack(side="left", padx=(0, 10))
    btn_color = "#4B4B4B" if ruta_actual else "#ECD925"
    btn_text = "🔄 Cambiar" if ruta_actual else "📂 Seleccionar"
    btn_fg = "#FFFFFF" if ruta_actual else "#282828"
    btn_codigos = tk.Button(file_frame, text=btn_text, font=("Inter", 10, "bold"), bg=btn_color, fg=btn_fg, relief="flat", padx=20, pady=5, command=lambda: seleccionar_archivo("codigos_cumple", lbl_codigos, btn_codigos))
    btn_codigos.pack(side="right")

    # Sección de Base General
    frame_base = tk.Frame(main_frame, bg="#F8F9FA", relief="flat", padx=20, pady=20)
    frame_base.pack(fill="x", pady=(0, 25))
    tk.Label(frame_base, text="📊 BASE GENERAL DE DATOS", font=("Inter", 12, "bold"), bg="#F8F9FA", fg="#282828").pack(anchor="w")
    tk.Label(frame_base, text="Archivo Excel principal con los datos del sistema", font=("Inter", 9), bg="#F8F9FA", fg="#4B4B4B").pack(anchor="w", pady=(2, 15))
    file_frame2 = tk.Frame(frame_base, bg="#F8F9FA")
    file_frame2.pack(fill="x")
    ruta_actual_base = config["rutas"].get("base_general", "")
    texto_inicial_base = os.path.basename(ruta_actual_base) if ruta_actual_base else "No seleccionado"
    color_inicial_base = "#282828" if ruta_actual_base else "#4B4B4B"
    lbl_base = tk.Label(file_frame2, text=texto_inicial_base, font=("Inter", 10), bg="#F8F9FA", fg=color_inicial_base, wraplength=400, justify="left")
    lbl_base.pack(side="left", padx=(0, 10))
    btn_color_base = "#4B4B4B" if ruta_actual_base else "#ECD925"
    btn_text_base = "🔄 Cambiar" if ruta_actual_base else "📂 Seleccionar"
    btn_fg_base = "#FFFFFF" if ruta_actual_base else "#282828"
    btn_base = tk.Button(file_frame2, text=btn_text_base, font=("Inter", 10, "bold"), bg=btn_color_base, fg=btn_fg_base, relief="flat", padx=20, pady=5, command=lambda: seleccionar_archivo("base_general", lbl_base, btn_base))
    btn_base.pack(side="right")

    # Estado
    estado_frame = tk.Frame(main_frame, bg="#FFFFFF")
    estado_frame.pack(fill="x", pady=(0, 20))
    lbl_estado = tk.Label(estado_frame, text="", font=("Inter", 10, "bold"), bg="#FFFFFF")
    lbl_estado.pack()

    # Botones de acción
    action_frame = tk.Frame(main_frame, bg="#FFFFFF")
    action_frame.pack(fill="x", pady=(0, 0))
    btn_frame = tk.Frame(action_frame, bg="#FFFFFF")
    btn_frame.pack()
    btn_limpiar = tk.Button(btn_frame, text="🗑️ LIMPIAR CONFIGURACIÓN", font=("Inter", 10, "bold"), bg="#4B4B4B", fg="#FFFFFF", relief="flat", padx=25, pady=10, command=limpiar_configuracion)
    btn_limpiar.pack(side="left", padx=10)
    btn_cerrar = tk.Button(btn_frame, text="❌ CERRAR", font=("Inter", 10, "bold"), bg="#4B4B4B", fg="#FFFFFF", relief="flat", padx=25, pady=10, command=ventana.destroy)
    btn_cerrar.pack(side="left", padx=10)

    # Actualizar estado inicial
    actualizar_estado()

    ventana.wait_window()
    return config["rutas"]

# --- PROGRAMA PRINCIPAL ---
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    configurar_rutas()
    root.mainloop()
