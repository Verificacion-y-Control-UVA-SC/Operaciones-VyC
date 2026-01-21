import os, re
import json
import pandas as pd
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, filedialog
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import tempfile
from datetime import datetime
from io import BytesIO
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors as rl_colors

COLORES = {
    "primario": "#ECD925",
    "secundario": "#282828",
    "exito": "#008d53",
    "peligro": "#d74a3d",
    "fondo": "#F8F9FA",
    "surface": "#FFFFFF",
    "texto_oscuro": "#282828",
    "accent": "#3498db",
    "warning": "#f39c12"
}

FUENTE = "Inter"

class VentanaDashboard(ctk.CTkToplevel):
    instancia_activa = None  # Para referencia global

    def __init__(self, parent=None):
        super().__init__(parent)
        VentanaDashboard.instancia_activa = self
        self.title("📊 Dashboard")
        self.geometry("1400x800")  # Aumenté el tamaño para acomodar más tarjetas
        self.configure(fg_color=COLORES["fondo"])
        self.center_window()

        # Configurar matplotlib para usar el backend TkAgg
        import matplotlib
        matplotlib.use('TkAgg')
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.crear_interfaz()
        self.actualizar_dashboard()
        self.cargar_historial_bases()

    def center_window(self):
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        ww, wh = 1400, 800
        x = (sw - ww) // 2
        y = (sh - wh) // 2
        self.geometry(f"{ww}x{wh}+{x}+{y}")

    def close(self):
        """Cerrar la ventana de dashboard."""
        try:
            VentanaDashboard.instancia_activa = None
            self.destroy()
        except Exception:
            pass

    def crear_interfaz(self):
        # HEADER
        header = ctk.CTkFrame(self, fg_color=COLORES["fondo"], height=100, corner_radius=0)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)

        header_content = ctk.CTkFrame(header, fg_color="transparent")
        header_content.pack(expand=True, fill="both", padx=30, pady=15)

        # Left labels
        left_labels = ctk.CTkFrame(header_content, fg_color="transparent")
        left_labels.pack(side="left", fill="both", expand=True)
        ctk.CTkLabel(left_labels, text="📊 Dashboard", font=(FUENTE, 22, "bold"), text_color=COLORES["texto_oscuro"]).pack(anchor="w")
        ctk.CTkLabel(left_labels, text="Análisis de datos y etiquetado", font=(FUENTE, 12), text_color=COLORES["texto_oscuro"]).pack(anchor="w", pady=(2,0))

        # Right actions (Cerrar)
        header_actions = ctk.CTkFrame(header_content, fg_color="transparent")
        header_actions.pack(side="right", anchor="e")
        btn_close = ctk.CTkButton(header_actions, text="Cerrar ✖", width=90, height=30, fg_color=COLORES["peligro"], text_color="white", command=self.close)
        btn_close.pack()

        # CONTENIDO PRINCIPAL
        main_content = ctk.CTkFrame(self, fg_color=COLORES["fondo"])
        main_content.grid(row=1, column=0, sticky="nsew", padx=20, pady=20)
        main_content.grid_columnconfigure(0, weight=1)
        main_content.grid_rowconfigure(1, weight=1)

        # TARJETAS - Ahora con 5 columnas
        metrics_frame = ctk.CTkFrame(main_content, fg_color="transparent")
        metrics_frame.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        for i in range(5):  # Cambiado a 5 columnas
            metrics_frame.grid_columnconfigure(i, weight=1, uniform="metrics")

        self.card_total = self.crear_tarjeta_metrica(metrics_frame, "🔸 Total UPCs", "0", COLORES["secundario"])
        self.card_total.grid(row=0, column=0, sticky="nsew", padx=(0,5))

        self.card_categorias = self.crear_tarjeta_metrica(metrics_frame, "📂 Total Categorías", "0", COLORES["secundario"])
        self.card_categorias.grid(row=0, column=1, sticky="nsew", padx=5)

        self.card_normal = self.crear_tarjeta_metrica(metrics_frame, "✅ Normales", "0", COLORES["exito"])
        self.card_normal.grid(row=0, column=2, sticky="nsew", padx=5)

        self.card_especial = self.crear_tarjeta_metrica(metrics_frame, "⚠️ Especiales", "0", COLORES["primario"])
        self.card_especial.grid(row=0, column=3, sticky="nsew", padx=5)

        self.card_metalicas = self.crear_tarjeta_metrica(metrics_frame, "🔩 Metálicas", "0", COLORES["warning"])
        self.card_metalicas.grid(row=0, column=4, sticky="nsew", padx=(5,0))

        # ... (el resto de la interfaz se mantiene igual)
        # CONTENIDO DIVIDIDO - Gráficas y archivos
        content_split = ctk.CTkFrame(main_content, fg_color=COLORES["fondo"])
        content_split.grid(row=1, column=0, sticky="nsew")
        content_split.grid_columnconfigure(0, weight=2)  # Gráficas
        content_split.grid_columnconfigure(1, weight=1)  # Archivos
        content_split.grid_rowconfigure(0, weight=1)

        # --- PANEL IZQUIERDO: GRÁFICAS ---
        charts_container = ctk.CTkFrame(content_split, fg_color=COLORES["fondo"])
        charts_container.grid(row=0, column=0, sticky="nsew", padx=(0,10))
        charts_container.grid_columnconfigure(0, weight=1)
        charts_container.grid_rowconfigure(0, weight=1)  # Fila para gráficas

        # Frame para contener ambas gráficas lado a lado
        graphs_frame = ctk.CTkFrame(charts_container, fg_color=COLORES["fondo"])
        graphs_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        graphs_frame.grid_columnconfigure(0, weight=1)  # Dona
        graphs_frame.grid_columnconfigure(1, weight=1)  # Barras
        graphs_frame.grid_rowconfigure(0, weight=1)

        # Gráfica de dona - Tipos de etiquetas
        dona_container = ctk.CTkFrame(graphs_frame, fg_color=COLORES["surface"], corner_radius=12)
        dona_container.grid(row=0, column=0, sticky="nsew", padx=(0,5))
        dona_container.grid_columnconfigure(0, weight=1)
        dona_container.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(dona_container, text="📊 Tipos de Etiquetas", 
                    font=(FUENTE, 14, "bold"), 
                    text_color=COLORES["texto_oscuro"]).grid(row=0, column=0, sticky="w", padx=15, pady=10)

        self.dona_chart_frame = ctk.CTkFrame(dona_container, fg_color=COLORES["surface"])
        self.dona_chart_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0,10))
        self.dona_chart_frame.grid_columnconfigure(0, weight=1)
        self.dona_chart_frame.grid_rowconfigure(0, weight=1)

        # Gráfica de barras - Medidas
        bars_container = ctk.CTkFrame(graphs_frame, fg_color=COLORES["surface"], corner_radius=12)
        bars_container.grid(row=0, column=1, sticky="nsew", padx=(5,0))
        bars_container.grid_columnconfigure(0, weight=1)
        bars_container.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(bars_container, text="📏 Distribución de Medidas", 
                    font=(FUENTE, 14, "bold"), 
                    text_color=COLORES["texto_oscuro"]).grid(row=0, column=0, sticky="w", padx=15, pady=10)

        self.bars_chart_frame = ctk.CTkFrame(bars_container, fg_color=COLORES["surface"])
        self.bars_chart_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0,10))
        self.bars_chart_frame.grid_columnconfigure(0, weight=1)
        self.bars_chart_frame.grid_rowconfigure(0, weight=1)

        # --- PANEL DERECHO: ARCHIVOS PROCESADOS ---
        files_panel = ctk.CTkFrame(content_split, fg_color=COLORES["surface"], corner_radius=12)
        files_panel.grid(row=0, column=1, sticky="nsew")
        files_panel.grid_columnconfigure(0, weight=1)
        files_panel.grid_rowconfigure(1, weight=1)

        # Título archivos procesados
        ctk.CTkLabel(files_panel, text="📁 Archivos Procesados", 
                    font=(FUENTE, 14, "bold"), 
                    text_color=COLORES["texto_oscuro"]).grid(row=0, column=0, sticky="w", padx=15, pady=(12,8))

        # Frame para la lista de archivos
        lista_container = ctk.CTkFrame(files_panel, fg_color=COLORES["surface"])
        lista_container.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0,10))
        lista_container.grid_columnconfigure(0, weight=1)
        lista_container.grid_rowconfigure(0, weight=1)

        # Text widget para listar archivos
        self.lista_archivos = tk.Text(
            lista_container,
            height=15,
            state="disabled",
            wrap="word",
            bg=COLORES["surface"],
            fg=COLORES["texto_oscuro"],
            font=(FUENTE, 10),
            bd=0,
            highlightthickness=0,
            padx=8,
            pady=8,
            selectbackground=COLORES["primario"]
        )
        
        scrollbar = ctk.CTkScrollbar(lista_container, command=self.lista_archivos.yview)
        self.lista_archivos.configure(yscrollcommand=scrollbar.set)
        
        self.lista_archivos.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # Botones de acción
        buttons_frame = ctk.CTkFrame(files_panel, fg_color="transparent")
        buttons_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(0,10))
        buttons_frame.grid_columnconfigure((0, 1), weight=1)

        btn_export = ctk.CTkButton(buttons_frame, text="📤 PDF", 
                                 command=self.exportar_pdf, 
                                 fg_color=COLORES["secundario"],
                                 hover_color="#4b4b4b", 
                                 text_color="white",
                                 height=32,
                                 font=(FUENTE, 11))
        btn_export.grid(row=0, column=0, sticky="ew", padx=(0,4))

        btn_borrar_sel = ctk.CTkButton(buttons_frame, text="🗑️ Borrar", 
                                     command=self.borrar_seleccionado, 
                                     fg_color=COLORES["peligro"], 
                                     hover_color="#d74a3d",
                                     text_color="white",
                                     height=32,
                                     font=(FUENTE, 11))
        btn_borrar_sel.grid(row=0, column=1, sticky="ew", padx=(4,0))

        btn_limpiar = ctk.CTkButton(buttons_frame, text="🧹 Limpiar", 
                                  command=self.limpiar_lista, 
                                  fg_color=COLORES["secundario"], 
                                  hover_color="#4b4b4b", 
                                  text_color="white",
                                  height=32,
                                  font=(FUENTE, 11))
        btn_limpiar.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6,0))

        # Estado interno de archivos procesados
        self.archivos_procesados = []

    def crear_tarjeta_metrica(self, parent, titulo, valor, color):
        """Crea una tarjeta de métrica con diseño consistente"""
        card = ctk.CTkFrame(parent, fg_color=color, corner_radius=12, height=100)
        card.grid_propagate(False)

        content = ctk.CTkFrame(card, fg_color="transparent")
        content.pack(expand=True, fill="both", padx=20, pady=15)

        # Título
        ctk.CTkLabel(content, text=titulo, 
                    font=(FUENTE, 12, "bold"), 
                    text_color="white").pack(anchor="w")
        
        # Valor
        valor_label = ctk.CTkLabel(content, text=valor, 
                                 font=(FUENTE, 20, "bold"), 
                                 text_color="white")
        valor_label.pack(anchor="w", pady=(8,0))

        # Guardar referencia para actualizaciones
        card.valor_label = valor_label

        return card

    def actualizar_tarjeta_metrica(self, card, nuevo_valor):
        """Actualiza el valor de una tarjeta métrica"""
        if hasattr(card, "valor_label"):
            card.valor_label.configure(text=str(nuevo_valor))

    def actualizar_dashboard(self):
        """Actualiza todos los datos del dashboard"""
        try:
            print("🔄 Actualizando dashboard...")
            
            # 🔸 CARGAR BASE DE DATOS GENERAL
            base_general_path = os.path.join("data", "BASE_GENERAL_ULTA_ETIQUETADO.json")
            base_norma_path = os.path.join("data", "BASE_NORMA_024.json")
            
            print(f"📁 Buscando archivos en:")
            print(f"  - General: {base_general_path}")
            print(f"  - Norma 024: {base_norma_path}")

            df_general = pd.DataFrame()
            df_norma = pd.DataFrame()

            # Cargar Base General
            if os.path.exists(base_general_path):
                print("✅ Archivo general encontrado, cargando datos...")
                df_general = self._cargar_base_datos(base_general_path)
                if df_general.empty:
                    print("⚠️ DataFrame general vacío después de la carga")
            else:
                print(f"❌ Archivo general no encontrado: {base_general_path}")

            # Cargar Base Norma 024
            if os.path.exists(base_norma_path):
                print("✅ Archivo norma 024 encontrado, cargando datos...")
                df_norma = self._cargar_base_datos(base_norma_path)
                if df_norma.empty:
                    print("⚠️ DataFrame norma 024 vacío después de la carga")
            else:
                print(f"❌ Archivo norma 024 no encontrado: {base_norma_path}")

            if df_general.empty and df_norma.empty:
                print("❌ No hay datos en ninguna base")
                self.mostrar_estado_vacio()
                return

            # 🔸 CALCULAR MÉTRICAS BÁSICAS DE BASE GENERAL
            if not df_general.empty:
                print(f"📈 DataFrame general cargado: {len(df_general)} filas")
                print(f"📋 Columnas general: {list(df_general.columns)}")

                # Buscar columnas en base general
                upc_col_general = self._buscar_columna(df_general, ['UPC', 'Upc', 'Upc/Ean'])
                cat_col_general = self._buscar_columna(df_general, ['CATEGORIA', 'CATEGORÍA', 'Categoria'])
                medidas_col_general = self._buscar_columna(df_general, ['MEDIDAS', 'Medidas'])

                print(f"🔍 Columnas general identificadas: UPC={upc_col_general}, Categoría={cat_col_general}, Medidas={medidas_col_general}")

                # Totales básicos de base general
                self.total_upcs = df_general[upc_col_general].nunique() if upc_col_general in df_general.columns else 0
                self.total_categorias = df_general[cat_col_general].nunique() if cat_col_general in df_general.columns else 0

                # Normales y Especiales de base general
                if medidas_col_general in df_general.columns:
                    # Contar especiales
                    self.especiales = df_general[medidas_col_general].astype(str).str.contains(
                        "REQUIERE ETIQUETADO ESPECIAL|NO IMPRIMIR HASTA TENER VISTO BUENO DE V&C", 
                        case=False, na=False
                    ).sum()
                    self.normales = len(df_general) - self.especiales
                    print(f"📊 Conteo general - especiales: {self.especiales}, normales: {self.normales}")
                else:
                    self.normales = 0
                    self.especiales = 0
                    print("⚠️ No se encontró columna de medidas en base general")
            else:
                self.total_upcs = 0
                self.total_categorias = 0
                self.normales = 0
                self.especiales = 0
                print("⚠️ No hay datos en base general")

            # 🔸 CALCULAR MÉTRICAS DE BASE NORMA 024 (Metálicas)
            if not df_norma.empty:
                print(f"📈 DataFrame norma 024 cargado: {len(df_norma)} filas")
                print(f"📋 Columnas norma 024: {list(df_norma.columns)}")
                
                # Contar total de registros en base norma 024 (etiquetas metálicas)
                self.metalicas = len(df_norma)
                print(f"🔩 Total etiquetas metálicas: {self.metalicas}")
            else:
                self.metalicas = 0
                print("⚠️ No hay datos en base norma 024")

            # 🔸 ACTUALIZAR INTERFAZ
            self.actualizar_tarjeta_metrica(self.card_total, self.total_upcs)
            self.actualizar_tarjeta_metrica(self.card_categorias, self.total_categorias)
            self.actualizar_tarjeta_metrica(self.card_normal, self.normales)
            self.actualizar_tarjeta_metrica(self.card_especial, self.especiales)
            self.actualizar_tarjeta_metrica(self.card_metalicas, self.metalicas)

            # 🔸 ACTUALIZAR GRÁFICAS (solo con datos de base general)
            print("🎨 Actualizando gráficas...")
            self.actualizar_grafica_dona()
            
            # Usar datos de medidas de base general para la gráfica de barras
            if not df_general.empty:
                medidas_col = self._buscar_columna(df_general, ['MEDIDAS', 'Medidas'])
                self.actualizar_grafica_barras_compacta(df_general, medidas_col)
            else:
                # Si no hay base general, intentar con base norma 024
                if not df_norma.empty:
                    medidas_col = self._buscar_columna(df_norma, ['MEDIDAS', 'Medidas'])
                    self.actualizar_grafica_barras_compacta(df_norma, medidas_col)
                else:
                    # Mostrar estado vacío en gráfica de barras
                    self._mostrar_grafica_vacia(self.bars_chart_frame)
            
            print("✅ Dashboard actualizado correctamente")

        except Exception as e:
            print(f"❌ Error actualizando dashboard: {e}")
            import traceback
            traceback.print_exc()
            self.mostrar_estado_vacio()

    def _cargar_base_datos(self, ruta_archivo):
        """Carga un archivo JSON y retorna un DataFrame"""
        try:
            with open(ruta_archivo, "r", encoding="utf-8") as f:
                data = json.load(f)
                print(f"📊 Datos JSON cargados desde {os.path.basename(ruta_archivo)}: {type(data)}")

            # Manejar diferentes estructuras de JSON
            if isinstance(data, dict):
                print("📂 JSON es diccionario, buscando lista...")
                # Buscar la lista principal
                for key, value in data.items():
                    if isinstance(value, list):
                        df = pd.DataFrame(value)
                        print(f"✅ Lista encontrada en clave '{key}': {len(value)} elementos")
                        return df
                else:
                    df = pd.DataFrame([data])
                    print("⚠️ No se encontró lista, usando diccionario directo")
                    return df
            elif isinstance(data, list):
                df = pd.DataFrame(data)
                print(f"✅ JSON es lista: {len(data)} elementos")
                return df
            else:
                print("❌ Formato JSON no reconocido")
                return pd.DataFrame()
                
        except json.JSONDecodeError as e:
            print(f"❌ Error leyendo JSON {ruta_archivo}: {e}")
            return pd.DataFrame()
        except Exception as e:
            print(f"❌ Error cargando {ruta_archivo}: {e}")
            return pd.DataFrame()

    def _buscar_columna(self, df, posibles_nombres):
        """Busca una columna en el DataFrame por posibles nombres"""
        for nombre in posibles_nombres:
            if nombre in df.columns:
                return nombre
        # Si no encuentra exacto, buscar case insensitive
        df_cols_lower = [col.lower() for col in df.columns]
        for nombre in posibles_nombres:
            if nombre.lower() in df_cols_lower:
                idx = df_cols_lower.index(nombre.lower())
                return df.columns[idx]
        return None

    def _mostrar_grafica_vacia(self, frame):
        """Muestra un mensaje de gráfica vacía"""
        for widget in frame.winfo_children():
            widget.destroy()
        no_data_label = ctk.CTkLabel(frame, 
                                   text="📊 No hay datos\npara mostrar",
                                   font=(FUENTE, 12),
                                   text_color=COLORES["texto_oscuro"])
        no_data_label.pack(expand=True, fill="both")

    def _normalizar_medida(self, medida):
        """Normaliza una medida para agrupar variaciones"""
        try:
            if pd.isna(medida) or medida == 'nan' or medida == '':
                return None
                
            medida_str = str(medida).strip().upper()
            
            # Casos especiales
            if "REQUIERE ETIQUETADO ESPECIAL" in medida_str:
                return "REQUIERE ETIQUETADO ESPECIAL"
            elif "NO IMPRIMIR HASTA TENER VISTO BUENO" in medida_str or "VISTO BUENO" in medida_str:
                return "NO IMPRIMIR - VISTO BUENO V&C"
            elif "ESPECIAL" in medida_str:
                return "ETIQUETADO ESPECIAL"
            
            # Normalizar medidas numéricas
            medida_str = re.sub(r'\s+', ' ', medida_str)
            medida_str = re.sub(r'(\d+)\s*MM\s*[X×]\s*(\d+)\s*MM', r'\1 mm × \2 mm', medida_str, flags=re.IGNORECASE)
            medida_str = re.sub(r'(\d+)\s*[X×]\s*(\d+)\s*MM', r'\1 × \2 mm', medida_str, flags=re.IGNORECASE)
            medida_str = re.sub(r'(\d+)\s*[X×]\s*(\d+)', r'\1 × \2', medida_str)
            
            return medida_str
        except Exception as e:
            print(f"⚠️ Error normalizando medida '{medida}': {e}")
            return None

    def actualizar_grafica_dona(self):
        """Actualiza la gráfica de dona para tipos de etiquetas"""
        print("🎨 Dibujando gráfica de dona...")
        
        # Limpiar gráfica anterior
        if hasattr(self, 'dona_chart_canvas') and self.dona_chart_canvas:
            try:
                self.dona_chart_canvas.get_tk_widget().destroy()
            except:
                pass

        # Limpiar frame
        for widget in self.dona_chart_frame.winfo_children():
            widget.destroy()

        # Verificar si hay datos para mostrar
        if self.normales == 0 and self.especiales == 0:
            print("⚠️ No hay datos para la gráfica de dona")
            no_data_label = ctk.CTkLabel(self.dona_chart_frame, 
                                       text="📊 No hay datos\npara mostrar",
                                       font=(FUENTE, 12),
                                       text_color=COLORES["texto_oscuro"])
            no_data_label.pack(expand=True, fill="both")
            return

        try:
            # Crear figura
            fig, ax = plt.subplots(figsize=(5, 4), dpi=80)  # Tamaño ajustado
            fig.patch.set_facecolor(COLORES["surface"])
            ax.set_facecolor(COLORES["surface"])

            # Datos para la gráfica
            labels = ["Normales", "Especiales"]
            sizes = [self.normales, self.especiales]
            colors = [COLORES["exito"], COLORES["warning"]]

            # Crear gráfica de dona
            wedges, texts, autotexts = ax.pie(sizes, labels=labels, colors=colors,
                                            autopct='%1.1f%%', startangle=90,
                                            wedgeprops=dict(width=0.3),
                                            textprops={'fontsize': 10, 'color': COLORES["texto_oscuro"], 'fontweight': 'bold'})

            # Mejorar la legibilidad de los porcentajes
            for autotext in autotexts:
                autotext.set_color('white')
                autotext.set_fontweight('bold')
                autotext.set_fontsize(11)

            # Estilizar las etiquetas
            for text in texts:
                text.set_fontweight('bold')
                text.set_fontsize(11)

            ax.axis('equal')
            ax.set_title('Distribución de Tipos', fontsize=13, fontweight='bold', 
                       color=COLORES["texto_oscuro"], pad=20)

            # Integrar en la interfaz
            self.dona_chart_canvas = FigureCanvasTkAgg(fig, master=self.dona_chart_frame)
            self.dona_chart_canvas.draw()
            self.dona_chart_canvas.get_tk_widget().pack(expand=True, fill="both", padx=5, pady=5)
            
            # Forzar actualización
            self.dona_chart_frame.update()
            print("✅ Gráfica de dona dibujada correctamente")
            
        except Exception as e:
            print(f"❌ Error dibujando gráfica de dona: {e}")
            import traceback
            traceback.print_exc()
            # Mostrar mensaje de error
            error_label = ctk.CTkLabel(self.dona_chart_frame, 
                                     text="❌ Error mostrando\ngráfica",
                                     font=(FUENTE, 12),
                                     text_color=COLORES["peligro"])
            error_label.pack(expand=True, fill="both")

    def actualizar_grafica_barras_compacta(self, df, medidas_col):
        """Actualiza la gráfica de barras compacta para medidas"""
        print("🎨 Dibujando gráfica de barras...")
        
        # Limpiar gráfica anterior
        if hasattr(self, 'bars_chart_canvas') and self.bars_chart_canvas:
            try:
                self.bars_chart_canvas.get_tk_widget().destroy()
            except:
                pass

        # Limpiar frame
        for widget in self.bars_chart_frame.winfo_children():
            widget.destroy()

        # Verificar si tenemos datos de medidas
        if medidas_col not in df.columns or df[medidas_col].isna().all():
            print("⚠️ No hay datos de medidas")
            no_data_label = ctk.CTkLabel(self.bars_chart_frame, 
                                       text="📏 No hay datos\nde medidas",
                                       font=(FUENTE, 12),
                                       text_color=COLORES["texto_oscuro"])
            no_data_label.pack(expand=True, fill="both")
            return

        try:
            # Obtener y normalizar medidas
            medidas_series = df[medidas_col].apply(self._normalizar_medida)
            
            # Filtrar valores nulos y contar frecuencia
            medidas_counts = medidas_series[medidas_series.notna()].value_counts()
            
            print(f"📊 Medidas encontradas: {len(medidas_counts)} tipos")
            
            # Si no hay medidas válidas
            if medidas_counts.empty:
                print("⚠️ No hay medidas válidas después de normalizar")
                no_data_label = ctk.CTkLabel(self.bars_chart_frame, 
                                           text="📏 No hay medidas\nválidas",
                                           font=(FUENTE, 12),
                                           text_color=COLORES["texto_oscuro"])
                no_data_label.pack(expand=True, fill="both")
                return

            # Ordenar por frecuencia y limitar a top 8
            medidas_counts = medidas_counts.nlargest(8).sort_values(ascending=True)
            print(f"📈 Top medidas: {dict(medidas_counts)}")

            # Crear figura para gráfica de barras VERTICALES
            fig, ax = plt.subplots(figsize=(6, 4), dpi=80)
            fig.patch.set_facecolor(COLORES["surface"])
            ax.set_facecolor(COLORES["surface"])

            # Colores para las barras
            colors = []
            for medida in medidas_counts.index:
                if any(especial in str(medida).upper() for especial in ["ESPECIAL", "VISTO BUENO", "NO IMPRIMIR"]):
                    colors.append(COLORES["warning"])
                else:
                    colors.append(COLORES["primario"])

            # Crear gráfica de barras VERTICALES
            bars = ax.bar(range(len(medidas_counts)), medidas_counts.values, color=colors, alpha=0.8, width=0.6)

            # Configurar ejes y etiquetas
            ax.set_xticks(range(len(medidas_counts)))
            
            # Rotar etiquetas para mejor legibilidad
            ax.set_xticklabels(medidas_counts.index, fontsize=9, rotation=45, ha='right')
            ax.set_ylabel('Cantidad', fontsize=11, color=COLORES["texto_oscuro"])
            
            # Añadir valores en las barras
            for i, (bar, count) in enumerate(zip(bars, medidas_counts.values)):
                height = bar.get_height()
                ax.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                       f'{count}', ha='center', va='bottom', fontsize=10, 
                       color=COLORES["texto_oscuro"], fontweight='bold')

            # Estilizar la gráfica
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_color(COLORES["texto_oscuro"])
            ax.spines['bottom'].set_color(COLORES["texto_oscuro"])
            ax.tick_params(colors=COLORES["texto_oscuro"], labelsize=10)
            ax.yaxis.label.set_color(COLORES["texto_oscuro"])

            ax.set_title('Top 8 Medidas Más Comunes', fontsize=13, fontweight='bold', 
                       color=COLORES["texto_oscuro"], pad=20)

            # Ajustar layout para evitar cortes
            plt.tight_layout()

            # Integrar en la interfaz
            self.bars_chart_canvas = FigureCanvasTkAgg(fig, master=self.bars_chart_frame)
            self.bars_chart_canvas.draw()
            self.bars_chart_canvas.get_tk_widget().pack(expand=True, fill="both", padx=5, pady=5)
            
            # Forzar actualización
            self.bars_chart_frame.update()
            print("✅ Gráfica de barras dibujada correctamente")
            
        except Exception as e:
            print(f"❌ Error dibujando gráfica de barras: {e}")
            import traceback
            traceback.print_exc()
            # Mostrar mensaje de error
            error_label = ctk.CTkLabel(self.bars_chart_frame, 
                                     text="❌ Error mostrando\ngráfica",
                                     font=(FUENTE, 12),
                                     text_color=COLORES["peligro"])
            error_label.pack(expand=True, fill="both")

    def mostrar_estado_vacio(self):
        """Muestra estado vacío en el dashboard"""
        print("📭 Mostrando estado vacío...")
        
        self.actualizar_tarjeta_metrica(self.card_total, "0")
        self.actualizar_tarjeta_metrica(self.card_categorias, "0")
        self.actualizar_tarjeta_metrica(self.card_normal, "0")
        self.actualizar_tarjeta_metrica(self.card_especial, "0")
        self.actualizar_tarjeta_metrica(self.card_metalicas, "0")
        
        # Limpiar gráficas
        if hasattr(self, 'dona_chart_canvas') and self.dona_chart_canvas:
            try:
                self.dona_chart_canvas.get_tk_widget().destroy()
            except:
                pass

        if hasattr(self, 'bars_chart_canvas') and self.bars_chart_canvas:
            try:
                self.bars_chart_canvas.get_tk_widget().destroy()
            except:
                pass
        
        # Mostrar mensajes de no datos en ambas gráficas
        for widget in self.dona_chart_frame.winfo_children():
            widget.destroy()
        no_data_dona = ctk.CTkLabel(self.dona_chart_frame, 
                                 text="📊 No hay datos\npara mostrar",
                                 font=(FUENTE, 12),
                                 text_color=COLORES["texto_oscuro"])
        no_data_dona.pack(expand=True, fill="both")

        for widget in self.bars_chart_frame.winfo_children():
            widget.destroy()
        no_data_bars = ctk.CTkLabel(self.bars_chart_frame, 
                                 text="📏 No hay datos\nde medidas",
                                 font=(FUENTE, 12),
                                 text_color=COLORES["texto_oscuro"])
        no_data_bars.pack(expand=True, fill="both")

    def limpiar_lista(self):
        """Limpia el historial de bases procesadas"""
        resp = messagebox.askyesno(
            "Confirmar limpieza",
            "¿Deseas limpiar TODO el historial de bases procesadas del Dashboard?\n"
            "⚠️ Tus archivos no serán eliminados."
        )
        if not resp:
            return

        historial_path = os.path.join(os.getcwd(), "data", "historial_bases.json")

        try:
            with open(historial_path, "w", encoding="utf-8") as f:
                json.dump([], f, ensure_ascii=False, indent=4)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo limpiar el historial:\n{e}")
            return

        self.archivos_procesados = []
        self.lista_archivos.configure(state="normal")
        self.lista_archivos.delete("1.0", "end")
        self.lista_archivos.insert("end", "📂 No hay bases generadas aún.\n")
        self.lista_archivos.configure(state="disabled")

        messagebox.showinfo("Listo", "✅ Historial limpiado correctamente.\nTus archivos permanecen intactos.")

    def borrar_seleccionado(self):
        """Borra el archivo seleccionado del historial"""
        sel_ranges = self.lista_archivos.tag_ranges("sel")
        if not sel_ranges:
            messagebox.showinfo("Seleccionar", "Selecciona una línea para borrar")
            return

        start = sel_ranges[0]
        line = str(start).split('.')[0]
        nombre = self.lista_archivos.get(f"{line}.0", f"{line}.end").strip()

        if not nombre:
            return

        resp = messagebox.askyesno(
            "Confirmar borrado",
            f"¿Deseas eliminar {nombre} del historial del Dashboard?\n"
            "Esta acción NO borrará el archivo físico."
        )
        if not resp:
            return

        historial_path = os.path.join(os.getcwd(), "data", "historial_bases.json")

        try:
            if os.path.exists(historial_path):
                with open(historial_path, "r", encoding="utf-8") as f:
                    historial = json.load(f)
                
                if isinstance(historial, dict):
                    historial = [historial]

                nuevo_historial = []
                for entry in historial:
                    if isinstance(entry, dict):
                        cargados = entry.get("archivos_cargados", {})
                        cargados = {k: v for k, v in cargados.items() if v != nombre}
                        entry["archivos_cargados"] = cargados
                        if cargados:
                            nuevo_historial.append(entry)

                with open(historial_path, "w", encoding="utf-8") as f:
                    json.dump(nuevo_historial, f, ensure_ascii=False, indent=4)

            self.archivos_procesados = [f for f in self.archivos_procesados if os.path.basename(f) != nombre]

            self.lista_archivos.configure(state="normal")
            self.lista_archivos.delete(f"{line}.0", f"{line}.end+1c")
            self.lista_archivos.configure(state="disabled")

            messagebox.showinfo("Listo", f"✅ {nombre} eliminado del historial del Dashboard.")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo borrar del historial:\n{e}")

    def mostrar_estado_vacio(self):
        """Muestra estado vacío en el dashboard"""
        print("📭 Mostrando estado vacío...")
        
        self.actualizar_tarjeta_metrica(self.card_total, "0")
        self.actualizar_tarjeta_metrica(self.card_categorias, "0")
        self.actualizar_tarjeta_metrica(self.card_normal, "0")
        self.actualizar_tarjeta_metrica(self.card_especial, "0")
        
        # Limpiar gráficas
        if hasattr(self, 'dona_chart_canvas') and self.dona_chart_canvas:
            try:
                self.dona_chart_canvas.get_tk_widget().destroy()
            except:
                pass

        if hasattr(self, 'bars_chart_canvas') and self.bars_chart_canvas:
            try:
                self.bars_chart_canvas.get_tk_widget().destroy()
            except:
                pass
        
        # Mostrar mensajes de no datos en ambas gráficas
        for widget in self.dona_chart_frame.winfo_children():
            widget.destroy()
        no_data_dona = ctk.CTkLabel(self.dona_chart_frame, 
                                 text="📊 No hay datos\npara mostrar",
                                 font=(FUENTE, 12),
                                 text_color=COLORES["texto_oscuro"])
        no_data_dona.pack(expand=True, fill="both")

        for widget in self.bars_chart_frame.winfo_children():
            widget.destroy()
        no_data_bars = ctk.CTkLabel(self.bars_chart_frame, 
                                 text="📏 No hay datos\nde medidas",
                                 font=(FUENTE, 12),
                                 text_color=COLORES["texto_oscuro"])
        no_data_bars.pack(expand=True, fill="both")

    def mostrar_archivos_procesados(self, archivos):
        """Muestra los archivos procesados en el TextBox"""
        self.lista_archivos.configure(state="normal")
        self.lista_archivos.delete("1.0", "end")

        if not archivos:
            self.lista_archivos.insert("end", "📂 No hay bases generadas aún.\n")
        else:
            for nombre, fecha in archivos:
                self.lista_archivos.insert("end", f"📄 {nombre} — {fecha}\n")

        self.lista_archivos.configure(state="disabled")

    def cargar_historial_bases(self):
        """Carga el historial de bases procesadas"""
        historial_path = os.path.join(os.getcwd(), "data", "historial_bases.json")

        if not os.path.exists(historial_path):
            self.mostrar_archivos_procesados([])
            return

        try:
            with open(historial_path, "r", encoding="utf-8") as f:
                historial = json.load(f)

            archivos_a_mostrar = []

            for entry in historial:
                if isinstance(entry, dict):
                    nombre = entry.get("nombre_archivo", "")
                    fecha = entry.get("fecha_generacion", "")
                    if nombre and fecha:
                        archivos_a_mostrar.append((nombre, fecha))

            self.mostrar_archivos_procesados(archivos_a_mostrar)

        except Exception as e:
            print(f"⚠️ Error cargando historial_bases.json: {e}")
            self.mostrar_archivos_procesados([])

    def exportar_pdf(self):
        from reportlab.pdfgen import canvas as pdf_canvas
        from reportlab.lib.pagesizes import letter
        from reportlab.lib.utils import ImageReader
        from reportlab.lib import colors as rl_colors
        from datetime import datetime
        import os
        import matplotlib.pyplot as plt
        from io import BytesIO

        # Preguntar al usuario dónde guardar el PDF
        ruta_guardar = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Archivos PDF", "*.pdf")],
            initialfile="Reporte de Base ULTA.pdf",
            title="Guardar reporte como..."
        )
        if not ruta_guardar:
            return

        try:
            # Datos principales
            total_codigos = getattr(self, 'total_upcs', 0)
            total_categorias_val = getattr(self, 'total_categorias', 0)
            normales = getattr(self, 'normales', 0)
            especiales = getattr(self, 'especiales', 0)
            # Leer historial actual para el PDF
            historial_path = os.path.join(os.getcwd(), "data", "historial_bases.json")
            archivos = []
            if os.path.exists(historial_path):
                with open(historial_path, "r", encoding="utf-8") as f:
                    historial = json.load(f)
                for entry in historial:
                    if isinstance(entry, dict):
                        nombre = entry.get("nombre_archivo")
                        if nombre:
                            archivos.append(nombre)

            # Preparar canvas
            c = pdf_canvas.Canvas(ruta_guardar, pagesize=letter)
            ancho, alto = letter
            pagina_actual = 1

            LOGO_PATH = os.path.join(os.getcwd(), 'img', 'logo_empresarial.jpeg')

            # Encabezado con franja amarilla y logo
            def dibujar_encabezado(titulo):
                c.setFillColor(rl_colors.HexColor('#ecd925'))
                c.rect(0, alto - 17, ancho, 17, fill=1, stroke=0)
                # Título
                c.setFillColor(rl_colors.HexColor('#282828'))
                c.setFont('Helvetica-Bold', 16)
                c.drawString(50, alto - 40, titulo)
                # Logo
                if os.path.exists(LOGO_PATH):
                    logo = ImageReader(LOGO_PATH)
                    c.drawImage(logo, ancho - 130, alto - 58, width=100, height=40, preserveAspectRatio=True)

            # Footer con número de página
            def dibujar_footer(pagina):
                c.setFillColor(rl_colors.HexColor('#282828'))
                c.rect(0, 0, ancho, 40, fill=1, stroke=0)
                c.setFillColor(rl_colors.HexColor('#FFFFFF'))
                c.setFont('Helvetica', 8)
                c.drawString(50, 18, 'Gestión de bases de etiquetado V&C')
                # Centro
                texto_centro = 'www.vandc.com'
                ancho_texto_centro = c.stringWidth(texto_centro, 'Helvetica', 8)
                c.drawString((ancho - ancho_texto_centro) / 2, 18, texto_centro)
                # Número de página
                texto_derecho = f'Página {pagina}'
                ancho_texto_derecho = c.stringWidth(texto_derecho, 'Helvetica', 8)
                c.drawString(ancho - ancho_texto_derecho - 50, 18, texto_derecho)

            # Nueva página
            def nueva_pagina(titulo):
                nonlocal pagina_actual, y
                dibujar_footer(pagina_actual)
                c.showPage()
                pagina_actual += 1
                dibujar_encabezado(titulo)
                y = alto - 120
                return y

            # Iniciar primera página
            y = alto - 120
            dibujar_encabezado('REPORTE DE ETIQUETADO ULTA')
            c.setFont('Helvetica', 10)
            c.drawString(50, alto - 70, f'Fecha: {datetime.now().strftime("%d/%m/%Y %H:%M")}')

            # Estadísticas principales
            c.setFont('Helvetica-Bold', 12)
            c.drawString(50, y, 'ESTADÍSTICAS PRINCIPALES')
            y -= 28
            c.setFont('Helvetica', 10)
            lineas = [
                f'• Total UPCs: {total_codigos}',
                f'• Total Categorías: {total_categorias_val}',
                f'• Normales: {normales}',
                f'• Especiales: {especiales}',
                f'• Total de archivos procesados: {len(archivos)}'
            ]
            for linea in lineas:
                if y < 120:
                    y = nueva_pagina('REPORTE DE ETIQUETADO')
                c.drawString(70, y, linea)
                y -= 18

            # Archivos procesados
            if archivos:
                if y < 140:
                    y = nueva_pagina('REPORTE DE ETIQUETADO')
                c.setFont('Helvetica-Bold', 12)
                c.drawString(50, y, 'BASES GENERADAS:')
                y -= 20
                c.setFont('Helvetica', 10)

                for archivo in archivos:
                    if y < 120:
                        y = nueva_pagina('REPORTE DE ETIQUETADO')
                        c.setFont('Helvetica-Bold', 12)
                        c.drawString(50, y, 'BASES GENERADAS:')
                        y -= 20
                        c.setFont('Helvetica', 10)

                    # Extraer nombre según tipo
                    if isinstance(archivo, str):
                        nombre = os.path.basename(archivo)
                    elif isinstance(archivo, dict):
                        nombre = archivo.get('nombre_archivo') or os.path.basename(archivo.get('ruta_archivo', ''))
                    else:
                        nombre = str(archivo)

                    c.drawString(70, y, f'• {nombre}')
                    y -= 16  # altura por línea, ajustada

            # Gráfica de pastel
            if total_codigos > 0:
                if y < 300:
                    y = nueva_pagina('REPORTE DE ETIQUETADO')
                etiquetas = ['Normales', 'Especiales']
                valores = [normales, especiales]
                colores_plot = ['#ECD925', '#282828']

                plt.figure(figsize=(6, 4), facecolor='white')  # fondo blanco para visibilidad
                wedges, texts, autotexts = plt.pie(
                    valores,
                    labels=etiquetas,  # nombres de categorías visibles
                    colors=colores_plot,
                    autopct='%1.1f%%',
                    startangle=90,
                    textprops={'fontsize': 12}  # texto de labels fuera del pastel
                )

                # Ajustar solo los números dentro del pastel
                for i, t in enumerate(autotexts):
                    if valores[i] > 0:
                        # blanco o negro según color del wedge
                        if colores_plot[i] in ['#282828']:  # wedge oscuro
                            t.set_color('white')
                        else:  # wedge claro
                            t.set_color('black')
                        t.set_fontsize(12)
                        t.set_fontweight('bold')

                plt.title('Distribución de Tipos de Etiquetas', fontsize=14, fontweight='bold', color='#282828', pad=10)
                plt.axis('equal')

                buf = BytesIO()
                plt.savefig(buf, format='PNG', dpi=150, bbox_inches='tight', facecolor='white')  # fondo blanco
                plt.close()
                buf.seek(0)
                c.drawImage(ImageReader(buf), 50, y - 260, width=420, height=260)

            # Footer última página
            dibujar_footer(pagina_actual)
            c.save()
            messagebox.showinfo('Éxito', f'PDF generado correctamente en:\n{ruta_guardar}')

        except Exception as e:
            messagebox.showerror('Error', f'No se pudo generar el PDF:\n{e}')

