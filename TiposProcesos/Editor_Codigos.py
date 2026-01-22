import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import json
import os
import numpy as np
import Rutas
import sys

def recurso_path(ruta_relativa):
    """
    Devuelve la ruta correcta del recurso, ya sea en desarrollo o en .exe.
    """
    try:
        # Ruta temporal dentro del ejecutable
        base_path = sys._MEIPASS
    except AttributeError:
        # Ruta normal durante desarrollo
        base_path = os.path.abspath(".")
    return os.path.join(base_path, ruta_relativa)


class EditorCodigos:
    def __init__(self, parent, archivo_excel, archivo_json):
        self.parent = parent
        self.ARCHIVO_CODIGOS = archivo_excel
        # Asegurar que siempre se guarde en la ruta que el dashboard lee
        self.ARCHIVO_JSON_DASHBOARD = recurso_path("datos/codigos_cumple.json")
        self.ARCHIVO_JSON = archivo_json  # Puede ser diferente, pero el dashboard siempre leerá de ARCHIVO_JSON_DASHBOARD
        self.df_codigos_cumple = pd.DataFrame()

        self.cargar_datos()
        self.crear_ventana()

    def cargar_datos(self):
        """Carga los datos desde Excel y JSON"""
        try:
            # Cargar Excel
            if os.path.exists(self.ARCHIVO_CODIGOS):
                self.df_codigos_cumple = pd.read_excel(self.ARCHIVO_CODIGOS)
                # Reemplazar NaN por cadenas vacías en CRITERIO cuando OBSERVACIONES es "CUMPLE"
                mask_cumple = self.df_codigos_cumple["OBSERVACIONES"].astype(str).str.upper() == "CUMPLE"
                self.df_codigos_cumple.loc[mask_cumple, "CRITERIO"] = ""
            else:
                self.df_codigos_cumple = pd.DataFrame(columns=["ITEM", "OBSERVACIONES", "CRITERIO"])

            # Cargar JSON (opcional) - pero el dashboard siempre leerá de ARCHIVO_JSON_DASHBOARD
            if os.path.exists(self.ARCHIVO_JSON):
                with open(self.ARCHIVO_JSON, "r", encoding="utf-8") as f:
                    data_json = json.load(f)
                    # sincronizar si quieres
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron cargar los datos: {str(e)}")

    def crear_ventana(self):
        self.ventana = tk.Toplevel(self.parent)
        self.ventana.title("📋 Editor de Códigos")
        
        # Tamaño de la ventana
        ancho_ventana = 900
        alto_ventana = 600
        self.ventana.geometry(f"{ancho_ventana}x{alto_ventana}")
        self.ventana.resizable(False, False)
        self.ventana.configure(bg="#FFFFFF")
        self.centrar_ventana(ancho_ventana, alto_ventana)

        # Título
        title_frame = tk.Frame(self.ventana, bg="#FFFFFF", pady=15)
        title_frame.pack(fill="x")
        tk.Label(title_frame, text="EDITOR DE CÓDIGOS", font=("Segoe UI", 20, "bold"),
                bg="#FFFFFF", fg="#282828").pack()

        # Buscador
        search_frame = tk.Frame(self.ventana, bg="#FFFFFF", pady=15)
        search_frame.pack(fill="x", padx=20)
        search_inner_frame = tk.Frame(search_frame, bg="#F5F5F5", relief="solid", bd=1)
        search_inner_frame.pack(fill="x", padx=10, pady=10)
        tk.Label(search_inner_frame, text="🔍 Buscar:", bg="#F5F5F5", fg="#282828",
                font=("INTER", 11)).pack(side="left", padx=10)
        
        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(search_inner_frame, textvariable=self.search_var,
                                    font=("INTER", 11), bg="#FFFFFF", fg="#282828",
                                    relief="flat", width=50)
        self.search_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        btn_search = tk.Button(search_inner_frame, text="Buscar", bg="#ECD925", fg="#282828",
                            font=("INTER", 9, "bold"), command=self.filtrar_tabla,
                            relief="flat", padx=15)
        btn_search.pack(side="left", padx=5)
        
        btn_clear = tk.Button(search_inner_frame, text="Limpiar", bg="#E0E0E0", fg="#282828",
                            font=("INTER", 9), command=self.actualizar_tabla,
                            relief="flat", padx=15)
        btn_clear.pack(side="left", padx=5)

        # Tabla con scrollbars
        table_frame = tk.Frame(self.ventana, bg="#FFFFFF")
        table_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        v_scrollbar = ttk.Scrollbar(table_frame)
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar = ttk.Scrollbar(table_frame, orient="horizontal")
        h_scrollbar.pack(side="bottom", fill="x")

        style = ttk.Style()
        style.configure("Treeview", 
                        font=("INTER", 10), rowheight=25,
                        background="#FFFFFF", fieldbackground="#FFFFFF", foreground="#282828")
        style.configure("Treeview.Heading", 
                        font=("INTER", 11, "bold"),
                        background="#ECD925", foreground="#282828", relief="flat")
        
        self.tree = ttk.Treeview(table_frame, columns=("ITEM", "OBSERVACIONES", "CRITERIO"), 
                                show="headings", yscrollcommand=v_scrollbar.set,
                                xscrollcommand=h_scrollbar.set)
        
        self.tree.heading("ITEM", text="ITEM")
        self.tree.heading("OBSERVACIONES", text="OBSERVACIONES")
        self.tree.heading("CRITERIO", text="CRITERIO")
        self.tree.column("ITEM", width=150, anchor="center", minwidth=100)
        self.tree.column("OBSERVACIONES", width=450, anchor="w", minwidth=300)
        self.tree.column("CRITERIO", width=150, anchor="w", minwidth=100)
        self.tree.pack(fill="both", expand=True)

        v_scrollbar.config(command=self.tree.yview)
        h_scrollbar.config(command=self.tree.xview)

        # Botones centrados en la parte inferior
        button_frame = tk.Frame(self.ventana, bg="#FFFFFF", pady=15)
        button_frame.pack(fill="x")
        
        inner_buttons = tk.Frame(button_frame, bg="#FFFFFF")
        inner_buttons.pack()

        buttons_config = [
            ("➕ Agregar Item", self.abrir_agregar_item, "#ECD925"),
            ("✏️ Editar Item", self.abrir_editar_item, "#ECD925"),
            ("📤 Subir Excel", self.importar_excel, "#ECD925"),
            ("🗑️ Eliminar Item", self.eliminar_item_principal, "#FF4444"),  # Rojo para indicar peligro
            ("❌ Cerrar", self.ventana.destroy, "#282828")
        ]
        
        for text, command, color in buttons_config:
            btn = tk.Button(inner_buttons, text=text, bg=color, fg="#FFFFFF" if color=="#282828" else "#282828",
                            font=("INTER", 10, "bold"), command=command, relief="flat", padx=15, pady=8)
            btn.pack(side="left", padx=5)

        self.actualizar_tabla()
        self.search_entry.bind("<Return>", lambda event: self.filtrar_tabla())

    def eliminar_item_principal(self):
        """Elimina el item seleccionado con confirmación"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Eliminar Item", "Seleccione un item de la tabla para eliminar")
            return
        
        # Obtener el valor del ITEM seleccionado
        item_values = self.tree.item(selected[0])["values"]
        if not item_values or item_values[0] in ["No hay datos", "No se encontraron resultados"]:
            return
            
        item_id = item_values[0]
        observaciones = item_values[1]
        
        # Mostrar alerta de confirmación
        confirmacion = messagebox.askyesno(
            "⚠️ CONFIRMAR ELIMINACIÓN",
            f"¿Está seguro que desea eliminar el item?\n\n"
            f"ITEM: {item_id}\n"
            f"OBSERVACIONES: {observaciones}\n\n"
            f"Esta acción no se puede deshacer.",
            icon="warning"
        )
        
        if not confirmacion:
            return
        
        try:
            # Encontrar y eliminar el item del DataFrame
            mask = self.df_codigos_cumple["ITEM"].astype(str) == str(item_id)
            if mask.any():
                # Eliminar el item
                self.df_codigos_cumple = self.df_codigos_cumple[~mask]
                
                # Guardar los cambios
                self.guardar_datos()
                
                # Actualizar la tabla
                self.actualizar_tabla()
                
                messagebox.showinfo("Eliminar Item", f"Item '{item_id}' eliminado correctamente")
            else:
                messagebox.showerror("Error", "No se pudo encontrar el item seleccionado")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error al eliminar item: {str(e)}")

    def centrar_ventana(self, ancho, alto):
        """Centra la ventana en la pantalla con el tamaño específico"""
        self.ventana.update_idletasks()
        x = (self.ventana.winfo_screenwidth() // 2) - (ancho // 2)
        y = (self.ventana.winfo_screenheight() // 2) - (alto // 2)
        self.ventana.geometry(f'{ancho}x{alto}+{x}+{y}')

    def actualizar_tabla(self):
        """Actualiza la tabla con los datos actuales"""
        for row in self.tree.get_children():
            self.tree.delete(row)
        
        if len(self.df_codigos_cumple) == 0:
            self.tree.insert("", "end", values=("No hay datos", "", ""))
            return
            
        for _, row in self.df_codigos_cumple.iterrows():
            # Formatear los valores para mostrar vacío en lugar de nan
            item = row["ITEM"]
            observaciones = row["OBSERVACIONES"]
            criterio = row["CRITERIO"]
            
            # Convertir a cadena y reemplazar 'nan' por vacío
            if pd.isna(criterio) or str(criterio).lower() == 'nan':
                criterio = ""
            
            self.tree.insert("", "end", values=(item, observaciones, criterio))

    def filtrar_tabla(self):
        """Filtra la tabla por el valor de búsqueda"""
        busqueda = self.search_var.get().lower()
        
        for row in self.tree.get_children():
            self.tree.delete(row)
            
        if not busqueda:
            self.actualizar_tabla()
            return
            
        # Crear máscara para la búsqueda en todos los campos
        mask = (
            self.df_codigos_cumple["ITEM"].astype(str).str.lower().str.contains(busqueda) |
            self.df_codigos_cumple["OBSERVACIONES"].astype(str).str.lower().str.contains(busqueda) |
            self.df_codigos_cumple["CRITERIO"].astype(str).str.lower().str.contains(busqueda)
        )
        
        resultados = self.df_codigos_cumple[mask]
        
        if len(resultados) == 0:
            self.tree.insert("", "end", values=("No se encontraron resultados", "", ""))
            return
            
        for _, row in resultados.iterrows():
            # Formatear los valores para mostrar vacío en lugar de nan
            item = row["ITEM"]
            observaciones = row["OBSERVACIONES"]
            criterio = row["CRITERIO"]
            
            # Convertir a cadena en lugar de nan
            if pd.isna(criterio) or str(criterio).lower() == 'nan':
                criterio = ""
            
            self.tree.insert("", "end", values=(item, observaciones, criterio))

    def abrir_agregar_item(self):
        AgregarItem(self)

    def abrir_editar_item(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Editar Item", "Seleccione an item de la tabla")
            return
        
        # Obtener el valor del ITEM seleccionado
        item_values = self.tree.item(selected[0])["values"]
        if not item_values or item_values[0] in ["No hay datos", "No se encontraron resultados"]:
            return
            
        item_id = item_values[0]
        
        # Encontrar el índice en el DataFrame
        try:
            mask = self.df_codigos_cumple["ITEM"].astype(str) == str(item_id)
            if mask.any():
                index = self.df_codigos_cumple[mask].index[0]
                EditorItem(self, index)
            else:
                messagebox.showerror("Error", "No se pudo encontrar el item seleccionado")
        except Exception as e:
            messagebox.showerror("Error", f"Error al editar item: {str(e)}")

    def guardar_datos(self):
        """Guarda los datos a Excel y JSON (en la ruta que el dashboard lee)"""
        try:
            # Asegurarse de que los valores NaN se guarden como vacíos
            self.df_codigos_cumple["CRITERIO"] = self.df_codigos_cumple["CRITERIO"].replace({np.nan: "", "nan": ""})
            
            # Guardar en Excel (archivo original)
            self.df_codigos_cumple.to_excel(self.ARCHIVO_CODIGOS, index=False)
            
            # Crear el directorio de la ruta del dashboard si no existe
            dashboard_dir = os.path.dirname(self.ARCHIVO_JSON_DASHBOARD)
            if dashboard_dir and not os.path.exists(dashboard_dir):
                os.makedirs(dashboard_dir)

            # Guardar en JSON en la ruta que el dashboard lee (siempre absoluta)
            self.df_codigos_cumple.to_json(self.ARCHIVO_JSON_DASHBOARD, orient="records", force_ascii=False, indent=4)

            # También guardar en el JSON original si es diferente
            if self.ARCHIVO_JSON != self.ARCHIVO_JSON_DASHBOARD:
                json_dir = os.path.dirname(self.ARCHIVO_JSON)
                if json_dir and not os.path.exists(json_dir):
                    os.makedirs(json_dir)
                self.df_codigos_cumple.to_json(self.ARCHIVO_JSON, orient="records", force_ascii=False, indent=4)

            messagebox.showinfo("Guardar", "Datos guardados correctamente. El dashboard se actualizará automáticamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron guardar los datos: {str(e)}")




    def importar_excel(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel para importar",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not file_path:
            return

        try:
            # --- Leer archivo Excel ---
            df_nuevo = pd.read_excel(file_path)

            # Columnas requeridas
            columnas_requeridas = {"ITEM", "CRITERIO"}
            if not columnas_requeridas.issubset(df_nuevo.columns):
                messagebox.showerror("Error", f"El archivo Excel debe contener las columnas: {', '.join(columnas_requeridas)}")
                return

            # Limpiar y convertir ITEM a número
            df_nuevo = df_nuevo.fillna("")
            df_nuevo["ITEM"] = (
                df_nuevo["ITEM"].astype(str)
                .str.extract(r"(\d+)")  # extrae solo números
                .astype(float)
                .astype("Int64")
            )

            # Guardar el CRITERIO original para mostrar en la columna "Obs. Nueva"
            df_nuevo["CRITERIO_ORIGINAL"] = df_nuevo["CRITERIO"].astype(str).str.strip()

            # Asignar CRITERIO automáticamente para la lógica interna
            df_nuevo["CRITERIO"] = df_nuevo["CRITERIO_ORIGINAL"].str.upper()
            df_nuevo["CRITERIO"] = df_nuevo["CRITERIO"].apply(lambda x: "" if x == "CUMPLE" else "REVISADO")

            df_nuevo = df_nuevo[["ITEM", "CRITERIO_ORIGINAL", "CRITERIO"]].copy()

            # Crear columnas de salida para el treeview
            df_nuevo["OBS_NEW"] = df_nuevo["CRITERIO_ORIGINAL"]  # Obs. Nueva = el valor original del archivo
            df_nuevo["CRIT_NEW"] = df_nuevo["CRITERIO"]           # Crit. Nueva = lógica interna (REVISADO o vacío)

            # 🔹 Filtrar filas que tienen Obs. Nueva vacía
            # 🔹 Filtrar filas que tienen Obs. Nueva vacía
            df_nuevo = df_nuevo[df_nuevo["OBS_NEW"].str.strip() != ""]

            # 🔹 Eliminar filas donde OBS_NEW contenga "REVISADO" (de la columna CRITERIO original del archivo)
            df_nuevo = df_nuevo[~df_nuevo["OBS_NEW"].str.upper().str.contains("REVISADO", na=False)]


            # Si no hay datos existentes, importar todo
            if self.df_codigos_cumple is None or self.df_codigos_cumple.empty:
                self.df_codigos_cumple = df_nuevo.copy()
                self.actualizar_tabla()
                self.guardar_datos()
                messagebox.showinfo("Importar Excel", "Todos los datos han sido importados correctamente.")
                return

            # Preparar datos existentes
            dict_existente = self.df_codigos_cumple.set_index("ITEM").to_dict("index")

            # Detectar cambios y nuevos items
            cambios = []
            nuevos_items = []
            actualizaciones = 0

            for _, row in df_nuevo.iterrows():
                item = row["ITEM"]
                obs_nuevo = row["OBS_NEW"]     # valor original del archivo
                crit_nuevo = row["CRIT_NEW"]   # REVISADO o vacío según la lógica

                if item in dict_existente:
                    obs_actual = dict_existente[item].get("OBSERVACIONES", "")
                    crit_actual = dict_existente[item].get("CRITERIO", "")

                    # Verificar si hay algún cambio en OBSERVACIONES o CRITERIO
                    if obs_actual != obs_nuevo or crit_actual != crit_nuevo:
                        cambios.append({
                            "item": item,
                            "obs_actual": obs_actual,
                            "crit_actual": crit_actual,
                            "obs_nuevo": obs_nuevo,
                            "crit_nuevo": crit_nuevo,
                            "tipo": "actualización"
                        })
                        actualizaciones += 1

                else:
                    cambios.append({
                        "item": item,
                        "obs_actual": "",
                        "crit_actual": "",
                        "obs_nuevo": obs_nuevo,
                        "crit_nuevo": crit_nuevo,
                        "tipo": "nuevo"
                    })
                    nuevos_items.append(item)


            # Crear ventana de revisión más compacta
            win = tk.Toplevel(self.parent if hasattr(self, "parent") else None)
            win.title("Revisión de cambios - Importar Excel")
            win.geometry("1200x550")
            win.configure(bg="#f5f5f5")
            win.minsize(900, 450)
            
            # Frame principal
            main_frame = ttk.Frame(win, padding="10")
            main_frame.pack(fill="both", expand=True)
            
            # Título y estadísticas
            title_frame = ttk.Frame(main_frame)
            title_frame.pack(fill="x", pady=(0, 10))
            
            ttk.Label(title_frame, text="Revisión de Cambios", 
                    font=("Arial", 12, "bold")).pack(anchor="w")
            
            stats_frame = ttk.Frame(title_frame)
            stats_frame.pack(fill="x", pady=(5, 0))
            
            ttk.Label(stats_frame, text=f"• Nuevos items: {len(nuevos_items)}", 
                    foreground="green", font=("Arial", 9)).pack(side="left", padx=(0, 15))
            ttk.Label(stats_frame, text=f"• Actualizaciones: {actualizaciones}", 
                    foreground="blue", font=("Arial", 9)).pack(side="left")
            
            # Frame para la tabla con scrollbars
            table_frame = ttk.Frame(main_frame)
            table_frame.pack(fill="both", expand=True)
            
            # Configurar scrollbars
            v_scrollbar = ttk.Scrollbar(table_frame, orient="vertical")
            v_scrollbar.pack(side="right", fill="y")
            
            h_scrollbar = ttk.Scrollbar(table_frame, orient="horizontal")
            h_scrollbar.pack(side="bottom", fill="x")
            
            # Crear treeview con nombres de columnas más cortos
            cols = ("ITEM", "TIPO", "OBS_ACT", "CRIT_ACT", "OBS_NEW", "CRIT_NEW", "ACTUALIZAR")
            col_names = {
                "ITEM": "Item",
                "TIPO": "Tipo",
                "OBS_ACT": "Obs. Actual",
                "CRIT_ACT": "Crit. Actual", 
                "OBS_NEW": "Obs. Nueva",
                "CRIT_NEW": "Crit. Nueva",
                "ACTUALIZAR": "Actualizar"
            }
            
            # Definir tree como variable local dentro de esta función
            tree = ttk.Treeview(table_frame, columns=cols, show="headings", 
                            yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set,
                            height=12, selectmode="extended")
            
            # Configurar columnas con autoajuste y nombres cortos
            column_configs = {
                "ITEM": {"width": 80, "minwidth": 60, "stretch": False},
                "TIPO": {"width": 90, "minwidth": 70, "stretch": False},
                "OBS_ACT": {"width": 180, "minwidth": 120, "stretch": True},
                "CRIT_ACT": {"width": 180, "minwidth": 120, "stretch": True},
                "OBS_NEW": {"width": 180, "minwidth": 120, "stretch": True},
                "CRIT_NEW": {"width": 180, "minwidth": 120, "stretch": True},
                "ACTUALIZAR": {"width": 100, "minwidth": 80, "stretch": False}
            }
            
            for col in cols:
                tree.heading(col, text=col_names[col])
                config = column_configs[col]
                tree.column(col, width=config["width"], minwidth=config["minwidth"], stretch=config["stretch"])
            
            # Insertar datos en el treeview
            for cambio in cambios:
                tipo_texto = "NUEVO" if cambio["tipo"] == "nuevo" else "ACTUAL"
                
                item_id = tree.insert("", "end", values=(
                    cambio["item"],
                    tipo_texto,
                    cambio["obs_actual"],
                    cambio["crit_actual"],
                    cambio["obs_nuevo"],
                    cambio["crit_nuevo"],
                    "Sí"  # Por defecto seleccionado
                ))
                
                # Aplicar color según el tipo
                if cambio["tipo"] == "nuevo":
                    tree.set(item_id, "TIPO", tipo_texto)
            
            tree.pack(fill="both", expand=True)
            
            # Configurar scrollbars
            v_scrollbar.config(command=tree.yview)
            h_scrollbar.config(command=tree.xview)
            
            # Autoajustar columnas al contenido después de insertar datos
            def autoajustar_columnas():
                for col in cols:
                    # Obtener el ancho máximo del contenido
                    max_width = 0
                    for item in tree.get_children():
                        cell_value = tree.set(item, col)
                        width = len(str(cell_value)) * 7 + 15  # Aproximación basada en longitud
                        if width > max_width:
                            max_width = width
                    
                    # Limitar el ancho máximo para no hacer columnas demasiado grandes
                    max_width = min(max_width, 300)
                    
                    # Ajustar la columna
                    tree.column(col, width=max_width)
            
            # Llamar después de un breve retraso para que se renderice el contenido
            win.after(100, autoajustar_columnas)
            
            # Función para editar celdas
            def editar_celda(event):
                region = tree.identify("region", event.x, event.y)
                if region == "cell":
                    item = tree.identify_row(event.y)
                    column = tree.identify_column(event.x)
                    col_index = int(column.replace("#", "")) - 1
                    col_name = cols[col_index]
                    
                    # No permitir editar columnas de solo lectura
                    if col_name in ["TIPO"]:
                        return
                    
                    # Obtener coordenadas y valor actual
                    x, y, width, height = tree.bbox(item, column)
                    current_value = tree.set(item, col_name)
                    
                    # Crear widget de edición
                    if col_name == "ACTUALIZAR":
                        # Combobox para Sí/No
                        edit_widget = ttk.Combobox(tree, values=["Sí", "No"], state="readonly")
                        edit_widget.set(current_value)
                    else:
                        # Entry para texto normal
                        edit_widget = ttk.Entry(tree)
                        edit_widget.insert(0, current_value)
                    
                    edit_widget.place(x=x, y=y, width=width, height=height)
                    edit_widget.focus()
                    edit_widget.select_range(0, tk.END)
                    
                    def guardar_edicion(event=None):
                        nuevo_valor = edit_widget.get()
                        tree.set(item, col_name, nuevo_valor)
                        edit_widget.destroy()
                    
                    def cancelar_edicion(event=None):
                        edit_widget.destroy()
                    
                    edit_widget.bind("<Return>", guardar_edicion)
                    edit_widget.bind("<Escape>", cancelar_edicion)
                    edit_widget.bind("<FocusOut>", guardar_edicion)
            
            tree.bind("<Double-1>", editar_celda)
            
            # Función para alternar selección
            def toggle_selection(select_all):
                for item in tree.get_children():
                    new_val = "Sí" if select_all else "No"
                    tree.set(item, "ACTUALIZAR", new_val)
            
            # Frame de botones de selección
            selection_frame = ttk.Frame(main_frame)
            selection_frame.pack(fill="x", pady=(10, 5))
            
            ttk.Button(selection_frame, text="Seleccionar Todos", 
                    command=lambda: toggle_selection(True)).pack(side="left", padx=(0, 5))
            ttk.Button(selection_frame, text="Deseleccionar Todos", 
                    command=lambda: toggle_selection(False)).pack(side="left", padx=(0, 10))
            
            # Función para aplicar cambios
            def aplicar_cambios():
                # Contadores para estadísticas
                aplicados_nuevos = 0
                aplicados_actualizaciones = 0
                
                # Aplicar cambios seleccionados
                for item_id in tree.get_children():
                    vals = tree.item(item_id)["values"]
                    seleccionado = vals[6]  # Columna ACTUALIZAR
                    
                    if seleccionado == "Sí":
                        item = vals[0]
                        obs_nuevo = vals[4]
                        crit_nuevo = vals[5]
                        tipo = "nuevo" if vals[1] == "NUEVO" else "actualización"
                        
                        # Actualizar el diccionario existente
                        dict_existente[item] = {"OBSERVACIONES": obs_nuevo, "CRITERIO": crit_nuevo}
                        
                        # Contar según el tipo
                        if tipo == "nuevo":
                            aplicados_nuevos += 1
                        else:
                            aplicados_actualizaciones += 1
                
                # Convertir de vuelta a DataFrame
                self.df_codigos_cumple = pd.DataFrame.from_dict(dict_existente, orient="index").reset_index()
                self.df_codigos_cumple = self.df_codigos_cumple.rename(columns={"index": "ITEM"})
                
                # Guardar los cambios (esto guardará en la ruta del dashboard)
                self.guardar_datos()
                
                # Actualizar la tabla principal
                self.actualizar_tabla()
                
                # Mostrar resumen
                resumen = f"Cambios aplicados y guardados exitosamente:\n\n• Nuevos items: {aplicados_nuevos}\n• Actualizaciones: {aplicados_actualizaciones}"
                messagebox.showinfo("Importar Excel", resumen)
                
                win.destroy()
            
            # Frame de botones de acción en la parte inferior
            action_frame = ttk.Frame(main_frame)
            action_frame.pack(fill="x", pady=(5, 0))
            
            # Botones alineados a la izquierda y derecha
            ttk.Button(action_frame, text="Cancelar", 
                    command=win.destroy).pack(side="right", padx=(5, 0))
            ttk.Button(action_frame, text="Aplicar Cambios", 
                    command=aplicar_cambios).pack(side="right", padx=(5, 0))
            
            # Centrar ventana
            win.transient(self.parent if hasattr(self, "parent") else None)
            win.grab_set()
            win.update_idletasks()
            width = win.winfo_width()
            height = win.winfo_height()
            x = (win.winfo_screenwidth() // 2) - (width // 2)
            y = (win.winfo_screenheight() // 2) - (height // 2)
            win.geometry(f"+{x}+{y}")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo importar el Excel: {str(e)}")




class AgregarItem:
    def __init__(self, editor: EditorCodigos):
        self.editor = editor
        self.ventana = tk.Toplevel(editor.ventana)
        self.ventana.title("Agregar Nuevo Item")
        
        # TAMAÑO ESPECÍFICO PARA VENTANA DE AGREGAR
        self.ventana.geometry("500x400")
        self.ventana.resizable(False, False)
        self.ventana.configure(bg="#FFFFFF")
        
        # Centrar ventana
        self.ventana.transient(editor.ventana)
        self.ventana.grab_set()
        self.centrar_ventana()

        # Frame principal
        main_frame = tk.Frame(self.ventana, bg="#FFFFFF", padx=20, pady=20)
        main_frame.pack(fill="both", expand=True)
        
        tk.Label(main_frame, text="AGREGAR NUEVO ITEM", font=("Segoe UI", 14, "bold"),
                 bg="#FFFFFF", fg="#282828").pack(pady=10)

        # Campos de entrada
        campos = [
            ("ITEM:", "item_entry"),
            ("OBSERVACIONES:", "obs_entry"),
            ("CRITERIO:", "crit_entry")
        ]
        
        for label_text, attr_name in campos:
            frame = tk.Frame(main_frame, bg="#FFFFFF")
            frame.pack(fill="x", pady=8)
            
            tk.Label(frame, text=label_text, bg="#FFFFFF", fg="#282828",
                     font=("Segoe UI", 10)).pack(anchor="w")
            
            entry = tk.Entry(frame, font=("Segoe UI", 10), bg="#FFFFFF", 
                            fg="#282828", relief="solid", bd=1)
            entry.pack(fill="x", pady=5)
            setattr(self, attr_name, entry)

        # 🔹 Vinculamos el evento para desactivar criterio si es "CUMPLE"
        self.obs_entry.bind("<KeyRelease>", self.verificar_observacion)

        # Botón guardar
        btn_frame = tk.Frame(main_frame, bg="#FFFFFF")
        btn_frame.pack(pady=20)
        
        tk.Button(btn_frame, text="💾 Guardar", bg="#ECD925", fg="#282828",
                  font=("Segoe UI", 10, "bold"), command=self.agregar_item,
                  relief="flat", padx=20, pady=8).pack(side="left", padx=10)
                  
        tk.Button(btn_frame, text="❌ Cancelar", bg="#E0E0E0", fg="#282828",
                  font=("Segoe UI", 10), command=self.ventana.destroy,
                  relief="flat", padx=20, pady=8).pack(side="left", padx=10)
        
    def centrar_ventana(self):
        """Centra la ventana en la pantalla"""
        self.ventana.update_idletasks()
        ancho = self.ventana.winfo_width()
        alto = self.ventana.winfo_height()
        x = (self.ventana.winfo_screenwidth() // 2) - (ancho // 2)
        y = (self.ventana.winfo_screenheight() // 2) - (alto // 2)
        self.ventana.geometry(f'+{x}+{y}')

    def verificar_observacion(self, event=None):
        """Desactiva la entrada de criterio si la observación es 'CUMPLE'"""
        if self.obs_entry.get().strip().upper() == "CUMPLE":
            self.crit_entry.delete(0, tk.END)
            self.crit_entry.config(state="disabled")
        else:
            self.crit_entry.config(state="normal")

    def agregar_item(self):
        item = self.item_entry.get().strip()
        observaciones = self.obs_entry.get().strip()
        
        # Si la entrada está deshabilitada, criterio queda vacío
        if self.crit_entry.cget("state") == "disabled":
            criterio = ""
        else:
            criterio = self.crit_entry.get().strip()
        
        if not item:
            messagebox.showwarning("Advertencia", "El campo ITEM no puede estar vacío")
            return
            
        # Verificar si el ITEM ya existe
        if item in self.editor.df_codigos_cumple["ITEM"].astype(str).values:
            messagebox.showwarning("Advertencia", f"El ITEM {item} ya existe")
            return
            
        # Seguridad extra: si la observación es "CUMPLE", limpiar criterio
        if observaciones.upper() == "CUMPLE":
            criterio = ""
            
        nuevo = {"ITEM": item, "OBSERVACIONES": observaciones, "CRITERIO": criterio}
        self.editor.df_codigos_cumple = pd.concat(
            [self.editor.df_codigos_cumple, pd.DataFrame([nuevo])],
            ignore_index=True
        )

        # 🔹 Guardar de inmediato en Excel y JSON (incluyendo la ruta del dashboard)
        self.editor.guardar_datos()

        # Refrescar tabla y cerrar
        self.editor.actualizar_tabla()
        self.ventana.destroy()
        messagebox.showinfo("Éxito", "Item agregado y guardado correctamente")

class EditorItem:
    def __init__(self, editor: EditorCodigos, index):
        self.editor = editor
        self.index = index
        self.ventana = tk.Toplevel(editor.ventana)
        self.ventana.title(f"Editar Item")
        
        # TAMAÑO ESPECÍFICO PARA VENTANA DE EDITAR
        self.ventana.geometry("500x400")
        self.ventana.resizable(False, False)
        
        self.ventana.configure(bg="#FFFFFF")
        
        # Centrar ventana
        self.ventana.transient(editor.ventana)
        self.ventana.grab_set()
        
        self.centrar_ventana()

        row = editor.df_codigos_cumple.iloc[index]

        # Frame principal
        main_frame = tk.Frame(self.ventana, bg="#FFFFFF", padx=20, pady=20)
        main_frame.pack(fill="both", expand=True)
        
        tk.Label(main_frame, text=f"EDITAR ITEM: {row['ITEM']}", font=("Segoe UI", 14, "bold"),
                 bg="#FFFFFF", fg="#282828").pack(pady=10)

        # Campos de entrada
        campos = [
            ("OBSERVACIONES:", "obs_entry", row["OBSERVACIONES"]),
            ("CRITERIO:", "crit_entry", self.format_value(row["CRITERIO"]))
        ]
        
        for label_text, attr_name, value in campos:
            frame = tk.Frame(main_frame, bg="#FFFFFF")
            frame.pack(fill="x", pady=8)
            
            tk.Label(frame, text=label_text, bg="#FFFFFF", fg="#282828",
                     font=("Segoe UI", 10)).pack(anchor="w")
            
            entry = tk.Entry(frame, font=("Segoe UI", 10), bg="#FFFFFF", 
                            fg="#282828", relief="solid", bd=1)
            entry.insert(0, value)
            entry.pack(fill="x", pady=5)
            
            setattr(self, attr_name, entry)

        # Botones
        btn_frame = tk.Frame(main_frame, bg="#FFFFFF")
        btn_frame.pack(pady=20)
        
        tk.Button(btn_frame, text="💾 Guardar", bg="#ECD925", fg="#282828",
                  font=("Segoe UI", 10, "bold"), command=self.guardar_cambios,
                  relief="flat", padx=20, pady=8).pack(side="left", padx=10)
                  
        tk.Button(btn_frame, text="❌ Cancelar", bg="#E0E0E0", fg="#282828",
                  font=("Segoe UI", 10), command=self.ventana.destroy,
                  relief="flat", padx=20, pady=8).pack(side="left", padx=10)

    def format_value(self, value):
        """Formatea el valor para mostrar vacío en lugar de nan"""
        if pd.isna(value) or str(value).lower() == 'nan':
            return ""
        return value

    def centrar_ventana(self):
        """Centra la ventana en la pantalla"""
        self.ventana.update_idletasks()
        ancho = self.ventana.winfo_width()
        alto = self.ventana.winfo_height()
        x = (self.ventana.winfo_screenwidth() // 2) - (ancho // 2)
        y = (self.ventana.winfo_screenheight() // 2) - (alto // 2)
        self.ventana.geometry(f'+{x}+{y}')

    def guardar_cambios(self):
        observaciones = self.obs_entry.get().strip()
        criterio = self.crit_entry.get().strip()
        
        # Si la observación es "CUMPLE", asegurar que el criterio esté vacío
        if observaciones.upper() == "CUMPLE":
            criterio = ""
        
        self.editor.df_codigos_cumple.at[self.index, "OBSERVACIONES"] = observaciones
        self.editor.df_codigos_cumple.at[self.index, "CRITERIO"] = criterio
        
        # Guardar cambios (esto guardará en la ruta del dashboard)
        self.editor.guardar_datos()
        
        self.ventana.destroy()
        messagebox.showinfo("Éxito", "Cambios guardados correctamente")

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    archivo_excel = recurso_path("datos/codigos_cumple.xlsx")
    archivo_json = recurso_path("datos/codigos_cumple.json")
  # Ahora ambos apuntan al mismo archivo
    app = EditorCodigos(root, archivo_excel, archivo_json)
    root.mainloop()