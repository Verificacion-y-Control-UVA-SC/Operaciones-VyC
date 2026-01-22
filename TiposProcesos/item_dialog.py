"""
Diálogo para agregar información de nuevos ítems
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, Any, Optional
import pandas as pd

class ItemDialog:
    def __init__(self, parent, item: int, existing_info: Dict[str, Any] = None):
        self.parent = parent
        self.item = item
        self.existing_info = existing_info or {}
        self.result = None
        
        # Crear ventana de diálogo
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Agregar Ítem Nuevo: {item}")
        self.dialog.geometry("500x400")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Centrar la ventana
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (500 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (400 // 2)
        self.dialog.geometry(f"500x400+{x}+{y}")
        
        self.create_widgets()
        self.load_existing_info()
    
    def create_widgets(self):
        """Crear los widgets del diálogo"""
        # Frame principal
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill="both", expand=True)
        
        # Título
        title_label = ttk.Label(main_frame, text=f"Información del Ítem: {self.item}", 
                               font=("Segoe UI", 14, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Frame para campos
        fields_frame = ttk.Frame(main_frame)
        fields_frame.pack(fill="both", expand=True)
        
        # Tipo de Proceso
        ttk.Label(fields_frame, text="Tipo de Proceso:", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 5))
        self.tipo_var = tk.StringVar()
        tipo_combo = ttk.Combobox(fields_frame, textvariable=self.tipo_var, 
                                 values=["ADHERIBLE", "COSTURA", "SIN NORMA"], 
                                 state="readonly", width=30)
        tipo_combo.pack(fill="x", pady=(0, 15))
        tipo_combo.set("SIN NORMA")
        
        # Norma
        ttk.Label(fields_frame, text="Norma:", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 5))
        self.norma_var = tk.StringVar()
        norma_entry = ttk.Entry(fields_frame, textvariable=self.norma_var, width=50)
        norma_entry.pack(fill="x", pady=(0, 15))
        
        # Descripción
        ttk.Label(fields_frame, text="Descripción:", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 5))
        self.desc_var = tk.StringVar()
        self.desc_entry = ttk.Entry(fields_frame, textvariable=self.desc_var, width=50)
        self.desc_entry.pack(fill="x", pady=(0, 15))
        
        # Criterio
        ttk.Label(fields_frame, text="Criterio:", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 5))
        self.criterio_var = tk.StringVar()
        criterio_combo = ttk.Combobox(fields_frame, textvariable=self.criterio_var,
                                     values=["CUMPLE", "NO CUMPLE", "PENDIENTE", "REVISADO"], 
                                     state="readonly", width=30)
        criterio_combo.pack(fill="x", pady=(0, 15))
        criterio_combo.set("PENDIENTE")
        
        # Frame para botones
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=(20, 0))
        
        # Botones
        ttk.Button(button_frame, text="Guardar", command=self.save_item, 
                  style="Accent.TButton").pack(side="right", padx=(10, 0))
        ttk.Button(button_frame, text="Cancelar", command=self.cancel).pack(side="right")
        
        # Configurar eventos
        self.dialog.protocol("WM_DELETE_WINDOW", self.cancel)
        norma_entry.focus()
    
    def load_existing_info(self):
        """Cargar información existente si la hay"""
        if self.existing_info:
            if 'tipo_proceso' in self.existing_info:
                self.tipo_var.set(self.existing_info['tipo_proceso'])
            if 'norma' in self.existing_info:
                self.norma_var.set(self.existing_info['norma'])
            if 'descripcion' in self.existing_info:
                self.desc_var.set(self.existing_info['descripcion'])
                # Hacer el campo de descripción de solo lectura si ya existe
                self.desc_entry.config(state='readonly')
            if 'criterio' in self.existing_info:
                self.criterio_var.set(self.existing_info['criterio'])
    
    def validate_fields(self) -> bool:
        """Validar que los campos requeridos estén llenos"""
        if not self.tipo_var.get().strip():
            messagebox.showerror("Error", "El tipo de proceso es requerido.")
            return False
        
        if not self.norma_var.get().strip():
            messagebox.showerror("Error", "La norma es requerida.")
            return False
        
        # Solo validar descripción si no existe información previa
        if not self.existing_info.get('descripcion') and not self.desc_var.get().strip():
            messagebox.showerror("Error", "La descripción es requerida.")
            return False
        
        if not self.criterio_var.get().strip():
            messagebox.showerror("Error", "El criterio es requerido.")
            return False
        
        return True
    
    def save_item(self):
        """Guardar la información del ítem"""
        if not self.validate_fields():
            return
        
        self.result = {
            'item': str(self.item),
            'tipo_proceso': self.tipo_var.get().strip(),
            'norma': self.norma_var.get().strip(),
            'descripcion': self.desc_var.get().strip(),
            'criterio': self.criterio_var.get().strip()
        }
        
        self.dialog.destroy()
    
    def cancel(self):
        """Cancelar la operación"""
        self.result = None
        self.dialog.destroy()
    
    def get_result(self) -> Optional[Dict[str, str]]:
        """Obtener el resultado del diálogo"""
        self.dialog.wait_window()
        return self.result

class BatchItemDialog:
    """Diálogo para procesar múltiples ítems nuevos"""
    
    def __init__(self, parent, new_items: list, df_reporte):
        self.parent = parent
        self.new_items = new_items
        self.df_reporte = df_reporte
        self.results = {}
        
        # Crear ventana de diálogo
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Procesar {len(new_items)} Ítems Nuevos")
        self.dialog.geometry("600x500")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Centrar la ventana
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (600 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (500 // 2)
        self.dialog.geometry(f"600x500+{x}+{y}")
        
        self.create_widgets()
    
    def create_widgets(self):
        """Crear los widgets del diálogo"""
        # Frame principal
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill="both", expand=True)
        
        # Título
        title_label = ttk.Label(main_frame, text=f"Procesar {len(self.new_items)} Ítems Nuevos", 
                               font=("Segoe UI", 14, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Lista de ítems
        ttk.Label(main_frame, text="Ítems nuevos detectados:", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        
        # Frame para la lista
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill="both", expand=True, pady=(10, 20))
        
        # Crear Treeview para mostrar ítems
        columns = ('item', 'status')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=10)
        
        self.tree.heading('item', text='Ítem')
        self.tree.heading('status', text='Estado')
        
        self.tree.column('item', width=100)
        self.tree.column('status', width=150)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Agregar ítems a la lista
        for item in self.new_items:
            self.tree.insert('', 'end', values=(item, 'Pendiente'))
        
        # Frame para botones
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=(20, 0))
        
        # Botones
        ttk.Button(button_frame, text="Procesar Todos", command=self.process_all, 
                  style="Accent.TButton").pack(side="right", padx=(10, 0))
        ttk.Button(button_frame, text="Procesar Seleccionado", command=self.process_selected).pack(side="right", padx=(10, 0))
        ttk.Button(button_frame, text="Cancelar", command=self.cancel).pack(side="right")
        
        # Configurar eventos
        self.dialog.protocol("WM_DELETE_WINDOW", self.cancel)
        self.tree.bind('<Double-1>', self.on_item_double_click)
    
    def on_item_double_click(self, event):
        """Manejar doble clic en un ítem"""
        selection = self.tree.selection()
        if selection:
            item_id = selection[0]
            item_value = self.tree.item(item_id)['values'][0]
            self.process_single_item(item_value)
    
    def process_selected(self):
        """Procesar ítem seleccionado"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Advertencia", "Por favor selecciona un ítem.")
            return
        
        item_id = selection[0]
        item_value = self.tree.item(item_id)['values'][0]
        self.process_single_item(item_value)
    
    def process_single_item(self, item):
        """Procesar un ítem individual"""
        # Obtener información del reporte para este ítem
        item_info = self.get_item_info_from_report(item)
        
        # Crear diálogo individual
        dialog = ItemDialog(self.dialog, item, item_info)
        result = dialog.get_result()
        
        if result:
            self.results[item] = result
            # Actualizar estado en la lista
            try:
                for item_id in self.tree.get_children():
                    if self.tree.item(item_id)['values'][0] == item:
                        self.tree.set(item_id, 'status', 'Completado')
                        break
            except tk.TclError:
                # Si el TreeView ya no existe, continuar
                pass
        else:
            # Marcar como cancelado
            try:
                for item_id in self.tree.get_children():
                    if self.tree.item(item_id)['values'][0] == item:
                        self.tree.set(item_id, 'status', 'Cancelado')
                        break
            except tk.TclError:
                # Si el TreeView ya no existe, continuar
                pass
    
    def process_all(self):
        """Procesar todos los ítems"""
        for item in self.new_items:
            if item not in self.results:
                self.process_single_item(item)
    
    def get_item_info_from_report(self, item):
        """Extraer información de un ítem desde el reporte"""
        info = {}
        
        try:
            # Buscar el ítem en el reporte
            item_rows = self.df_reporte[self.df_reporte['Num.Parte'].astype(str) == str(item)]
            
            if not item_rows.empty:
                row = item_rows.iloc[0]
                
                # Extraer norma si existe
                if 'NOMs' in row and pd.notna(row['NOMs']) and str(row['NOMs']).strip():
                    norma_value = str(row['NOMs']).strip()
                    if norma_value and norma_value != '0' and norma_value != 'nan' and norma_value != 'None':
                        info['norma'] = norma_value
                
                # Buscar descripción en múltiples columnas (más exhaustivo)
                descripcion_columns = [
                    'DESCRIPCION', 'DESCRIPCIÓN', 'DESCRIPTION', 'DESCRIP', 'PRODUCTO', 'NOMBRE',
                    'NOMBRE PRODUCTO', 'DESCRIPCIÓN DEL PRODUCTO', 'NOMBRE DEL PRODUCTO', 
                    'TITULO', 'TÍTULO', 'NOMBRE ARTICULO', 'DESCRIPCION ARTICULO'
                ]
                
                for col in descripcion_columns:
                    if col in row and pd.notna(row[col]) and str(row[col]).strip():
                        desc_value = str(row[col]).strip()
                        if desc_value and desc_value != 'nan' and desc_value != 'None' and len(desc_value) > 2:
                            info['descripcion'] = desc_value
                            print(f"Debug - Descripción encontrada para ítem {item} en columna '{col}': '{desc_value}'")
                            break
                
                # Si no se encontró en las columnas específicas, buscar en cualquier columna que contenga "descrip" o "nombre"
                if 'descripcion' not in info:
                    for col in row.index:
                        if any(palabra in col.upper() for palabra in ['DESCRIP', 'NOMBRE', 'PRODUCTO', 'TITULO']):
                            if pd.notna(row[col]) and str(row[col]).strip():
                                desc_value = str(row[col]).strip()
                                if desc_value and desc_value != 'nan' and desc_value != 'None' and len(desc_value) > 2:
                                    info['descripcion'] = desc_value
                                    print(f"Debug - Descripción encontrada para ítem {item} en columna '{col}': '{desc_value}'")
                                    break
                
                # Determinar tipo de proceso basado en la norma
                if 'norma' in info:
                    norma = info['norma'].upper()
                    if any(n in norma for n in ['NOM-004', 'NOM004', '004']):
                        info['tipo_proceso'] = 'COSTURA'
                    elif any(n in norma for n in ['NOM-050', 'NOM-015', 'NOM-024', 'NOM-141']):
                        info['tipo_proceso'] = 'ADHERIBLE'
                    else:
                        info['tipo_proceso'] = 'SIN NORMA'
                else:
                    info['tipo_proceso'] = 'SIN NORMA'
                
                # Criterio por defecto
                info['criterio'] = 'PENDIENTE'
                
                print(f"Debug - Información extraída para ítem {item}: {info}")
        
        except Exception as e:
            print(f"Error extrayendo información del reporte para ítem {item}: {e}")
        
        return info
    
    def cancel(self):
        """Cancelar la operación"""
        self.results = {}
        self.dialog.destroy()
    
    def get_results(self) -> Dict[int, Dict[str, str]]:
        """Obtener los resultados del procesamiento"""
        self.dialog.wait_window()
        return self.results
