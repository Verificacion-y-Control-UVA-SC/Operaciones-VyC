"""
Diálogo para gestionar la exportación e importación de bases de datos
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from data_manager import DataManager
import os

class DatabaseManagerDialog:
    def __init__(self, parent):
        self.parent = parent
        self.data_manager = DataManager()
        
        # Crear ventana de diálogo
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Gestor de Bases de Datos")
        self.dialog.geometry("700x600")
        self.dialog.resizable(True, True)  # Hacer redimensionable
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Centrar la ventana
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (700 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (600 // 2)
        self.dialog.geometry(f"700x600+{x}+{y}")
        
        self.create_widgets()
        self.update_status()
    
    def create_widgets(self):
        """Crear los widgets del diálogo"""
        # Frame principal
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill="both", expand=True)
        
        # Título
        title_label = ttk.Label(main_frame, text="Gestor de Bases de Datos", 
                               font=("Segoe UI", 16, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Frame para información de estado
        status_frame = ttk.LabelFrame(main_frame, text="Estado Actual", padding="10")
        status_frame.pack(fill="x", pady=(0, 20))
        
        # Información de bases de datos
        self.base_general_label = ttk.Label(status_frame, text="Base General: Cargando...", font=("Segoe UI", 10))
        self.base_general_label.pack(anchor="w", pady=2)
        
        self.inspeccion_label = ttk.Label(status_frame, text="Inspección: Cargando...", font=("Segoe UI", 10))
        self.inspeccion_label.pack(anchor="w", pady=2)
        
        self.historial_label = ttk.Label(status_frame, text="Historial: Cargando...", font=("Segoe UI", 10))
        self.historial_label.pack(anchor="w", pady=2)
        
        # Frame para exportación
        export_frame = ttk.LabelFrame(main_frame, text="Exportar a Excel", padding="10")
        export_frame.pack(fill="x", pady=(0, 20))
        
        # Botones de exportación
        export_buttons_frame = ttk.Frame(export_frame)
        export_buttons_frame.pack(fill="x")
        
        ttk.Button(export_buttons_frame, text="📤 Exportar Base General", 
                  command=self.export_base_general).pack(side="left", padx=(0, 10), pady=5)
        
        ttk.Button(export_buttons_frame, text="📤 Exportar Inspección", 
                  command=self.export_inspeccion).pack(side="left", padx=(0, 10), pady=5)
        
        ttk.Button(export_buttons_frame, text="📤 Exportar Historial", 
                  command=self.export_historial).pack(side="left", padx=(0, 10), pady=5)
        
        ttk.Button(export_buttons_frame, text="📤 Exportar Todo", 
                  command=self.export_all).pack(side="left", pady=5)
        
        # Frame para importación
        import_frame = ttk.LabelFrame(main_frame, text="Importar desde Excel", padding="10")
        import_frame.pack(fill="x", pady=(0, 20))
        
        # Botones de importación
        import_buttons_frame = ttk.Frame(import_frame)
        import_buttons_frame.pack(fill="x")
        
        ttk.Button(import_buttons_frame, text="📥 Importar Base General", 
                  command=self.import_base_general).pack(side="left", padx=(0, 10), pady=5)
        
        ttk.Button(import_buttons_frame, text="📥 Importar Inspección", 
                  command=self.import_inspeccion).pack(side="left", padx=(0, 10), pady=5)
        
        ttk.Button(import_buttons_frame, text="📥 Importar Historial", 
                  command=self.import_historial).pack(side="left", padx=(0, 10), pady=5)
        
        # Frame para acciones adicionales
        actions_frame = ttk.LabelFrame(main_frame, text="Acciones Adicionales", padding="10")
        actions_frame.pack(fill="x", pady=(0, 20))
        
        # Botones de acciones adicionales (más simples y visibles)
        ttk.Button(actions_frame, text="🔄 Recargar Datos", 
                  command=self.reload_data).pack(pady=5)
        
        ttk.Button(actions_frame, text="🗑️ Limpiar Historial", 
                  command=self.clear_historial).pack(pady=5)
        
        # Frame para botones de cierre
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=(20, 0))
        
        ttk.Button(button_frame, text="Cerrar", command=self.dialog.destroy).pack(side="right")
        
        # Configurar eventos
        self.dialog.protocol("WM_DELETE_WINDOW", self.dialog.destroy)
    
    def update_status(self):
        """Actualizar información de estado"""
        info = self.data_manager.get_data_info()
        
        # Base General
        if info['base_general']['exists']:
            self.base_general_label.config(
                text=f"Base General: ✅ {info['base_general']['records']} registros",
                foreground="green"
            )
        else:
            self.base_general_label.config(
                text="Base General: ❌ No encontrada",
                foreground="red"
            )
        
        # Inspección
        if info['inspeccion']['exists']:
            self.inspeccion_label.config(
                text=f"Inspección: ✅ {info['inspeccion']['records']} registros",
                foreground="green"
            )
        else:
            self.inspeccion_label.config(
                text="Inspección: ❌ No encontrada",
                foreground="red"
            )
        
        # Historial
        if info['historial']['exists']:
            self.historial_label.config(
                text=f"Historial: ✅ {info['historial']['records']} registros",
                foreground="green"
            )
        else:
            self.historial_label.config(
                text="Historial: ❌ No encontrado",
                foreground="red"
            )
    
    def export_base_general(self):
        """Exportar base general a Excel"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")],
            title="Exportar Base General",
            initialfile="BASE_GENERAL.xlsx"
        )
        
        if file_path:
            if self.data_manager.export_base_general_to_excel(file_path):
                messagebox.showinfo("Éxito", f"Base general exportada a:\n{file_path}")
            else:
                messagebox.showerror("Error", "No se pudo exportar la base general")
    
    def export_inspeccion(self):
        """Exportar inspección a Excel"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")],
            title="Exportar Inspección",
            initialfile="INSPECCION.xlsx"
        )
        
        if file_path:
            if self.data_manager.export_inspeccion_to_excel(file_path):
                messagebox.showinfo("Éxito", f"Inspección exportada a:\n{file_path}")
            else:
                messagebox.showerror("Error", "No se pudo exportar la inspección")
    
    def export_historial(self):
        """Exportar historial a Excel"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")],
            title="Exportar Historial",
            initialfile="HISTORIAL.xlsx"
        )
        
        if file_path:
            if self.data_manager.export_historial_to_excel(file_path):
                messagebox.showinfo("Éxito", f"Historial exportado a:\n{file_path}")
            else:
                messagebox.showerror("Error", "No se pudo exportar el historial")
    
    def export_all(self):
        """Exportar todas las bases de datos"""
        folder_path = filedialog.askdirectory(title="Seleccionar carpeta para exportar")
        
        if folder_path:
            success_count = 0
            
            # Exportar base general
            base_path = os.path.join(folder_path, "BASE_GENERAL.xlsx")
            if self.data_manager.export_base_general_to_excel(base_path):
                success_count += 1
            
            # Exportar inspección
            inspeccion_path = os.path.join(folder_path, "INSPECCION.xlsx")
            if self.data_manager.export_inspeccion_to_excel(inspeccion_path):
                success_count += 1
            
            # Exportar historial
            historial_path = os.path.join(folder_path, "HISTORIAL.xlsx")
            if self.data_manager.export_historial_to_excel(historial_path):
                success_count += 1
            
            if success_count == 3:
                messagebox.showinfo("Éxito", f"Todas las bases de datos exportadas a:\n{folder_path}")
            elif success_count > 0:
                messagebox.showwarning("Parcial", f"Se exportaron {success_count}/3 bases de datos a:\n{folder_path}")
            else:
                messagebox.showerror("Error", "No se pudo exportar ninguna base de datos")
    
    def import_base_general(self):
        """Importar base general desde Excel"""
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Base General",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        
        if file_path:
            if messagebox.askyesno("Confirmar", "¿Estás seguro de que quieres importar la base general?\nEsto sobrescribirá los datos actuales."):
                if self.data_manager.import_base_general_from_excel(file_path):
                    messagebox.showinfo("Éxito", "Base general importada correctamente")
                    self.update_status()
                else:
                    messagebox.showerror("Error", "No se pudo importar la base general")
    
    def import_inspeccion(self):
        """Importar inspección desde Excel"""
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Inspección",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        
        if file_path:
            if messagebox.askyesno("Confirmar", "¿Estás seguro de que quieres importar la inspección?\nEsto sobrescribirá los datos actuales."):
                if self.data_manager.import_inspeccion_from_excel(file_path):
                    messagebox.showinfo("Éxito", "Inspección importada correctamente")
                    self.update_status()
                else:
                    messagebox.showerror("Error", "No se pudo importar la inspección")
    
    def import_historial(self):
        """Importar historial desde Excel"""
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Historial",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        
        if file_path:
            if messagebox.askyesno("Confirmar", "¿Estás seguro de que quieres importar el historial?\nEsto sobrescribirá los datos actuales."):
                if self.data_manager.import_historial_from_excel(file_path):
                    messagebox.showinfo("Éxito", "Historial importado correctamente")
                    self.update_status()
                else:
                    messagebox.showerror("Error", "No se pudo importar el historial")
    
    def reload_data(self):
        """Recargar datos desde archivos"""
        try:
            # Recargar el gestor de datos
            self.data_manager = DataManager()
            self.update_status()
            messagebox.showinfo("Éxito", "Datos recargados correctamente")
        except Exception as e:
            messagebox.showerror("Error", f"Error al recargar datos: {e}")
    
    def clear_historial(self):
        """Limpiar historial"""
        if messagebox.askyesno("Confirmar", "¿Estás seguro de que quieres limpiar el historial?\nEsta acción no se puede deshacer."):
            try:
                self.data_manager.historial = []
                self.data_manager._save_historial()
                self.update_status()
                messagebox.showinfo("Éxito", "Historial limpiado correctamente")
            except Exception as e:
                messagebox.showerror("Error", f"Error al limpiar historial: {e}")
    
    def show(self):
        """Mostrar el diálogo"""
        self.dialog.wait_window()
