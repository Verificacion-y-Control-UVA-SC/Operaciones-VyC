import os
import json
import pandas as pd
import customtkinter as ctk
import re
from tkinter import filedialog, messagebox, ttk
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from typing import List, Dict, Any, Optional
import queue

# Configuración de estilos
STYLE = {
    "primario": "#ECD925",
    "secundario": "#282828",
    "fondo": "#F8F9FA",
    "surface": "#FFFFFF",
    "texto_oscuro": "#282828",
    "texto_claro": "#4b4b4b",
    "borde": "#BDC3C7",
    "exito": "#27AE60",
    "peligro": "#d74a3d",
    "hover_primario": "#ECD925",
    "button_hover": "#4b4b4b"
}

@dataclass
class ItemFactura:
    """Clase ligera para almacenar datos de factura"""
    upc: str
    factura: str
    descripcion: str
    cantidad_original: str
    cantidad_editada: str = ""
    escaneos: int = 0
    validacion: str = "No Llegó"
    observaciones: str = ""
    index: int = 0
    visible: bool = True
    modificado: bool = False  # Nuevo campo para rastrear cambios

@dataclass
class ItemLayout:
    """Clase ligera para almacenar datos de layout"""
    parte: str
    descripcion: str
    cantidad_original: str
    cantidad_editada: str = ""
    escaneos: int = 0
    validacion: str = "No Llegó"
    observaciones: str = ""
    index: int = 0
    visible: bool = True
    modificado: bool = False  # Nuevo campo para rastrear cambios

class TablaRapida(ttk.Treeview):
    """Treeview optimizado para máximo rendimiento"""
    def __init__(self, parent, columns, modo):
        super().__init__(parent, columns=columns, show='headings')
        
        self.modo = modo
        self.columnas = columns
        
        # Configuración con fuente más grande
        style = ttk.Style()
        style.configure("Treeview", font=("Inter", 11), rowheight=30)
        style.configure("Treeview.Heading", font=("Inter", 11, "bold"))
        
        # Configuración ultra rápida
        self.configure_columns()
        
        # Deshabilitar animaciones y efectos visuales
        self.configure(takefocus=False)
        
        # Cache para búsquedas rápidas
        self.cache_busqueda = {}
        
    def configure_columns(self):
        """Configura las columnas de forma optimizada"""
        if self.modo == "factura":
            widths = [140, 120, 250, 100, 100, 120, 200]
        else:
            widths = [140, 250, 100, 100, 120, 200]
        
        for i, col in enumerate(self.columnas):
            self.heading(col, text=col)
            self.column(col, width=widths[i], anchor='w' if i in [0, 1, 2, 6] else 'center')
    
    def insertar_items_batch(self, items):
        """Inserta múltiples items en lote para mejor rendimiento"""
        # Temporariamente deshabilitar actualizaciones
        self.config(displaycolumns='#all')
        
        # Insertar en lote
        valores = []
        for item in items:
            if self.modo == "factura":
                valores.append((
                    item.upc,
                    item.factura,
                    item.descripcion[:50] + "..." if len(item.descripcion) > 50 else item.descripcion,
                    item.cantidad_editada or item.cantidad_original,
                    str(item.escaneos),
                    item.validacion,
                    item.observaciones[:30] + "..." if len(item.observaciones) > 30 else item.observaciones
                ))
            else:
                valores.append((
                    item.parte,
                    item.descripcion[:50] + "..." if len(item.descripcion) > 50 else item.descripcion,
                    item.cantidad_editada or item.cantidad_original,
                    str(item.escaneos),
                    item.validacion,
                    item.observaciones[:30] + "..." if len(item.observaciones) > 30 else item.observaciones
                ))
        
        # Insertar todos a la vez
        for vals in valores:
            self.insert('', 'end', values=vals)
        
        # Rehabilitar actualizaciones
        self.update()

class EditorFacturacion(ctk.CTkToplevel):
    def __init__(self, parent, factura_data=None, layout_data=None, contador_escaneos=None):
        super().__init__(parent)
        self.parent = parent
        self.factura_data = factura_data or {"items": []}
        self.layout_data = layout_data or {"datos": []}
        self.contador_escaneos = contador_escaneos or {}
        
        # Determinar modo
        self.modo = "factura" if self.factura_data.get('items') else "layout"
        
        # Configuración de ventana
        self.title(f"Editor Ultra Rápido - {'FACTURA' if self.modo == 'factura' else 'LAYOUT'}")
        self.geometry("1000x500")  # Ligeramente más alto para fuente más grande
        self.configure(fg_color=STYLE["fondo"])
        
        # Hacer modal
        self.grab_set()
        
        # ✅ VARIABLES ULTRA OPTIMIZADAS
        self.items = []  # Lista de objetos ItemFactura/ItemLayout
        self.items_filtrados = []  # Índices de items visibles
        self.cambios = {}  # Diccionario rápido de cambios {codigo: cambios}
        self.cache_datos = {}  # Cache para acceso rápido
        
        # Para tracking de cambios
        self.hay_cambios_sin_guardar = False
        self.cambios_pendientes = []
        
        # Para threading
        self.executor = ThreadPoolExecutor(max_workers=2)
        self.carga_queue = queue.Queue()
        self.cancelar_carga = threading.Event()
        
        # Crear widgets
        self.crear_widgets_ultra_rapidos()
        
        # Cargar datos en background
        self.cargar_datos_background()
        
        self.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def crear_widgets_ultra_rapidos(self):
        """Crea todos los widgets optimizados"""
        main_frame = ctk.CTkFrame(self, fg_color=STYLE["fondo"])
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # Título
        titulo = ctk.CTkLabel(
            main_frame,
            text=f"EDITOR DE INFORMACIÓN - {'FACTURA' if self.modo == 'factura' else 'LAYOUT'}",
            font=("Inter", 20, "bold"),
            text_color=STYLE["texto_oscuro"]
        )
        titulo.pack(pady=(0, 15))
        
        # Barra superior de controles
        self.crear_barra_controles(main_frame)
        
        # Tabla principal
        self.crear_tabla_principal(main_frame)
        
        # Barra inferior de estadísticas
        self.crear_barra_estadisticas(main_frame)
    
    def crear_barra_controles(self, parent):
        """Crea la barra de controles superior"""
        controles_frame = ctk.CTkFrame(parent, fg_color=STYLE["fondo"], height=60)
        controles_frame.pack(fill="x", pady=(0, 10))
        
        # Lado izquierdo: Búsqueda y filtros
        izquierda_frame = ctk.CTkFrame(controles_frame, fg_color=STYLE["fondo"])
        izquierda_frame.pack(side="left", fill="x", expand=True)
        
        # Barra de búsqueda ultra rápida
        self.crear_barra_busqueda(izquierda_frame)
        
        # Lado derecho: Botones de acción
        derecha_frame = ctk.CTkFrame(controles_frame, fg_color=STYLE["fondo"])
        derecha_frame.pack(side="right", padx=(0, 10))
        
        # Botón para guardar
        self.btn_guardar = ctk.CTkButton(
            derecha_frame,
            text="💾 Guardar Todo",
            command=self.guardar_todo_rapido,
            width=120,
            height=40,
            fg_color=STYLE["primario"],
            hover_color=STYLE["hover_primario"],
            text_color=STYLE["texto_oscuro"],
            font=("Inter", 12, "bold")
        )
        self.btn_guardar.pack(side="left", padx=5)
        
        # Botón para exportar Excel
        self.btn_exportar = ctk.CTkButton(
            derecha_frame,
            text="📊 Exportar Excel",
            command=self.exportar_excel_rapido,
            width=120,
            height=40,
            fg_color=STYLE["exito"],
            hover_color="#2ecc71",
            text_color="white",
            font=("Inter", 12, "bold")
        )
        self.btn_exportar.pack(side="left", padx=5)
    
    def crear_barra_busqueda(self, parent):
        """Crea barra de búsqueda optimizada"""
        busqueda_frame = ctk.CTkFrame(parent, fg_color=STYLE["fondo"])
        busqueda_frame.pack(fill="x", pady=5)
        
        # Etiqueta
        ctk.CTkLabel(
            busqueda_frame,
            text="🔍 Buscar:",
            font=("Inter", 14, "bold"),
            text_color=STYLE["texto_oscuro"]
        ).pack(side="left", padx=(0, 10))
        
        # Campo de búsqueda principal
        self.entry_busqueda = ctk.CTkEntry(
            busqueda_frame,
            placeholder_text="UPC, Parte, Descripción...",
            width=300,
            height=40,
            font=("Inter", 12),
            corner_radius=8
        )
        self.entry_busqueda.pack(side="left", padx=(0, 10))
        self.entry_busqueda.bind('<KeyRelease>', self.filtrar_rapido)
        
        # Selector de campo de búsqueda
        self.combo_campo = ctk.CTkComboBox(
            busqueda_frame,
            values=["Todo", "Código", "Descripción", "Factura"],
            width=120,
            height=40,
            font=("Inter", 12),
            state="readonly"
        )
        self.combo_campo.set("Todo")
        self.combo_campo.pack(side="left", padx=(0, 10))
    
    def crear_tabla_principal(self, parent):
        """Crea la tabla principal ultra rápida"""
        # Frame para la tabla con scroll
        tabla_frame = ctk.CTkFrame(parent, fg_color=STYLE["surface"], corner_radius=10)
        tabla_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        # Definir columnas según modo
        if self.modo == "factura":
            columnas = ("UPC", "FACTURA", "DESCRIPCIÓN", "CANTIDAD", "ESCANEOS", "VALIDACIÓN")
        else:
            columnas = ("PARTE", "DESCRIPCIÓN", "CANTIDAD", "ESCANEOS", "VALIDACIÓN")
        
        # Crear Treeview optimizado con fuente más grande
        self.tabla = TablaRapida(tabla_frame, columnas, self.modo)
        
        # Configurar scrollbars
        scroll_y = ttk.Scrollbar(tabla_frame, orient="vertical", command=self.tabla.yview)
        scroll_x = ttk.Scrollbar(tabla_frame, orient="horizontal", command=self.tabla.xview)
        self.tabla.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        # Posicionar widgets
        self.tabla.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        
        # Configurar grid
        tabla_frame.grid_rowconfigure(0, weight=1)
        tabla_frame.grid_columnconfigure(0, weight=1)
        
        # Bind para edición rápida
        self.tabla.bind('<Double-1>', self.editar_celda_rapido)
        self.tabla.bind('<Return>', self.guardar_edicion_rapida)
        self.tabla.bind('<Escape>', self.cancelar_edicion)
        
        # Variables para edición
        self.celda_editando = None
        self.widget_edicion = None
    
    def crear_barra_estadisticas(self, parent):
        """Crea barra de estadísticas en la parte inferior"""
        stats_frame = ctk.CTkFrame(parent, fg_color=STYLE["surface"], height=50, corner_radius=8)
        stats_frame.pack(fill="x", pady=(5, 0))
        
        # Estadísticas en 3 columnas
        col1 = ctk.CTkFrame(stats_frame, fg_color=STYLE["surface"])
        col1.pack(side="left", fill="x", expand=True, padx=20, pady=10)
        
        col2 = ctk.CTkFrame(stats_frame, fg_color=STYLE["surface"])
        col2.pack(side="left", fill="x", expand=True, padx=20, pady=10)
        
        col3 = ctk.CTkFrame(stats_frame, fg_color=STYLE["surface"])
        col3.pack(side="left", fill="x", expand=True, padx=20, pady=10)
        
        # Columna 1: Totales
        self.label_total = ctk.CTkLabel(
            col1,
            text="Total: 0",
            font=("Inter", 12, "bold"),
            text_color=STYLE["texto_oscuro"]
        )
        self.label_total.pack(anchor="w")
        
        self.label_escaneados = ctk.CTkLabel(
            col1,
            text="Escaneados: 0",
            font=("Inter", 12),
            text_color=STYLE["exito"]
        )
        self.label_escaneados.pack(anchor="w")
        
        # Columna 2: Pendientes
        self.label_pendientes = ctk.CTkLabel(
            col2,
            text="Pendientes: 0",
            font=("Inter", 12),
            text_color=STYLE["peligro"]
        )
        self.label_pendientes.pack(anchor="w")
        
        self.label_validados = ctk.CTkLabel(
            col2,
            text="Validados: 0",
            font=("Inter", 12),
            text_color=STYLE["texto_claro"]
        )
        self.label_validados.pack(anchor="w")
        
        # Columna 3: Progreso
        self.label_progreso = ctk.CTkLabel(
            col3,
            text="Progreso: 0%",
            font=("Inter", 12),
            text_color=STYLE["secundario"]
        )
        self.label_progreso.pack(anchor="w")
        
        self.progressbar = ctk.CTkProgressBar(col3, width=200)
        self.progressbar.pack(anchor="w", pady=(5, 0))
        self.progressbar.set(0)
    
    def cargar_datos_background(self):
        """Carga datos en background usando threading"""
        # Mostrar mensaje de carga
        self.label_total.configure(text="Cargando datos...")
        
        # Ejecutar en thread
        self.executor.submit(self._cargar_datos_thread)
    
    def _cargar_datos_thread(self):
        """Thread para cargar datos"""
        try:
            if self.modo == "factura":
                items_raw = self.factura_data.get('items', [])
                self.items = self.procesar_factura_rapido(items_raw)
            else:
                items_raw = self.layout_data.get('datos', [])
                self.items = self.procesar_layout_rapido(items_raw)
            
            # Actualizar UI en thread principal
            self.after(0, self._carga_completada)
            
        except Exception as e:
            print(f"Error en carga: {e}")
            self.after(0, lambda: self.mostrar_error(f"Error cargando datos: {e}"))
    
    def procesar_factura_rapido(self, items_raw):
        """Procesa datos de factura de forma ultra rápida"""
        items = []
        
        for idx, item in enumerate(items_raw):
            # Extraer UPC rápidamente
            upc = self.extraer_upc_rapido(item)
            if not upc:
                continue
            
            # Obtener escaneos para este UPC
            escaneos = int(self.contador_escaneos.get(upc, 0) or 0)
            
            # Determinar validación automática: Llegó si tiene escaneos, No Llegó si no
            validacion_auto = "Llegó" if escaneos > 0 else "No Llegó"
            
            # Crear objeto ItemFactura
            item_obj = ItemFactura(
                upc=upc,
                factura=item.get('FACTURA', '') or self.factura_data.get('nombre_sin_ext', ''),
                descripcion=item.get('DESC. FACTURA', '') or item.get('DESC_FACTURA', ''),
                cantidad_original=item.get('CANTIDAD EN VU', '0') or item.get('CANTIDAD_EN_VU', '0'),
                escaneos=escaneos,
                validacion=validacion_auto,  # Validación automática según escaneos
                index=idx
            )
            
            # Aplicar cambios guardados si existen
            if hasattr(self.parent, 'cambios_factura'):
                for cambio in self.parent.cambios_factura:
                    if cambio.get('UPC') == upc:
                        item_obj.cantidad_editada = cambio.get('Cantidad_Factura_Editada', '')
                        # Solo mantener validación personalizada si no es la automática
                        if cambio.get('Validacion') != validacion_auto:
                            item_obj.validacion = cambio.get('Validacion', validacion_auto)
                        item_obj.observaciones = cambio.get('Observaciones', '')
                        break
            
            items.append(item_obj)
        
        return items
    
    def extraer_upc_rapido(self, item):
        """Extrae UPC de forma optimizada"""
        if not isinstance(item, dict):
            return None
        
        # Preferir campo explícito 'UPC' si existe (evita diferencias con '# ORDEN - ITEM')
        if 'UPC' in item and item.get('UPC'):
            valor_upc = str(item.get('UPC') or '')
            match = re.search(r'\d{6,14}', valor_upc)
            if match:
                return match.group()

        # Buscar en otros campos comunes
        campos_upc = ['# ORDEN - ITEM', '# ORDEN - ITEM ', 'ORDEN - ITEM', 'ORDEN-ITEM']
        for campo in campos_upc:
            if campo in item:
                valor = str(item[campo] or '')
                if valor:
                    # Buscar secuencias de dígitos
                    match = re.search(r'\d{6,14}', valor)
                    if match:
                        return match.group()

        return None
    
    def procesar_layout_rapido(self, items_raw):
        """Procesa datos de layout de forma ultra rápida"""
        items = []
        
        for idx, item in enumerate(items_raw):
            parte = item.get('Parte', '')
            if not parte:
                continue
            
            # Obtener escaneos para esta parte
            escaneos = int(self.contador_escaneos.get(parte, 0) or 0)
            
            # Determinar validación automática
            validacion_auto = "Llegó" if escaneos > 0 else "No Llegó"
            
            # Crear objeto ItemLayout
            item_obj = ItemLayout(
                parte=parte,
                descripcion=item.get('Descripción del producto', ''),
                cantidad_original=item.get('Cantidad', '0'),
                escaneos=escaneos,
                validacion=validacion_auto,  # Validación automática según escaneos
                index=idx
            )
            
            # Aplicar cambios guardados si existen
            if hasattr(self.parent, 'cambios_layout'):
                for cambio in self.parent.cambios_layout:
                    if cambio.get('Parte') == parte:
                        item_obj.cantidad_editada = cambio.get('Cantidad_Editada', '')
                        # Solo mantener validación personalizada si no es la automática
                        if cambio.get('Validacion') != validacion_auto:
                            item_obj.validacion = cambio.get('Validacion', validacion_auto)
                        item_obj.observaciones = cambio.get('Observaciones', '')
                        break
            
            items.append(item_obj)
        
        return items
    
    def _carga_completada(self):
        """Se llama cuando la carga se completa"""
        # Insertar items en tabla
        self.tabla.insertar_items_batch(self.items)
        
        # Actualizar estadísticas
        self.actualizar_estadisticas()
        
        # Actualizar UI
        self.label_total.configure(text=f"Total: {len(self.items)}")
        
        # Habilitar botones
        self.btn_guardar.configure(state="normal")
        self.btn_exportar.configure(state="normal")
        
        # Actualizar progreso
        escaneados = sum(1 for item in self.items if item.escaneos > 0)
        self.progressbar.set(escaneados / max(len(self.items), 1))
    
    def actualizar_estadisticas(self):
        """Actualiza todas las estadísticas de una vez"""
        if not self.items:
            return
        
        total = len(self.items)
        escaneados = sum(1 for item in self.items if item.escaneos > 0)
        pendientes = total - escaneados
        validados = sum(1 for item in self.items if item.validacion == "Llegó")
        parciales = sum(1 for item in self.items if item.validacion == "Parcial")
        no_llego = sum(1 for item in self.items if item.validacion == "No Llegó")
        progreso = (escaneados / total) * 100 if total > 0 else 0
        
        # Actualizar labels
        self.label_escaneados.configure(text=f"Escaneados: {escaneados}")
        self.label_pendientes.configure(text=f"Pendientes: {pendientes}")
        self.label_validados.configure(text=f"Validados: {validados} (Parciales: {parciales}, No Llegó: {no_llego})")
        self.label_progreso.configure(text=f"Progreso: {progreso:.1f}%")
        self.progressbar.set(escaneados / total)

    def actualizar_escaneo(self, codigo):
        """Notificación desde el parent cuando se escanea un código.
        Actualiza el contador del item correspondiente, marca validación y refresca vista/estadísticas.
        """
        try:
            actualizado = False
            # Normalizar codigo a str
            clave = str(codigo).strip()
            clave_digits = re.sub(r"\D", "", clave)

            for item in self.items:
                # Para modo factura usar upc, para layout usar parte
                target = item.upc if self.modo == 'factura' else item.parte
                target_str = str(target).strip()
                target_digits = re.sub(r"\D", "", target_str)

                # Comparaciones robustas: exacta o por dígitos
                match_exact = (target_str == clave)
                # Consider equal if digits-only match after stripping leading zeros
                match_digits = False
                if target_digits and clave_digits:
                    if target_digits == clave_digits:
                        match_digits = True
                    else:
                        # Normalize by removing leading zeros to handle cases like '0873...' vs '873...'
                        if target_digits.lstrip('0') == clave_digits.lstrip('0'):
                            match_digits = True

                if match_exact or match_digits:
                    # Sincronizar escaneos desde dict compartido: intentar varias claves posibles
                    try:
                        # Preferir la clave ya usada en el contador
                        item.escaneos = int(self.contador_escaneos.get(clave, 
                                            self.contador_escaneos.get(clave_digits, 
                                            self.contador_escaneos.get(target_str, 
                                            self.contador_escaneos.get(target_digits, 0)))) or 0)
                    except Exception:
                        item.escaneos = (item.escaneos or 0) + 1
                    # Marcar validación como Llegó si hay al menos 1 escaneo
                    if item.escaneos > 0:
                        item.validacion = "Llegó"
                    actualizado = True
                    # No romper: podría haber duplicados, actualizarlos todos

            if actualizado:
                # Refrescar tabla filtrada (rápido) y estadísticas
                try:
                    self.actualizar_tabla_filtrada()
                except Exception:
                    pass
                try:
                    self.actualizar_estadisticas()
                except Exception:
                    pass
        except Exception:
            pass
    
    def filtrar_rapido(self, event=None):
        """Filtra la tabla de forma ultra rápida"""
        texto = self.entry_busqueda.get().strip().lower()
        campo = self.combo_campo.get()
        
        if not texto:
            # Mostrar todos
            for item in self.items:
                item.visible = True
            self.actualizar_tabla_filtrada()
            return
        
        # Filtrar según campo
        for item in self.items:
            if campo == "Todo":
                item.visible = self.buscar_en_todo(item, texto)
            elif campo == "Código":
                item.visible = self.buscar_en_codigo(item, texto)
            elif campo == "Descripción":
                item.visible = self.buscar_en_descripcion(item, texto)
            elif campo == "Factura" and self.modo == "factura":
                item.visible = texto in item.factura.lower()
        
        # Actualizar tabla
        self.actualizar_tabla_filtrada()
        self.actualizar_estadisticas()
    
    def buscar_en_todo(self, item, texto):
        """Busca texto en todos los campos relevantes"""
        if self.modo == "factura":
            return (texto in item.upc.lower() or 
                    texto in item.descripcion.lower() or 
                    texto in item.factura.lower())
        else:
            return (texto in item.parte.lower() or 
                    texto in item.descripcion.lower())
    
    def buscar_en_codigo(self, item, texto):
        """Busca texto solo en código"""
        if self.modo == "factura":
            return texto in item.upc.lower()
        else:
            return texto in item.parte.lower()
    
    def buscar_en_descripcion(self, item, texto):
        """Busca texto solo en descripción"""
        return texto in item.descripcion.lower()
    
    def actualizar_tabla_filtrada(self):
        """Actualiza la tabla mostrando solo items visibles"""
        # Limpiar tabla
        for item in self.tabla.get_children():
            self.tabla.delete(item)
        
        # Insertar solo items visibles
        items_visibles = [item for item in self.items if item.visible]
        
        # Insertar en lote
        for item in items_visibles:
            if self.modo == "factura":
                self.tabla.insert('', 'end', values=(
                    item.upc,
                    item.factura,
                    item.descripcion[:50] + "..." if len(item.descripcion) > 50 else item.descripcion,
                    item.cantidad_editada or item.cantidad_original,
                    str(item.escaneos),
                    item.validacion,
                    item.observaciones[:30] + "..." if len(item.observaciones) > 30 else item.observaciones
                ))
            else:
                self.tabla.insert('', 'end', values=(
                    item.parte,
                    item.descripcion[:50] + "..." if len(item.descripcion) > 50 else item.descripcion,
                    item.cantidad_editada or item.cantidad_original,
                    str(item.escaneos),
                    item.validacion,
                    item.observaciones[:30] + "..." if len(item.observaciones) > 30 else item.observaciones
                ))
        
        # Actualizar contador
        self.label_total.configure(text=f"Mostrando: {len(items_visibles)} de {len(self.items)}")
    
    def editar_celda_rapido(self, event):
        """Permite editar una celda con doble clic"""
        # Obtener celda seleccionada
        region = self.tabla.identify_region(event.x, event.y)
        if region != "cell":
            return
        
        # Obtener fila y columna
        item_id = self.tabla.identify_row(event.y)
        col = self.tabla.identify_column(event.x)
        
        if not item_id:
            return
        
        # Obtener valores actuales
        valores = list(self.tabla.item(item_id, 'values'))
        col_index = int(col[1:]) - 1
        valor_actual = valores[col_index]
        
        # Obtener coordenadas
        x, y, width, height = self.tabla.bbox(item_id, col)
        
        # Crear widget de edición según columna
        if self.tabla.columnas[col_index] in ["CANTIDAD", "OBSERVACIONES"]:
            self.crear_entry_edicion(x, y, width, height, valor_actual, item_id, col_index)
        elif self.tabla.columnas[col_index] == "VALIDACIÓN":
            self.crear_combo_validacion(x, y, width, height, valor_actual, item_id, col_index)
    
    def crear_entry_edicion(self, x, y, width, height, valor, item_id, col_index):
        """Crea un Entry para editar"""
        # Limpiar widget anterior
        if self.widget_edicion:
            self.widget_edicion.destroy()
        
        # Crear nuevo Entry
        self.widget_edicion = ttk.Entry(self.tabla, font=("Inter", 11))
        self.widget_edicion.insert(0, valor)
        self.widget_edicion.place(x=x, y=y, width=width, height=height)
        
        # Guardar contexto
        self.celda_editando = (item_id, col_index)
        
        # Configurar eventos
        self.widget_edicion.focus_set()
        self.widget_edicion.select_range(0, 'end')
        self.widget_edicion.bind('<Return>', self.guardar_edicion_rapida)
        self.widget_edicion.bind('<Escape>', self.cancelar_edicion)
    
    def crear_combo_validacion(self, x, y, width, height, valor, item_id, col_index):
        """Crea un ComboBox para validación"""
        if self.widget_edicion:
            self.widget_edicion.destroy()
        
        self.widget_edicion = ttk.Combobox(
            self.tabla, 
            values=["Llegó", "No Llegó", "Parcial", "No Aplica"],
            state="readonly",
            font=("Inter", 11)
        )
        self.widget_edicion.set(valor)
        self.widget_edicion.place(x=x, y=y, width=width, height=height)
        
        self.celda_editando = (item_id, col_index)
        self.widget_edicion.focus_set()
        self.widget_edicion.bind('<<ComboboxSelected>>', lambda e: self.guardar_edicion_rapida())
        self.widget_edicion.bind('<Escape>', self.cancelar_edicion)
        self.widget_edicion.bind('<Return>', lambda e: self.guardar_edicion_rapida())
    
    def guardar_edicion_rapida(self, event=None):
        """Guarda la edición actual"""
        if not self.celda_editando or not self.widget_edicion:
            return
        
        item_id, col_index = self.celda_editando
        
        try:
            nuevo_valor = self.widget_edicion.get()
            
            # Actualizar tabla
            valores = list(self.tabla.item(item_id, 'values'))
            valores[col_index] = nuevo_valor
            self.tabla.item(item_id, values=valores)
            
            # Actualizar objeto en memoria
            self.actualizar_item_en_memoria(item_id, col_index, nuevo_valor)
            
            # Limpiar edición
            self.cancelar_edicion()
            
            # Actualizar estadísticas si es validación
            if self.tabla.columnas[col_index] == "VALIDACIÓN":
                self.actualizar_estadisticas()
                
        except Exception as e:
            print(f"Error al guardar edición: {e}")
            self.cancelar_edicion()
    
    def actualizar_item_en_memoria(self, item_id, col_index, nuevo_valor):
        """Actualiza el objeto en memoria"""
        # Encontrar índice del item en la tabla visible
        all_items = self.tabla.get_children()
        if item_id not in all_items:
            return
            
        item_index = all_items.index(item_id)
        
        # Buscar el item correspondiente en items visibles
        items_visibles = [item for item in self.items if item.visible]
        if item_index >= len(items_visibles):
            return
            
        item_visible = items_visibles[item_index]
        
        # Encontrar el item original en self.items
        for item in self.items:
            if self.modo == "factura":
                if item.upc == item_visible.upc and item.factura == item_visible.factura:
                    columna = self.tabla.columnas[col_index]
                    
                    # Actualizar según columna
                    if columna == "CANTIDAD":
                        if item.cantidad_editada != nuevo_valor:
                            item.cantidad_editada = nuevo_valor
                            item.modificado = True
                            self.hay_cambios_sin_guardar = True
                    elif columna == "VALIDACIÓN":
                        if item.validacion != nuevo_valor:
                            item.validacion = nuevo_valor
                            item.modificado = True
                            self.hay_cambios_sin_guardar = True
                    elif columna == "OBSERVACIONES":
                        if item.observaciones != nuevo_valor:
                            item.observaciones = nuevo_valor
                            item.modificado = True
                            self.hay_cambios_sin_guardar = True
                    break
            else:
                if item.parte == item_visible.parte:
                    columna = self.tabla.columnas[col_index]
                    
                    # Actualizar según columna
                    if columna == "CANTIDAD":
                        if item.cantidad_editada != nuevo_valor:
                            item.cantidad_editada = nuevo_valor
                            item.modificado = True
                            self.hay_cambios_sin_guardar = True
                    elif columna == "VALIDACIÓN":
                        if item.validacion != nuevo_valor:
                            item.validacion = nuevo_valor
                            item.modificado = True
                            self.hay_cambios_sin_guardar = True
                    elif columna == "OBSERVACIONES":
                        if item.observaciones != nuevo_valor:
                            item.observaciones = nuevo_valor
                            item.modificado = True
                            self.hay_cambios_sin_guardar = True
                    break
    
    def cancelar_edicion(self, event=None):
        """Cancela la edición actual"""
        if self.widget_edicion:
            try:
                self.widget_edicion.destroy()
            except:
                pass
            self.widget_edicion = None
        self.celda_editando = None
    
    def guardar_todo_rapido(self):
        """Guarda todos los cambios de forma ultra rápida"""
        try:
            cambios = []
            items_modificados = 0
            
            for item in self.items:
                # Solo guardar items que han sido modificados
                if item.modificado:
                    items_modificados += 1
                    
                    if self.modo == "factura":
                        cambios.append({
                            'TIPO': 'factura',
                            'UPC': item.upc,
                            'Factura': item.factura,
                            'Cantidad_Factura_Original': item.cantidad_original,
                            'Cantidad_Factura_Editada': item.cantidad_editada or item.cantidad_original,
                            'Cantidad_Escaneada': str(item.escaneos),
                            'Validacion': item.validacion,
                            'Observaciones': item.observaciones,
                            'Descripcion': item.descripcion
                        })
                    else:
                        cambios.append({
                            'TIPO': 'layout',
                            'Parte': item.parte,
                            'Cantidad_Original': item.cantidad_original,
                            'Cantidad_Editada': item.cantidad_editada or item.cantidad_original,
                            'Cantidad_Escaneada': str(item.escaneos),
                            'Validacion': item.validacion,
                            'Observaciones': item.observaciones,
                            'Descripcion': item.descripcion
                        })
            
            if items_modificados == 0:
                messagebox.showinfo(
                    "Información", 
                    "No hay cambios para guardar.",
                    parent=self
                )
                return
            
            # Guardar en parent
            if self.modo == 'factura':
                self.parent.cambios_factura = cambios
            else:
                self.parent.cambios_layout = cambios
            
            # Guardar en archivo
            self.guardar_cambios_archivo_rapido(cambios)
            
            # Resetear estado de modificados
            for item in self.items:
                item.modificado = False
            self.hay_cambios_sin_guardar = False
            
            messagebox.showinfo(
                "✅ Guardado Exitoso", 
                f"Se guardaron {items_modificados} registros",
                parent=self
            )
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron guardar los cambios:\n{e}", parent=self)
    
    def guardar_cambios_archivo_rapido(self, cambios):
        """Guarda cambios en archivo de forma optimizada"""
        try:
            timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
            archivo = os.path.join("data", f"cambios_{self.modo}_{timestamp}.json")
            
            # Crear directorio si no existe
            os.makedirs("data", exist_ok=True)
            
            # Guardar con separadores compactos
            with open(archivo, 'w', encoding='utf-8') as f:
                json.dump(cambios, f, separators=(',', ':'), ensure_ascii=False)
            
            # Guardar también como actual
            archivo_actual = os.path.join("data", f"cambios_{self.modo}_actual.json")
            with open(archivo_actual, 'w', encoding='utf-8') as f:
                json.dump(cambios, f, separators=(',', ':'), ensure_ascii=False)
                
        except Exception as e:
            print(f"Error guardando archivo: {e}")
    
    def exportar_excel_rapido(self):
        """Exporta a Excel de forma ultra rápida - INCLUYE TODAS LAS COLUMNAS"""
        try:
            # Crear DataFrame con TODAS las columnas de la tabla
            if self.modo == "factura":
                data = []
                for item in self.items:
                    data.append({
                        'UPC': item.upc,
                        'FACTURA': item.factura,
                        'DESCRIPCIÓN': item.descripcion,
                        'CANTIDAD_ORIGINAL': item.cantidad_original,
                        'CANTIDAD_EDITADA': item.cantidad_editada or item.cantidad_original,
                        # 'ESCANEOS': item.escaneos,
                        'VALIDACIÓN': item.validacion,
                        # 'OBSERVACIONES': item.observaciones,  # ¡INCLUYE OBSERVACIONES!
                    })
                columnas_exportar = ['UPC', 'FACTURA', 'DESCRIPCIÓN', 'CANTIDAD_ORIGINAL', 
                                    'CANTIDAD_EDITADA', 'VALIDACIÓN']
            else:
                data = []
                for item in self.items:
                    data.append({
                        'PARTE': item.parte,
                        'DESCRIPCIÓN': item.descripcion,
                        'CANTIDAD_ORIGINAL': item.cantidad_original,
                        'CANTIDAD_EDITADA': item.cantidad_editada or item.cantidad_original,
                        # 'ESCANEOS': item.escaneos,
                        'VALIDACIÓN': item.validacion,
                        # 'OBSERVACIONES': item.observaciones,  # ¡INCLUYE OBSERVACIONES!
                    })
                columnas_exportar = ['PARTE', 'DESCRIPCIÓN', 'CANTIDAD_ORIGINAL', 
                                    'CANTIDAD_EDITADA', 'VALIDACIÓN']
            
            df = pd.DataFrame(data, columns=columnas_exportar)
            
            # Agregar totales
            df = self.agregar_totales_dataframe(df)
            
            # Pedir ubicación para guardar
            timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
            tipo = "Factura" if self.modo == "factura" else "Layout"
            archivo_default = f"Reporte_{tipo}_{timestamp}.xlsx"
            
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=archivo_default,
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            
            if file_path:
                # Guardar con formato
                self.guardar_excel_con_formato_rapido(df, file_path)
                
                messagebox.showinfo(
                    "✅ Exportación Exitosa", 
                    f"Reporte guardado:\n{file_path}\n\n"
                    f"Total de registros: {len(df)-1}\n"
                    f"Columnas exportadas: {', '.join(columnas_exportar)}",
                    parent=self
                )
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar:\n{e}", parent=self)
    
    def agregar_totales_dataframe(self, df):
        """Agrega fila de totales al DataFrame sumando columnas numéricas como ENTEROS"""
        try:
            df_calc = df.copy()

            # Convertir columnas a ENTERO
            for col in ['CANTIDAD_ORIGINAL', 'CANTIDAD_EDITADA']:
                if col in df_calc.columns:
                    df_calc[col] = (
                        pd.to_numeric(df_calc[col], errors='coerce')
                        .fillna(0)
                        .astype(int)
                    )

            # Calcular totales por columna
            total_original = int(df_calc['CANTIDAD_ORIGINAL'].sum())
            total_editada = int(df_calc['CANTIDAD_EDITADA'].sum())

            # Construir fila de totales
            if self.modo == "factura":
                total_row = {
                    'UPC': 'TOTALES',
                    'FACTURA': '',
                    'DESCRIPCIÓN': '',
                    'CANTIDAD_ORIGINAL': total_original,
                    'CANTIDAD_EDITADA': total_editada,
                    'VALIDACIÓN': ''
                }
            else:
                total_row = {
                    'PARTE': 'TOTALES',
                    'DESCRIPCIÓN': '',
                    'CANTIDAD_ORIGINAL': total_original,
                    'CANTIDAD_EDITADA': total_editada,
                    'VALIDACIÓN': ''
                }

            # Asegurar todas las columnas
            for col in df.columns:
                total_row.setdefault(col, '')

            # Agregar fila final
            return pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

        except Exception as e:
            print(f"Error agregando totales: {e}")
            return df

    def guardar_excel_con_formato_rapido(self, df, file_path):
        """Guarda Excel con formato optimizado - EXPORTA TODAS LAS COLUMNAS"""
        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # Guardar DataFrame con todas las columnas
                df.to_excel(writer, sheet_name='Reporte', index=False)
                
                # Formato básico
                workbook = writer.book
                worksheet = writer.sheets['Reporte']
                
                # Ajustar anchos de columnas
                for i, column in enumerate(df.columns):
                    column_width = max(df[column].astype(str).map(len).max(), len(column)) + 2
                    worksheet.column_dimensions[chr(65 + i)].width = min(column_width, 50)
                
                # Formato para la fila de totales (últltima fila)
                if 'TOTALES' in df.iloc[-1].values:
                    last_row = len(df) + 1  # +1 porque Excel empieza en 1
                    for col in range(1, len(df.columns) + 1):
                        cell = worksheet.cell(row=last_row, column=col)
                        cell.font = cell.font.copy(bold=True)
                        cell.fill = cell.fill.copy(patternType="solid", fgColor="FFFF00")  # Amarillo
                
            print(f"✅ Excel exportado con {len(df.columns)} columnas: {list(df.columns)}")
        except Exception as e:
            raise e
    
    def mostrar_error(self, mensaje):
        """Muestra mensaje de error"""
        messagebox.showerror("Error", mensaje, parent=self)
    
    def on_close(self):
        """Maneja el cierre de la ventana - SOLO PREGUNTA SI HAY CAMBIOS"""
        try:
            # Verificar si hay cambios sin guardar
            if self.hay_cambios_sin_guardar:
                respuesta = messagebox.askyesnocancel(
                    "Guardar cambios",
                    "Tiene cambios sin guardar. ¿Desea guardar los cambios antes de cerrar?",
                    parent=self
                )
                
                if respuesta is None:  # Cancelar
                    return  # No cerrar la ventana
                elif respuesta:  # Sí
                    self.guardar_todo_rapido()
                    # Después de guardar, verificar si se guardó correctamente
                    if not self.hay_cambios_sin_guardar:
                        # Proceder a cerrar
                        pass
                    else:
                        # Si todavía hay cambios (error al guardar), no cerrar
                        return
            
            # Limpiar y cerrar
            self.cancelar_carga.set()
            self.executor.shutdown(wait=False)
            
            self.grab_release()
            self.destroy()
            
        except Exception as e:
            print(f"Error al cerrar: {e}")
            try:
                self.destroy()
            except:
                pass