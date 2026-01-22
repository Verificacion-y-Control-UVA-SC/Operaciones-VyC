"""
Módulo de gestión de datos para reemplazar archivos Excel
Utiliza JSON para datos estructurados y Pickle para objetos complejos
"""

import json
import pickle
import os
from pathlib import Path
from typing import Dict, List, Any, Optional
import pandas as pd
from datetime import datetime

class DataManager:
    """Gestor de datos que reemplaza la funcionalidad de archivos Excel"""
    
    def __init__(self, data_dir: str = "data"):
        self.data_dir = Path(data_dir)
        self.data_dir.mkdir(exist_ok=True)
        
        # Definir rutas de archivos
        self.base_general_path = self.data_dir / "base_general.json"
        self.inspeccion_path = self.data_dir / "inspeccion.json"
        self.historial_path = self.data_dir / "historial.pkl"
        
        # Inicializar datos en memoria
        self.base_general = {}
        self.inspeccion = {}
        self.historial = []
        self._base_items_set = set()
        self._base_general_index = {}  # Índice para búsquedas rápidas
        self._inspeccion_index = {}    # Índice para búsquedas rápidas
        
        # Cargar datos existentes
        self._load_all_data()
    
    def _load_all_data(self):
        """Cargar todos los datos desde archivos"""
        try:
            if self.base_general_path.exists():
                with open(self.base_general_path, 'r', encoding='utf-8') as f:
                    self.base_general = json.load(f)
                # Crear índices para búsquedas rápidas
                if self.base_general and 'data' in self.base_general:
                    self._create_base_general_index()
            
            if self.inspeccion_path.exists():
                with open(self.inspeccion_path, 'r', encoding='utf-8') as f:
                    self.inspeccion = json.load(f)
                # Crear índices para búsquedas rápidas
                if self.inspeccion and 'data' in self.inspeccion:
                    self._create_inspeccion_index()
            
            if self.historial_path.exists():
                with open(self.historial_path, 'rb') as f:
                    self.historial = pickle.load(f)
        except Exception as e:
            print(f"Error cargando datos: {e}")
            # Inicializar estructuras vacías en caso de error
            self.base_general = None
            self.inspeccion = None
            self.historial = []
            self._base_items_set = set()
            self._base_general_index = {}
            self._inspeccion_index = {}
    
    def _create_base_general_index(self):
        """Crear índice para búsquedas rápidas en base general"""
        self._base_general_index = {}
        self._base_items_set = set()
        
        if self.base_general and 'data' in self.base_general:
            for i, record in enumerate(self.base_general['data']):
                ean = str(record.get('EAN', ''))
                if ean:
                    self._base_general_index[ean] = i
                    self._base_items_set.add(ean)
    
    def _create_inspeccion_index(self):
        """Crear índice para búsquedas rápidas en inspección"""
        self._inspeccion_index = {}
        
        if self.inspeccion and 'data' in self.inspeccion:
            for i, record in enumerate(self.inspeccion['data']):
                item = str(record.get('ITEM', ''))
                if item:
                    self._inspeccion_index[item] = i
    
    def _save_base_general(self):
        """Guardar datos de base general en JSON"""
        try:
            with open(self.base_general_path, 'w', encoding='utf-8') as f:
                json.dump(self.base_general, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Error guardando base general: {e}")
    
    def _save_inspeccion(self):
        """Guardar datos de inspección en JSON"""
        try:
            with open(self.inspeccion_path, 'w', encoding='utf-8') as f:
                json.dump(self.inspeccion, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Error guardando inspección: {e}")
    
    def _save_historial(self):
        """Guardar historial en Pickle"""
        try:
            with open(self.historial_path, 'wb') as f:
                pickle.dump(self.historial, f)
        except Exception as e:
            print(f"Error guardando historial: {e}")
    
    def migrate_from_excel(self, base_excel_path: str, inspeccion_excel_path: str, historial_excel_path: str = None):
        """Migrar datos desde archivos Excel existentes"""
        try:
            # Migrar base general
            if os.path.exists(base_excel_path):
                df_base = pd.read_excel(base_excel_path)
                self.base_general = {
                    'columns': df_base.columns.tolist(),
                    'data': df_base.to_dict('records'),
                    'metadata': {
                        'migrated_at': datetime.now().isoformat(),
                        'source_file': base_excel_path,
                        'total_records': len(df_base)
                    }
                }
                self._save_base_general()
                print(f"Base general migrada: {len(df_base)} registros")
            
            # Migrar inspección
            if os.path.exists(inspeccion_excel_path):
                df_inspeccion = pd.read_excel(inspeccion_excel_path)
                self.inspeccion = {
                    'columns': df_inspeccion.columns.tolist(),
                    'data': df_inspeccion.to_dict('records'),
                    'metadata': {
                        'migrated_at': datetime.now().isoformat(),
                        'source_file': inspeccion_excel_path,
                        'total_records': len(df_inspeccion)
                    }
                }
                self._save_inspeccion()
                print(f"Inspección migrada: {len(df_inspeccion)} registros")
            
            # Migrar historial si existe
            if historial_excel_path and os.path.exists(historial_excel_path):
                df_historial = pd.read_excel(historial_excel_path)
                self.historial = df_historial.to_dict('records')
                self._save_historial()
                print(f"Historial migrado: {len(df_historial)} registros")
            
            return True
        except Exception as e:
            print(f"Error en migración: {e}")
            return False
    
    def get_base_general_df(self) -> pd.DataFrame:
        """Obtener base general como DataFrame"""
        if self.base_general and 'data' in self.base_general:
            return pd.DataFrame(self.base_general['data'])
        return pd.DataFrame()
    
    def get_inspeccion_df(self) -> pd.DataFrame:
        """Obtener inspección como DataFrame"""
        if self.inspeccion and 'data' in self.inspeccion:
            return pd.DataFrame(self.inspeccion['data'])
        return pd.DataFrame()
    
    def get_historial_df(self) -> pd.DataFrame:
        """Obtener historial como DataFrame"""
        if self.historial:
            return pd.DataFrame(self.historial)
        return pd.DataFrame()
    
    def add_to_historial(self, new_records: List[Dict[str, Any]]):
        """Agregar nuevos registros al historial"""
        # Convertir a DataFrame para usar drop_duplicates
        current_df = self.get_historial_df()
        new_df = pd.DataFrame(new_records)
        
        if not current_df.empty:
            # Combinar y eliminar duplicados por ITEM
            combined_df = pd.concat([current_df, new_df], ignore_index=True)
            final_df = combined_df.drop_duplicates(subset=['ITEM'], keep='last')
        else:
            final_df = new_df
        
        self.historial = final_df.to_dict('records')
        self._save_historial()
    
    def export_to_excel(self, data: pd.DataFrame, file_path: str):
        """Exportar DataFrame a Excel (para compatibilidad)"""
        try:
            data.to_excel(file_path, index=False)
            return True
        except Exception as e:
            print(f"Error exportando a Excel: {e}")
            return False
    
    def export_base_general_to_excel(self, file_path: str):
        """Exportar base general a Excel"""
        try:
            df = self.get_base_general_df()
            if not df.empty:
                df.to_excel(file_path, index=False)
                print(f"Base general exportada: {len(df)} registros")
                return True
            else:
                print("No hay datos en la base general para exportar")
                return False
        except Exception as e:
            print(f"Error exportando base general: {e}")
            return False
    
    def export_inspeccion_to_excel(self, file_path: str):
        """Exportar inspección a Excel"""
        try:
            df = self.get_inspeccion_df()
            if not df.empty:
                df.to_excel(file_path, index=False)
                print(f"Inspección exportada: {len(df)} registros")
                return True
            else:
                print("No hay datos en inspección para exportar")
                return False
        except Exception as e:
            print(f"Error exportando inspección: {e}")
            return False
    
    def export_historial_to_excel(self, file_path: str):
        """Exportar historial a Excel"""
        try:
            df = self.get_historial_df()
            if not df.empty:
                df.to_excel(file_path, index=False)
                print(f"Historial exportado: {len(df)} registros")
                return True
            else:
                print("No hay datos en historial para exportar")
                return False
        except Exception as e:
            print(f"Error exportando historial: {e}")
            return False
    
    def import_base_general_from_excel(self, file_path: str):
        """Importar base general desde Excel"""
        try:
            df = pd.read_excel(file_path)
            self.base_general = {
                'columns': df.columns.tolist(),
                'data': df.to_dict('records'),
                'metadata': {
                    'imported_at': datetime.now().isoformat(),
                    'source_file': file_path,
                    'total_records': len(df)
                }
            }
            self._save_base_general()
            # Actualizar el set de ítems para búsqueda rápida
            if self.base_general and 'data' in self.base_general:
                self._base_items_set = {str(record.get('EAN', '')) for record in self.base_general['data']}
            print(f"Base general importada: {len(df)} registros")
            return True
        except Exception as e:
            print(f"Error importando base general: {e}")
            return False
    
    def import_inspeccion_from_excel(self, file_path: str):
        """Importar inspección desde Excel"""
        try:
            df = pd.read_excel(file_path)
            self.inspeccion = {
                'columns': df.columns.tolist(),
                'data': df.to_dict('records'),
                'metadata': {
                    'imported_at': datetime.now().isoformat(),
                    'source_file': file_path,
                    'total_records': len(df)
                }
            }
            self._save_inspeccion()
            print(f"Inspección importada: {len(df)} registros")
            return True
        except Exception as e:
            print(f"Error importando inspección: {e}")
            return False
    
    def import_historial_from_excel(self, file_path: str):
        """Importar historial desde Excel"""
        try:
            df = pd.read_excel(file_path)
            self.historial = df.to_dict('records')
            self._save_historial()
            print(f"Historial importado: {len(df)} registros")
            return True
        except Exception as e:
            print(f"Error importando historial: {e}")
            return False
    
    def get_data_info(self) -> Dict[str, Any]:
        """Obtener información sobre los datos almacenados"""
        return {
            'base_general': {
                'exists': self.base_general_path.exists(),
                'records': len(self.base_general.get('data', [])) if self.base_general else 0,
                'last_modified': self.base_general_path.stat().st_mtime if self.base_general_path.exists() else None
            },
            'inspeccion': {
                'exists': self.inspeccion_path.exists(),
                'records': len(self.inspeccion.get('data', [])) if self.inspeccion else 0,
                'last_modified': self.inspeccion_path.stat().st_mtime if self.inspeccion_path.exists() else None
            },
            'historial': {
                'exists': self.historial_path.exists(),
                'records': len(self.historial),
                'last_modified': self.historial_path.stat().st_mtime if self.historial_path.exists() else None
            }
        }
    
    def item_exists_in_base(self, item: str) -> bool:
        """Verificar si un ítem existe en la base general (O(1))"""
        return str(item) in self._base_items_set
    
    def get_base_general_record_by_ean(self, ean: str) -> Optional[Dict[str, Any]]:
        """Obtener registro de base general por EAN (O(1))"""
        ean_str = str(ean)
        if ean_str in self._base_general_index:
            idx = self._base_general_index[ean_str]
            return self.base_general['data'][idx]
        return None
    
    def get_inspeccion_record_by_item(self, item: str) -> Optional[Dict[str, Any]]:
        """Obtener registro de inspección por ITEM (O(1))"""
        item_str = str(item)
        if item_str in self._inspeccion_index:
            idx = self._inspeccion_index[item_str]
            return self.inspeccion['data'][idx]
        return None
    
    def add_new_item_to_base(self, item: str, tipo_proceso: str, norma: str, descripcion: str):
        """Agregar un nuevo ítem a la base general"""
        if not self.base_general:
            self.base_general = {'data': [], 'columns': ['EAN', 'DESCRIPTION', 'CODIGO FORMATO']}
        
        new_record = {
            'EAN': str(item),
            'DESCRIPTION': descripcion,
            'CODIGO FORMATO': tipo_proceso
        }
        
        item_str = str(item)
        
        # Usar índice para búsqueda rápida
        if item_str in self._base_general_index:
            # Actualizar registro existente
            idx = self._base_general_index[item_str]
            self.base_general['data'][idx] = new_record
        else:
            # Agregar nuevo registro
            self.base_general['data'].append(new_record)
            idx = len(self.base_general['data']) - 1
            self._base_general_index[item_str] = idx
            self._base_items_set.add(item_str)
        
        self._save_base_general()
    
    def add_new_item_to_inspeccion(self, item: str, criterio: str):
        """Agregar un nuevo ítem a la inspección"""
        if not self.inspeccion:
            self.inspeccion = {'data': [], 'columns': ['ITEM', 'INFORMACION FALTANTE']}
        
        new_record = {
            'ITEM': str(item),
            'INFORMACION FALTANTE': criterio
        }
        
        item_str = str(item)
        
        # Usar índice para búsqueda rápida
        if item_str in self._inspeccion_index:
            # Actualizar registro existente
            idx = self._inspeccion_index[item_str]
            self.inspeccion['data'][idx] = new_record
        else:
            # Agregar nuevo registro
            self.inspeccion['data'].append(new_record)
            idx = len(self.inspeccion['data']) - 1
            self._inspeccion_index[item_str] = idx
        
        self._save_inspeccion()
    
    def get_new_items_from_report(self, report_items: List[int]) -> List[int]:
        """Obtener lista de ítems nuevos que no están en la base de datos (O(n))"""
        # Usar set comprehension para mayor velocidad
        report_items_set = {str(item) for item in report_items}
        new_items_set = report_items_set - self._base_items_set
        
        return [int(item) for item in new_items_set]
