"""Módulo para gestionar los puntos normativos según las diferentes normas NOM."""
def obtener_puntos_normativos(norma: str):
    """Devuelve una lista de puntos normativos según la norma con mejor formato."""
    mapa_puntos = {
        "NOM-141": [
            "• Denominación genérica o específica (opcional)",
            "• Leyenda 'Contenido' o 'Contenido neto' (no obligatoria)",
            "• Nombre o razón social del responsable del producto",
            "• País de origen",
            "• Número de lote",
            "• Leyendas precautorias",
            "• Instrucciones de uso (obligatorias en los siguientes casos):"
            "  - Tintes, colorantes o coloración",
            "  - Decolorantes",
            "  - Permanentes",
            "  - Alisadores permanentes",
            "  - Productos para la piel cuya función primaria sea la protección solar",
            "  - Bronceadores o autobronceadores",
            "  - Depilatorios o epilatorios",
            "  - O cualquier otro producto que lo requiera",
            "• Listado de ingredientes"
        ],
        
        "NOM-004": [
            "📋 ROPA DE CASA:",
            "• Insumos de mayor a menor porcentaje",
            "• Importador",
            "• País de origen",
            "• Instrucciones de cuidado",
            "• Marca",
            "• Medidas",
            
            "👕 PRENDA DE VESTIR:",
            "• Insumos de mayor a menor porcentaje",
            "• Importador",
            "• País de origen",
            "• Instrucciones de cuidado",
            "• Marca",
            "• Talla",
            
            "🧵 TEXTILES:",
            "• Insumos de mayor a menor porcentaje",
            "• Importador",
            "• País de origen",
            "• Marca",
            "• Medidas",
        ],
        
        "NOM-050": [
            "• Marca",
            "• Denominación (si no se identifica a simple vista)",
            "• Contenido (si no se identifica a simple vista)",
            "• Importador",
            "• País de origen"
        ],

        "NOM-020": [
            "• Importador",
            "• País de origen",
            "• Insumos cuando aplique forro"
        ],

        "NOM-015": [
            "• Etiquetado de alimentos y bebidas",
            "• Información nutrimental",
            "• Lista de ingredientes y aditivos",
            "• Contenido neto"
        ],
        
        "NOM-024": [
            "🔌 ELECTRÓNICOS, ELÉCTRICOS Y ELECTRODOMÉSTICOS – MANUAL:",
            "• Marca",
            "• Denominación",
            "• País de origen",
            "• Importador",
            "• Contenido cuando aplique",
            "• Características eléctricas",
            
            "🔧 REPUESTOS CONSUMIBLES Y DESECHABLES:",
            "• Marca",
            "• Denominación",
            "• País de origen",
            "• Importador",
            
            "📝 NOTA: Se utilizan dos tipos de etiquetas:",
            "• Etiqueta blanca (comercial)",
            "• Etiqueta metálica (del producto)"
        ],
    }
    
    return mapa_puntos.get(norma.upper(), ["❌ No se encontraron puntos normativos definidos para esta norma."])
def obtener_normas_disponibles():
    """Devuelve una lista de todas las normas disponibles en el sistema."""
    return [
        "NOM-141", "NOM-004", "NOM-015", "NOM-050", 
        "NOM-020", "NOM-024"
    ]

def validar_norma(norma: str):
    """Valida si una norma existe en el sistema."""
    normas = obtener_normas_disponibles()
    return norma.upper() in normas

