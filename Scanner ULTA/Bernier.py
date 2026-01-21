# Bernier.py
import customtkinter as ctk
from Configuracion import extraer_valor_numerico
import math

# Usar los mismos estilos que Scanner.py para consistencia
STYLE = {
    "fondo": "#F8F9FA",
    "borde": "#BDC3C7", 
    "secundario": "#282828",
    "exito": "#27AE60",
    "texto_oscuro": "#282828",
    "texto_claro": "#4b4b4b"
}

FONT_TEXT = ("Inter", 12)

def dibujar_regla_bernier(canvas, declaracion_contenido):
    """Dibuja la regla vertical (0–10 mm = 1 cm) usando solo el tamaño declarado."""
    canvas.delete("all")

    # Medidas del canvas
    canvas.update_idletasks()
    width = canvas.winfo_width() or 100
    raw_height = canvas.winfo_height() or 0
    MIN_HEIGHT = 180
    height = raw_height if raw_height >= MIN_HEIGHT else MIN_HEIGHT

    # Texto original (preferimos mostrar exactamente lo que viene en el dato)
    original_str = str(declaracion_contenido).strip()
    valor_declarado = extraer_valor_numerico(declaracion_contenido)
    try:
        valor_num = float(valor_declarado) if valor_declarado is not None else None
    except Exception:
        valor_num = None

    # Definir rango máximo dinámico: por defecto 10 mm (1 cm).
    rango_max = 10
    if valor_num and valor_num > rango_max:
        rango_max = int(math.ceil(valor_num / 5.0) * 5)
    DESIRED_RULE_HEIGHT = 280  # px, altura visual objetivo de la regla
    available = max(0, height - 40)  # dejar algo de espacio superior/inferior
    alto_util = DESIRED_RULE_HEIGHT if available >= DESIRED_RULE_HEIGHT else available

    # Centrar verticalmente la regla dentro del canvas
    top_margin = max(10, (height - alto_util) // 2)
    bottom_margin = height - top_margin - alto_util
    margen_sup = top_margin
    margen_inf = bottom_margin
    x_regla = width // 2

    # Fondo de la regla
    canvas.create_rectangle(
        x_regla - 15, margen_sup - 5,
        x_regla + 15, height - margen_inf + 5,
        fill=STYLE["fondo"], outline=STYLE["borde"], width=1
    )

    # Dibujar líneas y marcas. Etiquetamos cada 5 unidades para no saturar.
    step = 1
    for i in range(0, rango_max + 1, step):
        # mapea i en [0..rango_max] a coordenadas dentro de la zona de la regla
        y = height - margen_inf - (i / rango_max) * alto_util
        if i % 5 == 0:
            canvas.create_line(x_regla - 15, y, x_regla + 15, y, width=2, fill=STYLE["secundario"])
            canvas.create_text(x_regla - 30, y, text=f"{i}", anchor="e",
                            fill=STYLE["secundario"], font=FONT_TEXT)
        else:
            canvas.create_line(x_regla - 10, y, x_regla + 10, y, width=1, fill=STYLE["secundario"])
    if valor_num and valor_num > 0:
        # Evitar overflow: si valor_num > rango_max, la barra llenará hasta el tope
        altura_mm = (min(valor_num, rango_max) / rango_max) * alto_util
        y_tope = height - margen_inf - altura_mm

        canvas.create_rectangle(
            x_regla - 8, y_tope,
            x_regla + 8, height - margen_inf,
            fill=STYLE["exito"], outline="#2E7D32", width=2
        )

        # Marcar el valor exacto en la regla: preferimos mostrar la cadena original
        if original_str and any(ch.isdigit() for ch in original_str):
            texto_val = original_str
        else:
            try:
                texto_val = f"{valor_num:.1f} mm"
            except Exception:
                texto_val = str(valor_declarado)

        canvas.create_text(x_regla + 40, y_tope - 10,
                        text=texto_val,
                        fill=STYLE["exito"], font=FONT_TEXT, anchor="w")
    else:
        # No hay número: mostrar la cadena original si existe, sino 'Sin medida'
        texto_val = original_str if original_str else " "
        canvas.create_text(x_regla + 40, margen_sup + 6,
                        text=texto_val,
                        fill=STYLE["texto_claro"], font=FONT_TEXT, anchor="w")

    # Etiquetas superior e inferior
    canvas.create_text(x_regla, margen_sup - 10,
                    text="10 mm (1 cm)", fill=STYLE["secundario"], font=FONT_TEXT)
    canvas.create_text(x_regla, height - margen_inf + 15,
                    text="0 mm", fill=STYLE["secundario"], font=FONT_TEXT)
    
