from datetime import datetime, timedelta

def calcular_domingo_asociado(fecha=None):
    """
    Calcula el domingo asociado según un patrón de rangos de fechas
    y devuelve un string formato 'Sun, 7 Sep, 2025'.
    """
    if fecha is None:
        fecha = datetime.today().date()
    
    FECHA_INICIO = datetime(2025, 10, 14).date()
    DOMINGO_BASE = datetime(2025, 9, 7).date()
    
    dias_transcurridos = (fecha - FECHA_INICIO).days
    
    if dias_transcurridos < 0:
        semanas_hacia_atras = (abs(dias_transcurridos) + 6) // 7
        domingo_resultado = DOMINGO_BASE - timedelta(weeks=semanas_hacia_atras)
    elif dias_transcurridos <= 5:
        domingo_resultado = DOMINGO_BASE
    else:
        dias_post_primera = dias_transcurridos - 6
        semanas_extra = (dias_post_primera // 7) + 1
        domingo_resultado = DOMINGO_BASE + timedelta(weeks=semanas_extra)
    
    # Formatear sin cero inicial en el día
    abreviatura_dia = domingo_resultado.strftime('%a')
    abreviatura_mes = domingo_resultado.strftime('%b')
    dia = domingo_resultado.day
    anio = domingo_resultado.year
    
    return f"{abreviatura_dia}, {dia} {abreviatura_mes}, {anio}"


# Uso del script
domingo = calcular_domingo_asociado()
print(domingo)  # Ejemplo de salida: "Sun, 14 Sep, 2025"
