from datetime import datetime, timedelta

# Fecha de hoy
hoy = datetime.today().date()

# Cálculo de cuántos días restar para llegar al último domingo
# weekday(): lunes=0 … domingo=6 → (weekday + 1) % 7 da 0 si hoy es domingo
dias_desde_domingo = (hoy.weekday() + 1) % 7

# Fecha del último domingo
ultimo_domingo = hoy - timedelta(days=dias_desde_domingo)

# Formateo como 'YYYY-MM-DD'
fecha_str = ultimo_domingo.strftime('%Y-%m-%d')

# Ejemplo de salida
print(fecha_str)
