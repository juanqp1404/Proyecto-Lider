"""
Funciones de acciones específicas en Power BI
"""
from pywinauto import keyboard, mouse
import time


def refrescar_datos():
    """Refresca los datos del modelo"""
    print("→ Refrescando datos...")
    keyboard.send_keys('^r')
    time.sleep(5)
    print("✓ Datos refrescados")


def guardar_archivo():
    """Guarda el archivo .pbix"""
    print("→ Guardando archivo...")
    keyboard.send_keys('^s')
    time.sleep(2)
    print("✓ Archivo guardado")


def exportar_visual_coordenadas(x: int, y: int, nombre_archivo: str = None):
    """
    Exporta datos de un visual usando coordenadas de pantalla
    
    Args:
        x, y: Coordenadas del visual
        nombre_archivo: Nombre del archivo a guardar (opcional)
    """
    print(f"→ Haciendo clic en visual ({x}, {y})...")
    mouse.click(coords=(x, y))
    time.sleep(1)
    
    print("→ Abriendo menú contextual...")
    keyboard.send_keys('+{F10}')  # Shift+F10
    time.sleep(1)
    
    print("→ Navegando al menú 'Exportar datos'...")
    keyboard.send_keys('{DOWN}{DOWN}{ENTER}')
    time.sleep(2)
    
    if nombre_archivo:
        print(f"→ Guardando como '{nombre_archivo}'...")
        keyboard.send_keys(nombre_archivo)
        keyboard.send_keys('{ENTER}')
        time.sleep(2)
    
    print("✓ Exportación completada")

def refrescar_ps_dispatching():
    mouse.click(coords=(1897, 519))
    time.sleep(2)
    mouse.click(coords=(1693, 156))
    time.sleep(2)
    mouse.click(coords=(1573, 148))

def refrescar_sap_buyers():
    mouse.click(coords=(1897, 595))
    time.sleep(2)
    mouse.click(coords=(1693, 169))
    time.sleep(2)
    mouse.click(coords=(1573, 174))

def exportar_assignation_history():
    mouse.click(coords=(1238,801))
    time.sleep(2)
    mouse.click(coords=(1238, 769))
    time.sleep(2)
    mouse.click(coords=(1349, 442))
    time.sleep(4)
    mouse.click(coords=(1072, 557))

def esperar_carga(segundos: int = 10):
    """Espera a que el dashboard cargue"""
    print(f"→ Esperando {segundos} segundos a que cargue...")
    time.sleep(segundos)
    print("✓ Carga completada")


def cerrar_powerbi():
    """Cierra Power BI Desktop"""
    print("→ Cerrando Power BI...")
    keyboard.send_keys('%{F4}')  # Alt+F4
    time.sleep(1)
    print("✓ Power BI cerrado")
