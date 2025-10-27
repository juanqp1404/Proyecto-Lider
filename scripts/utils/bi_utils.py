import subprocess
import time
import psutil
from pywinauto import Desktop
from typing import Optional


def abrir_powerbi(archivo_pbix: str, tiempo_espera: int = 25):
    """
    Abre un archivo .pbix con Power BI Desktop
    
    Args:
        archivo_pbix: Ruta completa al archivo .pbix
        tiempo_espera: Segundos a esperar después de abrir (default: 25)
    """
    subprocess.Popen([archivo_pbix], shell=True)
    time.sleep(tiempo_espera)

def abrir_powerbi(archivo_pbix: str, tiempo_espera: int = 25):
    """
    Abre un archivo .pbix con Power BI Desktop
    
    Args:
        archivo_pbix: Ruta completa al archivo .pbix
        tiempo_espera: Segundos a esperar después de abrir (default: 25)
    """
    subprocess.Popen([archivo_pbix], shell=True)
    time.sleep(tiempo_espera)

def obtener_pid_powerbi() -> Optional[int]:
    """
    Obtiene el PID de Power BI Desktop
    
    Returns:
        PID del proceso o None si no se encuentra
    """
    # Método 1: Buscar por nombre de proceso
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if 'PBIDesktop' in proc.info['name']:
                return proc.info['pid']
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    
    # Método 2: Buscar por ventana
    desktop = Desktop(backend="uia")
    for ventana in desktop.windows():
        try:
            if "Power BI" in ventana.window_text():
                return ventana.process_id()
        except Exception:
            pass
    
    return None

def abrir_o_conectar_powerbi(archivo_pbix: str, tiempo_espera: int = 25) -> int:
    """
    Detecta si Power BI ya está abierto. Si está abierto usa ese proceso,
    si no lo abre.
    
    Args:
        archivo_pbix: Ruta completa al archivo .pbix
        tiempo_espera: Segundos a esperar después de abrir (si es necesario)
        
    Returns:
        PID del proceso de Power BI
    """
    # Verificar si Power BI ya está abierto
    print("→ Verificando si Power BI ya está abierto...")
    pid = obtener_pid_powerbi()
    
    if pid:
        print(f"✓ Power BI ya está abierto (PID: {pid})")
        return pid
    
    # Si no está abierto, abrirlo
    print("→ Power BI no está abierto, iniciando...")
    abrir_powerbi(archivo_pbix, tiempo_espera)
    
    # Obtener el PID del proceso recién abierto
    pid = obtener_pid_powerbi()
    
    if pid:
        print(f"✓ Power BI abierto correctamente (PID: {pid})")
        return pid
    else:
        print("✗ Error: No se pudo obtener el PID después de abrir Power BI")
        return None
    
def enfocar_ventana(pid: int):
    """
    Pone la ventana de Power BI en primer plano
    
    Args:
        pid: Process ID de Power BI
    """
    from pywinauto.application import Application
    
    app = Application(backend="uia").connect(process=pid)
    ventana = app.window(title_re=".*Power BI.*")
    
    # Múltiples métodos para asegurar que se enfoque
    ventana.set_focus()
    ventana.restore()  # Por si está minimizada
    ventana.bring_to_top()
    
    return ventana

