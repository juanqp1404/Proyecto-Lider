import re
from playwright.sync_api import Playwright, sync_playwright, expect
import time
import psutil

def kill_edge_processes():
    """
    Verifica si hay procesos de Edge ejecutándose y los cierra antes de continuar.
    """
    edge_process_names = ["msedge.exe", "msedgedriver.exe"]
    processes_killed = False
    
    print("Verificando procesos de Edge...")
    
    for proc in psutil.process_iter(['name', 'pid']):
        try:
            # Verifica si el proceso es Edge o EdgeDriver
            if proc.info['name'] in edge_process_names:
                print(f"Cerrando proceso: {proc.info['name']} (PID: {proc.info['pid']})")
                proc.kill()  # Usa terminate() en lugar de kill() si prefieres un cierre más suave
                processes_killed = True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            # Ignora procesos que ya no existen o a los que no tienes acceso
            pass
    
    if processes_killed:
        print("Procesos de Edge cerrados. Esperando 2 segundos...")
        time.sleep(2)  # Espera para asegurar que los procesos se cierren completamente
    else:
        print("No se encontraron procesos de Edge en ejecución.")


def run(playwright: Playwright) -> None:
    user_data_dir = r"C:/Users/xanti/AppData/Local/Microsoft/Edge/User Data"
    
    # Usa launch_persistent_context en lugar de launch
    print("Abriendo contexto persistente...")
    context = playwright.chromium.launch_persistent_context(
        user_data_dir,
        channel="msedge",
        headless=False,
        args=[
            "--start-fullscreen"
        ],
    )

    page = context.new_page()
   
    page.goto("https://www.youtube.com/", wait_until="networkidle")
    page.wait_for_timeout(8000)
    page.get_by_role("link", name="Tú").click()
    page.wait_for_timeout(8000)
    target_element = page.locator('ytd-rich-shelf-renderer').filter(has_text='Ver más tarde 274 vídeos Ver').get_by_label("CA7RIEL & Paco Amoroso - BABY GANGSTA (Visualizer)")

    target_element.scroll_into_view_if_needed()

    target_element.click()

    page.wait_for_timeout(50000)
    # ---------------------
    context.close()


# Primero cierra todos los procesos de Edge
kill_edge_processes()

# Luego ejecuta Playwright
with sync_playwright() as playwright:
    run(playwright)
