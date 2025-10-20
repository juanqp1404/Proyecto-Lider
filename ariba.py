import os
import re
import time
from playwright.sync_api import Playwright, sync_playwright, expect
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

# TODO: cerrar proceso msedge si esta en segundo plano con task manager antes de ejecutar

def run(playwright: Playwright) -> None:
    user_data_dir = r"C:/Users/JQuintero27/AppData/Local/Microsoft/Edge/User Data"
    carpeta_actual = "./"
    # browser = playwright.chromium.launch(headless=False,executable_path='C:/Program Files/Google/Chrome/Application/chrome.exe')
    context = playwright.chromium.launch_persistent_context(user_data_dir, headless=False, channel="msedge")
    page = context.new_page()
    page.goto("https://s1.ariba.com/Sourcing/Main/aw?awh=r&awssk=ODTzCxAIbpbHwlV1&realm=schlumberger")
    # time.sleep(40)
    page.get_by_role("button", name="Manage").click()
    time.sleep(2)
    page.get_by_role("menuitem", name="Queues").click()
    time.sleep(2)
    page.locator("#text__c_npbb").click()
    time.sleep(2)
    page.get_by_role("option", name="Requisition").click()
    time.sleep(2)
    page.get_by_role("textbox", name="From:").fill("Sun, 7 Sep, 2025")
    time.sleep(2)
    page.get_by_title("Run this search").click()
    time.sleep(2)
    # page.get_by_role("button", id="_7msd8").click()
    page.query_selector('#_7msd8')
    time.sleep(2)

    with page.expect_download() as download_info:
        page.get_by_role("menuitem", name="Export all Rows").click()

    download = download_info.value

    new_filename = "DF"
    print(f"Nombre sugerido del archivo: {download.suggested_filename}")
    print(f"Descargando el archivo como: {new_filename}")
    print(f"URL del archivo: {download}")
    time.sleep(20)
    # Guardar el archivo en la carpeta actual con su nombre original
    ruta_destino = os.path.join(carpeta_actual, new_filename)
    download.save_as(ruta_destino)

    # page.goto("https://example.com/")
    # page.get_by_role("heading", name="Example Domain").click()
    # page.get_by_role("link", name="More information...").click()
    # page.get_by_role("link", name="IANA-managed Reserved Domains").click()
    # page.get_by_role("link", name="XN--HLCJ6AYA9ESC7A").click()

    # ---------------------
    # time.sleep(15)

    # page.fill('#i0116','JQuintero27@slb.com')
    # time.sleep(305)

    context.close()
    # browser.close()

kill_edge_processes()

with sync_playwright() as playwright:
    run(playwright)
