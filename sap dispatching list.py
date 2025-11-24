from datetime import datetime, timedelta
import os
import re
import time
from playwright.sync_api import Playwright, sync_playwright, expect
import pandas as pd
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
    user_data_dir = r"C:/Users/JQuintero27/AppData/Local/Microsoft/Edge/User Data"
    carpeta_actual = "./"
    # browser = playwright.chromium.launch(headless=False,executable_path='C:/Program Files/Google/Chrome/Application/chrome.exe')
    context = playwright.chromium.launch_persistent_context(user_data_dir, headless=False, channel="msedge")
    page = context.new_page()
    page.goto("https://slb001.sharepoint.com/sites/BogotaPSC/Lists/SAP%20Dispatching%20List%20January%202019/AllItems.aspx?FilterField1=Author&FilterValue1=Juan%20Quintero&FilterType1=User&sortField=Created&isAscending=false&viewid=ba980fc8%2D7c86%2D4e24%2D8f25%2Dd39c7c4d7aeb")
    page.wait_for_selector('role=menuitem[name="Export"]', state="visible")
    page.get_by_role("menuitem", name="Export").click()
    page.wait_for_selector('role=menuitem[name="Export to CSV"]', state="visible")
    
    with page.expect_download(timeout=450) as download_info:
       page.get_by_role("menuitem", name="Export to CSV", exact=True).click()

    download = download_info.value

    new_filename = "sap_dispatching_list.csv"
    print(f"Nombre sugerido del archivo: {download.suggested_filename}")
    print(f"Descargando el archivo como: {new_filename}")
    print(f"URL del archivo: {download}")
    # Guardar el archivo en la carpeta actual con su nombre original
    ruta_destino = os.path.join(carpeta_actual,"data/sharepoint/", new_filename)
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
