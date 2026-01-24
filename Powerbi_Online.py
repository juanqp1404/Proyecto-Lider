from datetime import datetime, timedelta
import os
import re
import time
from playwright.sync_api import Playwright, sync_playwright, expect
import pandas as pd
import psutil
from datetime import datetime, timedelta

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



def calcular_domingo_asociado(fecha=None):
    """
    Calcula el domingo asociado según un patrón de rangos de fechas
    y devuelve un string formato 'M/D/YYYY', por ejemplo '1/24/2026'.
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

    # Formato M/D/YYYY sin ceros a la izquierda
    mes = domingo_resultado.month
    dia = domingo_resultado.day
    anio = domingo_resultado.year

    return f"{mes}/{dia}/{anio}"

def run(playwright: Playwright) -> None:
    user_data_dir = os.path.join(
    os.getenv('LOCALAPPDATA'),  # C:/Users/{usuario}/AppData/Local
    'Microsoft',
    'Edge',
    'User Data'
    )
    carpeta_actual = "./"
    # browser = playwright.chromium.launch(headless=False,executable_path='C:/Program Files/Google/Chrome/Application/chrome.exe')
    context = playwright.chromium.launch_persistent_context(user_data_dir, headless=False, channel="msedge")
    page = context.new_page()
    page.goto("https://app.powerbi.com/groups/me/reports/470121da-3902-4467-a4a4-85cba55102be/ReportSection?experience=power-bi")

    # Espera que la página termine de cargar antes de buscar elementos
    # page.wait_for_load_state('networkidle')

    #page.wait_for_selector('role=menuitem[name="Export"]', state="visible", timeout=50000)
   #page.get_by_role("menuitem", name="Export").click()

   #page.wait_for_selector('role=menuitem[name="Export to CSV"]', state="visible", timeout=50000)
    page.get_by_test_id('collapse-pages-pane-btn').click()
    page.get_by_role("textbox", name="End date. Available input").fill("11/24/2025")
    container= page.locator('div[title="Assignation History"]')
    container.scroll_into_view_if_needed()
    container.hover(force=True)
    page.get_by_test_id("visual-more-options-btn").click()
  
    #page.mouse.wheel(500,0)
    #page.get_by_test_id("visual-more-options-btn").click()
    select=page.get_by_test_id("pbimenu-item.Export data")
    select.scroll_into_view_if_needed()
    #page.get_by_test_id("pbimenu-item.Export data").click()
    select.click()
   
    with page.expect_download(timeout=450000) as download_info:  # 45 segundos para la descarga
         page.get_by_test_id("export-btn").click()
      #page.get_by_role("menuitem", name="Export to CSV", exact=True).click()


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
