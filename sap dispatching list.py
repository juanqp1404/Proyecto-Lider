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



def fecha_filtro() -> str:
    """
    Retorna la fecha de ejecución en formato 'M/D/YYYY'.

    Regla:
    - En general: usa la fecha de hoy.
    - Excepción: si es lunes y la hora actual es <= 7:30 AM,
      usa la fecha del viernes anterior.
    """
    ahora = datetime.now()
    fecha = ahora.date()

    # Lunes = 0 (lunes..domingo = 0..6)
    if ahora.weekday() == 0 and ahora.time() <= datetime.strptime("07:30", "%H:%M").time():
        # Lunes → viernes anterior (3 días antes)
        fecha = (ahora - timedelta(days=3)).date()

    mes = fecha.month
    dia = fecha.day
    anio = fecha.year

    return f"{mes}/{dia}/{anio}"

def xlsx_a_csv(ruta_xlsx: str, ruta_csv: str, hoja: str | int = 0) -> None:
    """
    Convierte un XLSX a CSV sin dañar el contenido.
    - hoja: nombre de la hoja o índice (0 = primera).
    """
    df = pd.read_excel(ruta_xlsx, sheet_name=hoja, engine="openpyxl")
    df.to_csv(ruta_csv, index=False, encoding="utf-8-sig")  # utf-8 con BOM para Excel

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

    fecha = fecha_filtro()

    page.get_by_test_id('collapse-pages-pane-btn').click()
    page.get_by_role("textbox", name="End date. Available input").fill(fecha)
    container= page.locator('div[title="Assignation History"]')
    container.scroll_into_view_if_needed()
    container.hover(force=True)
    page.get_by_test_id("visual-more-options-btn").click()
  
    select=page.get_by_test_id("pbimenu-item.Export data")

    select.scroll_into_view_if_needed()
    select.click()
    with page.expect_download(timeout=450000) as download_info:  # 45 segundos para la descarga
         page.get_by_test_id("export-btn").click()
    
    download = download_info.value

    new_filename = "sap_dispatching_list.xlsx"
    ruta_xlsx = os.path.join(carpeta_actual, "data/sharepoint/", new_filename)
    ruta_csv = os.path.join(carpeta_actual, "data/sharepoint/", "sap_dispatching_list.csv")

    print(f"Nombre sugerido del archivo: {download.suggested_filename}")
    print(f"Descargando XLSX a: {ruta_xlsx}")
    download.save_as(ruta_xlsx)
    
    # Convertir inmediatamente a CSV

    try:
        df = pd.read_excel(ruta_xlsx, engine="openpyxl")
        df.to_csv(ruta_csv, index=False, encoding="utf-8-sig")
        print(f"CSV guardado en: {ruta_csv}")
        # os.remove(ruta_xlsx)
        
    except Exception as e:
        print(f"Error convirtiendo: {e}")

    context.close()

kill_edge_processes()

with sync_playwright() as playwright:
    run(playwright)
