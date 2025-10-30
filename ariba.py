from datetime import datetime, timedelta
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

def run(playwright: Playwright) -> None:
    fecha_filtrado = calcular_domingo_asociado()
    user_data_dir = r"C:/Users/JQuintero27/AppData/Local/Microsoft/Edge/User Data"
    carpeta_actual = "./"
    # browser = playwright.chromium.launch(headless=False,executable_path='C:/Program Files/Google/Chrome/Application/chrome.exe')
    context = playwright.chromium.launch_persistent_context(user_data_dir, headless=False, channel="msedge")
    page = context.new_page()
    page.goto("https://s1.ariba.com/Sourcing/Main/aw?awh=r&awssk=ODTzCxAIbpbHwlV1&realm=schlumberger")
    # time.sleep(40)
    page.wait_for_selector("#_c19zzd")
    page.get_by_role("button", name="Manage").click()
    time.sleep(2)
    page.get_by_role("menuitem", name="Queues").click()
    page.wait_for_selector("#text__c_npbb")
    page.locator("#text__c_npbb").click()
    page.wait_for_selector("#text__c_npbb")
    page.get_by_role("option", name="Requisition").click()
    time.sleep(2)
    page.get_by_role("textbox", name="From:").fill(fecha_filtrado)
    time.sleep(2)
    page.get_by_title("Run this search").click()
    page.wait_for_selector("#_7msd8 > div")
    # page.get_by_role("button", id="_7msd8").click()
    page.query_selector('#_7msd8 > div').click()
    time.sleep(2)

    with page.expect_download() as download_info:
        page.get_by_role("menuitem", name="Export all Rows").click()

    download = download_info.value

    new_filename = "DF.xls"
    print(f"Nombre sugerido del archivo: {download.suggested_filename}")
    print(f"Descargando el archivo como: {new_filename}")
    print(f"URL del archivo: {download}")
    time.sleep(20)
    # Guardar el archivo en la carpeta actual con su nombre original
    ruta_destino = os.path.join(carpeta_actual,"data/", new_filename) #./data/DF.xls
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
