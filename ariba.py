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


def run(playwright: Playwright) -> None:
    # browser = playwright.chromium.launch(headless=False,executable_path='C:/Program Files/Google/Chrome/Application/chrome.exe')
    browser = playwright.chromium.launch_persistent_context(headless=False,executable_path="C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe")
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://s1.ariba.com/Sourcing/Main/aw?awh=r&awssk=ODTzCxAIbpbHwlV1&realm=schlumberger")
    # page.goto("https://example.com/")
    # page.get_by_role("heading", name="Example Domain").click()
    # page.get_by_role("link", name="More information...").click()
    # page.get_by_role("link", name="IANA-managed Reserved Domains").click()
    # page.get_by_role("link", name="XN--HLCJ6AYA9ESC7A").click()

    # ---------------------
    time.sleep(15)

    page.fill('#i0116','JQuintero27@slb.com')
    time.sleep(305)

    context.close()
    browser.close()

kill_edge_processes()

with sync_playwright() as playwright:
    run(playwright)
