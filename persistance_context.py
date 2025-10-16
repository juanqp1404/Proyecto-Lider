import re
from playwright.sync_api import Playwright, sync_playwright, expect
import time


def run(playwright: Playwright) -> None:
    # Especifica la ruta de tu User Data Directory de Chrome
    # user_data_dir = r"C:/Users/xanti/AppData/Local/Google/Chrome/User Data/"
    user_data_dir = r"C:/Users/xanti/AppData/Local/Microsoft/Edge/User Data"
    # user_data_dir = r"C:\Users\xanti\Desktop\chrome_profile_temp"
    
    # Usa launch_persistent_context en lugar de launch
    print("Abriendo contexto persistente...")
    context = playwright.chromium.launch_persistent_context(
        user_data_dir,
        channel="msedge",
        headless=False,
        # args= [
            # '--profile-directory=Profile 3',
            # '--disable-features=DevToolsDebuggingRestrictions',
            # '--app=https://www.youtube.com'
        # ]
        # Para buscar el nombre del perfil, ve a chrome://version/ y busca "Perfil de ruta"
    )
    print("Contexto persistente abierto.")
    # Con persistent context, ya tienes el contexto creado
    # y puedes acceder a la página directamente

    # print(context.pages)
    # if context.pages:
    #     page = context.pages[0]  # Usa la página about:blank existente
    # else:
    #     page = context.new_page()

    print("Páginas No abiertas:")
    page = context.new_page()
    print("Páginas abiertas:")
   
    page.goto("https://www.youtube.com/", wait_until="networkidle")
    page.wait_for_timeout(8000)  # Espera 5 segundos para asegurarte de que la página cargue completamente
    page.get_by_role("link", name="Tú").click()
    # page.wait_for_timeout(5000)
    # page.locator("button-view-model").filter(has_text="Ver todo").get_by_label("Ver todo").click()
    # page.wait_for_timeout(5000)
    # page.get_by_title("Listas de reproducción").get_by_role("link").click()
    page.wait_for_timeout(8000)
    # page.get_by_text("CA7RIEL & Paco Amoroso - BABY GANGSTA (Visualizer)Activar el").click()
    page.locator("ytd-rich-shelf-renderer").filter(has_text="Historial Ver todo 2:").get_by_label("CA7RIEL & Paco Amoroso - BABY").click()

    page.wait_for_timeout(8000)
    # ---------------------
    context.close()


with sync_playwright() as playwright:
    run(playwright)
