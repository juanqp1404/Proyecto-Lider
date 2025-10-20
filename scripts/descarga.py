import os
from playwright.sync_api import Playwright, sync_playwright

def run(playwright: Playwright) -> None:
    # Obtener la ruta absoluta de la carpeta del script
    # carpeta_actual = os.getcwd()
    carpeta_actual = './'

    browser = playwright.chromium.launch(headless=False, slow_mo=1000)
    # Configurar el contexto para aceptar descargas en la carpeta actual
    context = browser.new_context(accept_downloads=True)
    page = context.new_page()
    page.goto("https://the-internet.herokuapp.com/download")

    # Interceptar la descarga antes del clic
    with page.expect_download() as download_info:
        page.get_by_role("link", name="Blue Simple Professional CV").click()
    download = download_info.value

    # dowwload_name = download.suggested_filename
    new_filename = "Mi_CV_Descargado.pdf"
    
    print(f"Nombre sugerido del archivo: {download.suggested_filename}")
    print(f"Descargando el archivo como: {new_filename}")
    print(f"URL del archivo: {download}")

    # Guardar el archivo en la carpeta actual con su nombre original
    ruta_destino = os.path.join(carpeta_actual, new_filename)
    download.save_as(ruta_destino)

    print(f"Archivo descargado en: {ruta_destino}")

    context.close()
    browser.close()

with sync_playwright() as playwright:
    run(playwright)
