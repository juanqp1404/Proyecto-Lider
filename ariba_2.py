from datetime import datetime, timedelta
import os
import re
import time
from typing import Optional
import psutil
from playwright.sync_api import Playwright, sync_playwright, expect, TimeoutError


def kill_edge_processes():
    """Cierra procesos de Edge de forma segura."""
    edge_process_names = ["msedge.exe", "msedgedriver.exe"]
    processes_killed = False
    
    print("🔍 Verificando procesos de Edge...")
    
    for proc in psutil.process_iter(['name', 'pid']):
        try:
            if proc.info['name'] in edge_process_names:
                print(f"🛑 Cerrando {proc.info['name']} (PID: {proc.info['pid']})")
                proc.terminate()  # Más suave que kill()
                processes_killed = True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess) as e:
            continue
    
    if processes_killed:
        print("⏳ Esperando 3s para cierre completo...")
        time.sleep(3)
    else:
        print("✅ No hay procesos Edge activos.")


def calcular_domingo_asociado(fecha=None) -> str:
    """Calcula domingo asociado con formato 'Sun, 7 Sep, 2025'."""
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
    
    return domingo_resultado.strftime('%a, %e %b, %Y')  # %e sin cero inicial


def wait_with_retry(page, selector: str, max_retries: int = 5, base_delay: float = 1.0) -> None:
    """Espera un selector con retry exponencial."""
    for attempt in range(max_retries):
        try:
            page.wait_for_selector(selector, state='visible', timeout=5000)
            return
        except TimeoutError:
            if attempt == max_retries - 1:
                raise
            delay = base_delay * (2 ** attempt)  # Backoff exponencial
            print(f"⚠️  Retry {attempt + 1}/{max_retries} para '{selector}' (esperando {delay:.1f}s)")
            time.sleep(delay)


def safe_click_with_retry(page, selector: str, max_retries: int = 5) -> None:
    """Clic seguro con scroll, hover y retry."""
    for attempt in range(max_retries):
        try:
            # Scroll y hover para asegurar visibilidad
            elem = page.locator(selector).first
            elem.scroll_into_view_if_needed()
            elem.hover(timeout=3000)
            time.sleep(0.5)  # Pequeña pausa post-hover
            
            elem.click(timeout=5000)
            return
        except TimeoutError:
            if attempt == max_retries - 1:
                raise
            print(f"⚠️  Retry clic {attempt + 1}/{max_retries} para '{selector}'")
            time.sleep(1)


def run(playwright: Playwright) -> None:
    """Ejecución principal robusta."""
    os.makedirs("./data", exist_ok=True)
    
    fecha_filtrado = calcular_domingo_asociado()
    print(f"📅 Fecha de filtrado calculada: {fecha_filtrado}")
    
    user_data_dir = r"C:/Users/JQuintero27/AppData/Local/Microsoft/Edge/User Data"
    
    # Lanzamiento robusto del browser
    try:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir, 
            headless=False, 
            channel="msedge",
            args=[
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-gpu",
                "--disable-extensions",
                #"--window-size=1920,1080"
            ]
        )
        page = context.new_page()
        #page.set_viewport_size({"width": 1920, "height": 1080})
        
        print("🌐 Navegando a Ariba...")
        page.goto("https://s1.ariba.com/Sourcing/Main/aw?awh=r&awssk=ODTzCxAIbpbHwlV1&realm=schlumberger", 
                 wait_until="domcontentloaded", timeout=30000)
        
        # PASO 1: Manage -> Queues
        print("📋 Click Manage > Queues...")
        wait_with_retry(page, "role=button[name='Manage']")
        safe_click_with_retry(page, "role=button[name='Manage']")
        
        wait_with_retry(page, "role=menuitem[name='Queues']")
        safe_click_with_retry(page, "role=menuitem[name='Queues']")
        
        # PASO 2: Filtro Requisition
        print("🔍 Configurando filtro Requisition...")
        wait_with_retry(page, "#text__c_npbb")
        page.locator("#text__c_npbb").click()
        
        wait_with_retry(page, "role=option[name='Requisition']")
        page.get_by_role("option", name="Requisition").click()
        
        # PASO 3: Fecha From
        print(f"📆 Ingresando fecha: {fecha_filtrado}")
        wait_with_retry(page, "role=textbox[name='From:']")
        from_field = page.get_by_role("textbox", name="From:")
        from_field.fill("")
        from_field.fill(fecha_filtrado)
        
        # PASO 4: Run search
        print("▶️ Ejecutando búsqueda...")
        wait_with_retry(page, "[title='Run this search']")
        safe_click_with_retry(page, "[title='Run this search']")
        
        # Esperar resultados (método más robusto)
        page.wait_for_load_state('networkidle', timeout=15000)
        print("⏳ Esperando que carguen resultados...")
        time.sleep(3)  # Pausa mínima para estabilidad
        
        # PASO 5: Export
        print("📥 Exportando tabla...")
        wait_with_retry(page, "div[title='Table Options Menu']", max_retries=3)
        safe_click_with_retry(page, "div[title='Table Options Menu']")
        
        wait_with_retry(page, "role=menuitem[name='Export all Rows']", max_retries=3)
        
        # Descarga robusta SIN sleep(20)
        download_path = "./data/"
        download = page.expect_download(timeout=30000)
        safe_click_with_retry(page, "role=menuitem[name='Export all Rows']")
        
        new_filename = "DF.xls"
        ruta_destino = os.path.join(download_path, new_filename)
        
        print(f"💾 Guardando como: {ruta_destino}")
        download.save_as(ruta_destino)
        print("✅ Descarga completada exitosamente!")
        
        context.close()
        
    except Exception as e:
        print(f"💥 Error crítico: {e}")
        print("🔄 Intentando cierre limpio...")
        try:
            context.close()
        except:
            pass
        raise


if __name__ == "__main__":
    print("🚀 Iniciando scraping Ariba robusto...")
    kill_edge_processes()
    
    try:
        with sync_playwright() as playwright:
            run(playwright)
        print("🎉 Proceso completado sin errores!")
    except KeyboardInterrupt:
        print("\n⏹️  Proceso interrumpido por usuario.")
    except Exception as e:
        print(f"\n💥 Fallo final: {e}")
        print("💡 Revisa logs arriba y ejecuta kill_edge_processes() manualmente si es necesario.")
