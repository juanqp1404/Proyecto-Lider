"""
Plantilla de automatización de Power BI
Modifica las secciones según lo que necesites
"""
from pywinauto.application import Application
from utils.bi_utils import abrir_o_conectar_powerbi
from utils.acciones_powerbi import (
    refrescar_datos,
    guardar_archivo,
    exportar_visual_coordenadas,
    esperar_carga,
    refrescar_ps_dispatching,
    refrescar_sap_buyers,
    exportar_assignation_history,
    cerrar_powerbi
)


def main():
    # ========================================
    # 1. CONFIGURACIÓN (modifica aquí)
    # ========================================
    archivo = r"C:\Users\JQuintero27\Downloads\Workload Dispatching 6 (1) 1.pbix"
    
    
    # ========================================
    # 2. CONECTAR A POWER BI
    # ========================================
    pid = abrir_o_conectar_powerbi(archivo, tiempo_espera=25)
    
    if not pid:
        print("✗ Error: No se pudo conectar a Power BI")
        return
    
    print(f"→ Conectando (PID: {pid})...")
    app = Application(backend="uia").connect(process=pid)
    print("✓ Conectado!\n")
    
    
    # ========================================
    # 3. ESPERAR A QUE CARGUE EL DASHBOARD
    # ========================================
    esperar_carga(10)
    
    
    # ========================================
    # 4. ACCIONES (modifica según necesites)
    # ========================================
    # refrescar_ps_dispatching()
    # esperar_carga(420)
    # refrescar_sap_buyers()
    # esperar_carga(25)
    exportar_assignation_history()
    # esperar_carga(10)
    # Ejemplo 1: Refrescar datos
    # refrescar_datos()
    
    # Ejemplo 2: Exportar datos de un visual
    # (Primero identifica las coordenadas del gráfico)
    # exportar_visual_coordenadas(x=715, y=442, nombre_archivo="datos_conflicto.csv")
    
    # Ejemplo 3: Guardar archivo
    # guardar_archivo()
    
    # Ejemplo 4: Exportar múltiples visuales
    # exportar_visual_coordenadas(x=500, y=400, nombre_archivo="mapa_departamentos.csv")
    # exportar_visual_coordenadas(x=900, y=400, nombre_archivo="grafico_hechos.csv")
    
    
    # ========================================
    # 5. FINALIZAR
    # ========================================
    print("\n✓ Proceso completado!")
    
    # Opcional: cerrar Power BI
    # cerrar_powerbi()


if __name__ == "__main__":
    main()
