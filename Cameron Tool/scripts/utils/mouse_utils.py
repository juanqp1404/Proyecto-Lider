import pyautogui
import time

def capturar_posicion():
    """Captura la posición actual del mouse"""
    x, y = pyautogui.position()
    return x, y


def mostrar_posicion_continua(duracion=10):
    """Muestra la posición del mouse continuamente"""
    print(f"Mostrando posición durante {duracion} segundos...")
    print("Mueve el mouse por la pantalla:\n")
    
    inicio = time.time()
    while time.time() - inicio < duracion:
        x, y = capturar_posicion()
        print(f"\rPosición: ({x:4d}, {y:4d})", end="", flush=True)
        time.sleep(0.1)
    
    print("\n✓ Captura finalizada")


def capturar_clicks(num_clicks=3):
    """
    Captura coordenadas cada vez que haces clic
    
    Args:
        num_clicks: Número de clics a capturar
    """
    print(f"Harás {num_clicks} clics. Se capturará la posición antes de cada clic.")
    print("Presiona Enter para iniciar...\n")
    input()
    
    coordenadas = []
    
    for i in range(num_clicks):
        print(f"\n[{i+1}/{num_clicks}] Mueve el mouse al elemento y presiona Enter...")
        input()
        
        x, y = capturar_posicion()
        coordenadas.append((x, y))
        print(f"  ✓ Capturado: ({x}, {y})")
    
    print("\n" + "="*50)
    print("COORDENADAS CAPTURADAS:")
    print("="*50)
    
    for i, (x, y) in enumerate(coordenadas, 1):
        print(f"{i}. pyautogui.click({x}, {y})")
    
    return coordenadas


def menu_interactivo():
    """Menú interactivo para probar coordenadas"""
    while True:
        print("\n" + "="*50)
        print("HERRAMIENTA DE COORDENADAS")
        print("="*50)
        print("1. Ver posición actual del mouse")
        print("2. Mostrar posición continuamente (10 segundos)")
        print("3. Capturar múltiples clics")
        print("4. Salir")
        print("="*50)
        
        opcion = input("\nElige una opción (1-4): ").strip()
        
        if opcion == "1":
            x, y = capturar_posicion()
            print(f"✓ Posición actual: ({x}, {y})")
        
        elif opcion == "2":
            duracion = input("¿Cuántos segundos? (default 10): ").strip()
            duracion = int(duracion) if duracion.isdigit() else 10
            mostrar_posicion_continua(duracion)
        
        elif opcion == "3":
            num_clicks = input("¿Cuántos clics? (default 3): ").strip()
            num_clicks = int(num_clicks) if num_clicks.isdigit() else 3
            coordenadas = capturar_clicks(num_clicks)
            
            # Guardar en archivo
            guardar = input("\n¿Guardar en archivo? (s/n): ").strip().lower()
            if guardar == 's':
                with open("coordenadas_capturadas.py", "w") as f:
                    f.write("# Coordenadas capturadas\n\n")
                    for i, (x, y) in enumerate(coordenadas, 1):
                        f.write(f"coord_{i} = ({x}, {y})\n")
                print("✓ Guardado en: coordenadas_capturadas.py")
        
        elif opcion == "4":
            print("✓ Saliendo...")
            break
        
        else:
            print("✗ Opción no válida")


if __name__ == "__main__":
    menu_interactivo()
