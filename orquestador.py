#!/usr/bin/env python3
"""
Orquestador ETL Cameron Indirect - Windows compatible.
"""

import subprocess
import sys
import time
from pathlib import Path
from datetime import datetime
import logging

# Logs Windows-friendly
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(message)s',
    handlers=[
        logging.FileHandler('./logs/orquestador.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ],
    encoding='utf-8'  # Para Windows
)
logger = logging.getLogger(__name__)

Path("./logs").mkdir(exist_ok=True)

FLUJOS = [
    ("ariba", ["python", "ariba.py"]),
    ("manual", ["jupyter", "execute", "Cameron Indirect Manual Assignment Tool.ipynb"]),
    ("sap_buyers", ["python", "sap buyers.py"]),
    ("sap_dispatching", ["python", "sap dispatching list.py"]),
    ("workload", ["python", "workload_.py"]),
    ("asignaciones", ["python", "asignaciones.py"]),
]

MAX_RETRIES = 3
RETRY_DELAY = 15

def ejecutar_con_retry(nombre: str, comando: list) -> bool:
    for intento in range(1, MAX_RETRIES + 1):
        logger.info(f"[INFO] [{nombre}] Intento {intento}/{MAX_RETRIES}")
        
        try:
            resultado = subprocess.run(
                comando, capture_output=True, text=True, timeout=300, cwd="."
            )
            
            if resultado.returncode == 0:
                logger.info(f"[OK] [{nombre}] Completado")
                return True
            else:
                logger.error(f"[ERROR] [{nombre}] Return code {resultado.returncode}")
                if resultado.stderr:
                    logger.error(f"[STDERR] {resultado.stderr[:300]}...")
                
        except subprocess.TimeoutExpired:
            logger.error(f"[TIMEOUT] [{nombre}] 5min excedido")
        except FileNotFoundError as e:
            logger.error(f"[NOT FOUND] [{nombre}] {e}")
        except Exception as e:
            logger.error(f"[EXCEPTION] [{nombre}] {e}")
        
        if intento < MAX_RETRIES:
            logger.info(f"[WAIT] [{nombre}] Retry en {RETRY_DELAY}s...")
            time.sleep(RETRY_DELAY)
    
    logger.error(f"[FAILED] [{nombre}] Falló definitivamente")
    return False

def main():
    logger.info("=" * 60)
    logger.info(f"ORQUESTADOR ETL - {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    logger.info("=" * 60)
    
    exitosos = 0
    fallidos = []
    
    for nombre, comando in FLUJOS:
        if ejecutar_con_retry(nombre, comando):
            exitosos += 1
        else:
            fallidos.append(nombre)
    
    logger.info("=" * 60)
    logger.info(f"RESULTADO: {exitosos}/6 OK | Fallidos: {fallidos}")
    logger.info("Salidas: ./final/ | Logs: ./logs/orquestador.log")
    logger.info("=" * 60)
    
    sys.exit(0 if exitosos == 6 else 1)

if __name__ == "__main__":
    main()
