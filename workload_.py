import os
from datetime import datetime, time
from pathlib import Path
import pandas as pd

pd.set_option('display.max_columns', None)


with open("workload_debug.txt", "w", encoding="utf-8") as f:
    f.write(f"CWD: {os.getcwd()}\n")
    f.write(f"__file__: {__file__}\n")
    f.write(f"SCRIPT_DIR: {Path(__file__).parent.resolve()}\n")
    f.write(f"./data/sharepoint: {Path('./data/sharepoint').exists()}\n")
    f.write(f"sap_buyers.csv: {Path('./data/sharepoint/sap_buyers.csv').exists()}\n")
    f.write(f"sap_buyers RESUELTA: {Path('./data/sharepoint/sap_buyers.csv').resolve()}\n")
    f.write(f"TAMAÑO sap_buyers: {Path('./data/sharepoint/sap_buyers.csv').stat().st_size if Path('./data/sharepoint/sap_buyers.csv').exists() else 'NO'}\n")

print("DEBUG guardado en workload_debug.txt")  # ← ESTO SÍ sale en orquestador

def ensure_output_dir(path: str = "./data/final") -> None:
    os.makedirs(path, exist_ok=True)


def parse_today_fixed() -> datetime:
    """
    Para pruebas: fija 'hoy' al 2025-11-26 09:45.
    En producción, cambia a: return datetime.now()
    """
    # return datetime(2025, 11, 26, 9, 45)
    # return datetime(2025, 12, 2, 9, 45)
    return datetime.now()


def load_data(
    buyers_path: str = "./data/sharepoint/sap_buyers.csv",
    dispatching_path: str = "./data/sharepoint/sap_dispatching_list.csv",
) -> tuple[pd.DataFrame, pd.DataFrame]:
    
    with open("load_data_debug.txt", "w", encoding="utf-8") as f:
        f.write("=== LOAD_DATA DEBUG ===\n")
        f.write(f"1. cwd: {os.getcwd()}\n")
        f.write(f"2. buyers_path: {Path(buyers_path).resolve()}\n")
        
        # TEST 1: Lectura cruda de bytes
        try:
            with open(buyers_path, 'rb') as file:
                first_bytes = file.read(100)
                has_bom = first_bytes.startswith(b'\xef\xbb\xbf')
                f.write(f"3. Primeros 100 bytes (hex): {first_bytes.hex()[:50]}...\n")
                f.write(f"4. Detecta BOM: {has_bom}\n")  # ← Fix: variable fuera de f-string
        except Exception as e:
            f.write(f"3. Error lectura cruda: {e}\n")
        
        f.write("=== FIN DEBUG ===\n")
    
    print("DEBUG guardado en load_data_debug.txt")
    
    # Carga FINAL robusta
    df_buyers = pd.read_csv(buyers_path, encoding='utf-8-sig', low_memory=False, on_bad_lines='skip')
    df_dispatching = pd.read_csv(dispatching_path, encoding='utf-8-sig', low_memory=False, on_bad_lines='skip')
    
    return df_buyers, df_dispatching



def normalize_dispatching_dates(df_dispatching: pd.DataFrame) -> pd.DataFrame:
    """
    Normaliza la columna Created de df_dispatching a tipo datetime (solo fecha).
    Asume que Created viene como texto 'mm/dd/YYYY HH:MM AM/PM' o con hora opcional.
    """
    df = df_dispatching.copy()

    # Asegurar string y quedarnos solo con la parte de fecha
    df["Created"] = df["Created"].astype(str).str.split(" ").str[0]

    # Dejar que pandas infiera el formato
    df["Created"] = pd.to_datetime(df["Created"], errors="coerce")

    if df["Created"].isna().any():
        bad_rows = df[df["Created"].isna()]
        raise ValueError(
            f"Hay filas en df_dispatching con fecha 'Created' inválida:\n"
            f"{bad_rows[['SC Number', 'Created']].head()}"
        )

    df["Created_date"] = df["Created"].dt.date
    return df


def filter_dispatching_for_today(df_dispatching: pd.DataFrame, today: datetime) -> pd.DataFrame:
    """
    Filtra df_dispatching para quedarte solo con las filas de 'hoy' (por fecha).
    """
    df = normalize_dispatching_dates(df_dispatching)
    today_date = today.date()
    df_today = df[df["Created_date"] == today_date].copy()
    return df_today


def determine_execution_shift(now: datetime) -> str:
    """
    Determina el shift de ejecución, solo para metadata.
    Ejemplo simple:
      - 07:00–17:00 -> '7 to 5'
      - 09:00–19:00 -> '9 to 7'
      - resto -> 'off-shift'
    Ajusta las reglas si hace falta.
    """
    current_time = now.time()

    if time(7, 0) <= current_time < time(17, 0):
        return "7 to 5"
    elif time(9, 0) <= current_time < time(19, 0):
        return "9 to 7"
    else:
        return "off-shift"


def clean_percentage_column(series: pd.Series) -> pd.Series:
    """
    Convierte valores tipo '100%' o '0%' a float 1.0, 0.0, etc.
    Si no se puede convertir, devuelve NaN.
    """
    return (
        series.astype(str)
        .str.strip()
        .str.replace("%", "", regex=False)
        .replace({"": None, "nan": None})
        .astype(float)
        / 100.0
    )

def create_workload_by_subcategory(df_buyers, df_dispatching_today, subcategory_value):
    df_b = df_buyers.copy()
    df_b.columns = [c.strip() for c in df_b.columns]
    
    # Fix NaN en Sub-Category
    mask = df_b["Sub-Category"].astype(str).str.strip() == subcategory_value
    df_b_sub = df_b[mask].copy()
    
    if df_b_sub.empty:
        return pd.DataFrame()
    
    # Fix clean_percentage_column
    series_clean = clean_percentage_column(df_b_sub["Workload / Availability"])
    df_b_sub["Availability_SAP"] = series_clean.fillna(0)
    df_b_sub["Urgent_enabled"] = df_b_sub["Available For Urgencies"].astype(str).str.strip()
    
    # Fix filtro
    df_b_active = df_b_sub[df_b_sub["Availability_SAP"] > 0].copy()
    
    if df_b_active.empty:
        return pd.DataFrame()
    
    # Fix groupby vacío
    if df_dispatching_today.empty:
        df_counts = pd.DataFrame({"Buyer Alias": [], "Count of SC Number": []})
    else:
        df_d = df_dispatching_today.copy()
        df_d.columns = [c.strip() for c in df_d.columns]
        df_counts = df_d.groupby("Buyer Alias", as_index=False)["SC Number"].count()
        df_counts = df_counts.rename(columns={"SC Number": "Count of SC Number"})
    
    # Merge seguro
    df_workload = df_b_active.merge(df_counts, on="Buyer Alias", how="left")
    df_workload["Count of SC Number"] = df_workload["Count of SC Number"].fillna(0)
    
    cols = ["Buyer Alias", "Count of SC Number", "Availability_SAP", "Shift", "Urgent_enabled"]
    return df_workload[cols].sort_values("Count of SC Number", ascending=False).reset_index(drop=True)





def main() -> None:
    # 1. Preparar salida
    output_dir = "./data/final"
    ensure_output_dir(output_dir)

    # 2. Definir fecha/hora de ejecución
    hoy = parse_today_fixed()
    execution_shift = determine_execution_shift(hoy)
    print(f"Ejecutando ETL para fecha: {hoy.date()}, hora: {hoy.time()}, shift_detectado: {execution_shift}")

    # 3. Cargar datos
    df_buyers, df_dispatching = load_data()

    # 4. Filtrar dispatching a 'hoy'
    df_dispatching_today = filter_dispatching_for_today(df_dispatching, hoy)

    # 5. Crear workloads para CAM IND NAM y CAM IND LAM
    df_workload_nam = create_workload_by_subcategory(
        df_buyers=df_buyers,
        df_dispatching_today=df_dispatching_today,
        subcategory_value="CAM IND NAM",
    )

    df_workload_lam = create_workload_by_subcategory(
        df_buyers=df_buyers,
        df_dispatching_today=df_dispatching_today,
        subcategory_value="CAM IND LAM",
    )

    # 6. Exportar CSVs finales
    nam_path = os.path.join(output_dir, "workload_cam_ind_nam.csv")
    lam_path = os.path.join(output_dir, "workload_cam_ind_lam.csv")

    df_workload_nam.to_csv(nam_path, index=False, encoding="utf-8")
    df_workload_lam.to_csv(lam_path, index=False, encoding="utf-8")

    print(f"Exportado: {nam_path}")
    print(f"Exportado: {lam_path}")


if __name__ == "__main__":
    main()
