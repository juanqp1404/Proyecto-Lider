import os
from datetime import datetime, time

import pandas as pd

pd.set_option('display.max_columns', None)

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
    df_buyers = pd.read_csv(buyers_path)
    df_dispatching = pd.read_csv(dispatching_path)
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

def create_workload_by_subcategory(
    df_buyers: pd.DataFrame,
    df_dispatching_today: pd.DataFrame,
    subcategory_value: str,
) -> pd.DataFrame:
    """
    Workload = TODOS buyers con Workload > 0% de la categoría (Yes + No urgencias).
    """
    df_b = df_buyers.copy()
    df_b.columns = [c.strip() for c in df_b.columns]

    # 1. TODOS buyers de la subcategoría
    df_b_sub = df_b[df_b["Sub-Category"] == subcategory_value].copy()
    
    # 2. Limpiar ANTES de filtrar
    df_b_sub["Availability_SAP"] = clean_percentage_column(df_b_sub["Workload / Availability"])
    df_b_sub["Urgent_enabled"] = df_b_sub["Available For Urgencies"].astype(str).str.strip()
    
    # 3. SOLO filtra Workload > 0% (NO toca urgencias)
    df_b_active = df_b_sub[df_b_sub["Availability_SAP"] > 0].copy()
    
    print(f"🔍 {subcategory_value}: {len(df_b_active)} buyers activos, "
          f"Urgentes: {len(df_b_active[df_b_active['Urgent_enabled']=='Yes'])}, "
          f"No urgentes: {len(df_b_active[df_b_active['Urgent_enabled']=='No'])}")
    
    # 4. Contar SC Number de dispatching (puede estar vacío)
    df_d = df_dispatching_today.copy()
    df_d.columns = [c.strip() for c in df_d.columns]
    
    df_counts = (
        df_d.groupby("Buyer Alias", as_index=False)["SC Number"]
        .count()
        .rename(columns={"SC Number": "Count of SC Number"})
    )
    
    # 5. MERGE INVERTIDO: parte de ACTIVE BUYERS, agrega counts (0 si no hay)
    df_workload = df_b_active.merge(
        df_counts, 
        on="Buyer Alias", 
        how="left"
    ).fillna({"Count of SC Number": 0})
    
    # 6. Columnas finales
    return df_workload.sort_values("Count of SC Number", ascending=False)[
        ["Buyer Alias", "Count of SC Number", "Availability_SAP", "Shift", "Urgent_enabled"]
    ].reset_index(drop=True)


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
