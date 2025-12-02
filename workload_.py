import os
from datetime import datetime, time

import pandas as pd


def ensure_output_dir(path: str = "./data/final") -> None:
    os.makedirs(path, exist_ok=True)


def parse_today_fixed() -> datetime:
    """
    Para pruebas: fija 'hoy' al 2025-11-26 09:45.
    En producción, cambia a: return datetime.now()
    """
    #return datetime(2025, 11, 26, 9, 45)
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
    Crea df_workload para un subcategory específico (CAM IND NAM o CAM IND LAM).
    - Filtra buyers por Sub-Category.
    - Se queda solo con Workload / Availability > 0.
    - Cuenta SC Number por Buyer Alias en df_dispatching_today.
    - Trae Availability_SAP desde df_buyers (Workload / Availability).
    - Añade Shift.
    - Añade Urgent_enabled manteniendo 'Yes'/'No' de 'Available For Urgencies'.
    """

    df_b = df_buyers.copy()

    # Normalizar nombres de columnas
    df_b.columns = [c.strip() for c in df_b.columns]

    # Columnas necesarias en df_buyers
    required_cols = {
        "Buyer Alias",
        "Sub-Category",
        "Workload / Availability",
        "Shift",
        "Available For Urgencies",
    }
    missing = required_cols - set(df_b.columns)
    if missing:
        raise KeyError(f"Faltan columnas en df_buyers: {missing}")

    # Filtrar por subcategoría (CAM IND NAM o CAM IND LAM)
    df_b_sub = df_b[df_b["Sub-Category"] == subcategory_value].copy()

    # Limpiar Workload / Availability como porcentaje numérico
    df_b_sub["Availability_SAP"] = clean_percentage_column(df_b_sub["Workload / Availability"])

    # Filtrar solo los que tengan disponibilidad > 0
    df_b_sub = df_b_sub[df_b_sub["Availability_SAP"] > 0]

    if df_b_sub.empty:
        raise ValueError(f"No hay buyers con Availability_SAP > 0 para Sub-Category = {subcategory_value}")

    # Normalizar Available For Urgencies: dejar 'Yes'/'No' limpio en Urgent_enabled # !BUG: Solo salen los que tienen "Yes" en urgent_enabled
    df_b_sub["Urgent_enabled"] = df_b_sub["Available For Urgencies"].astype(str).str.strip()

    # df_dispatching_today: contar SC Number por Buyer Alias
    df_d = df_dispatching_today.copy()
    df_d.columns = [c.strip() for c in df_d.columns]

    required_disp_cols = {"SC Number", "Buyer Alias"}
    missing_disp = required_disp_cols - set(df_d.columns)
    if missing_disp:
        raise KeyError(f"Faltan columnas en df_dispatching: {missing_disp}")

    df_counts = (
        df_d.groupby("Buyer Alias", as_index=False)["SC Number"]
        .count()
        .rename(columns={"SC Number": "Count of SC Number"})
    )

    # Merge con buyers filtrados por subcategoría
    df_workload = df_counts.merge(
        df_b_sub[["Buyer Alias", "Availability_SAP", "Shift", "Urgent_enabled"]],
        on="Buyer Alias",
        how="left",
    )

    # Quitamos filas donde no se encontró buyer válido
    df_workload = df_workload.dropna(subset=["Availability_SAP"])

    # Ordenar por carga descendente (opcional)
    df_workload = df_workload.sort_values("Count of SC Number", ascending=False).reset_index(drop=True)

    return df_workload


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
