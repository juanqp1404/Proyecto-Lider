import os
import re
from datetime import datetime, time
from typing import List, Tuple

import pandas as pd


# ------------------ UTILIDADES GENERALES ------------------ 

def ensure_output_dir(path: str = "./data/final") -> None:
    os.makedirs(path, exist_ok=True)


def parse_today_fixed() -> datetime:
    """
    Para pruebas: fija 'hoy' al 2025-11-26 09:45.
    En producción, cambia a: return datetime.now()
    """
    #return datetime(2025, 11, 26, 9, 45)
    # return datetime(2025, 12, 2, 9, 45)
    return datetime.now()


def load_workloads(
    lam_path: str = "./data/final/workload_cam_ind_lam.csv",
    nam_path: str = "./data/final/workload_cam_ind_nam.csv",
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    df_lam = pd.read_csv(lam_path)
    df_nam = pd.read_csv(nam_path)

    # Normalizar columnas clave
    for df in (df_lam, df_nam):
        df.columns = [c.strip() for c in df.columns]
        if "Buyer Alias" not in df.columns:
            raise KeyError("Los workloads deben contener la columna 'Buyer Alias'.")
        if "Shift" not in df.columns:
            raise KeyError("Los workloads deben contener la columna 'Shift'.")
        if "Urgent_enabled" not in df.columns:
            raise KeyError("Los workloads deben contener la columna 'Urgent_enabled' (Yes/No).")

    return df_lam, df_nam


def load_resultados(path: str = "./data/Resultado.xlsx") -> pd.DataFrame:
    df = pd.read_excel(path)
    df.columns = [c.strip() for c in df.columns]
    return df


def split_special_prs(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Separa los PR que terminan en -V2 o -V3 en un df aparte (ps_special).
    Usa la columna PR.
    """
    df = df.copy()
    df["ID"] = df["ID"].astype(str).str.strip()

    mask_special = df["ID"].str.contains(r"-V", case=False, na=False)
    df_special = df[mask_special].copy()
    df_normal = df[~mask_special].copy()

    return df_normal, df_special


# ------------------ SHIFTS Y URGENCIAS ------------------ #

def parse_shift_to_range(shift_str: str) -> Tuple[time, time]:
    """
    Convierte un string tipo '7 to 5', '9 to 7', '6 to 4' en (hora_inicio, hora_fin).
    Se interpreta como turno dentro del mismo día:
      '6 to 4' -> 06:00–16:00
      '7 to 5' -> 07:00–17:00
      '9 to 7' -> 09:00–19:00
    Si en algún momento tienes turnos que cruzan medianoche, habrá que ajustar.
    """
    if not isinstance(shift_str, str):
        return None, None

    s = shift_str.strip().lower()
    m = re.match(r"(\d+)\s*to\s*(\d+)", s)
    if not m:
        return None, None

    start_hour = int(m.group(1))
    end_hour = int(m.group(2))

    # Interpretación: office-style, mismo día
    start = time(start_hour, 0)
    end = time(end_hour, 0)
    return start, end


def is_in_shift(current_time: time, start: time, end: time) -> bool:
    """
    Devuelve True si current_time está dentro del rango [start, end) para turnos del mismo día.
    Si en el futuro hay turnos que cruzan medianoche, aquí se ajusta.
    """
    if start is None or end is None:
        return False

    # Turno normal (ej: 6->16, 7->17, 9->19)
    if start < end:
        return start <= current_time < end
    # Turno que cruza medianoche (ej: 18->4) -> NO usado por ahora, pero se deja listo
    else:
        return current_time >= start or current_time < end


def filter_buyers_by_shift(df_workload: pd.DataFrame, execution_time: time) -> pd.DataFrame:
    """
    Filtra buyers que estén dentro de su shift actual según la hora de ejecución.
    """
    df = df_workload.copy()
    available_rows = []

    for idx, row in df.iterrows():
        start, end = parse_shift_to_range(row.get("Shift"))
        if is_in_shift(execution_time, start, end):
            available_rows.append(idx)

    df_available = df.loc[available_rows].reset_index(drop=True)
    return df_available


def filter_buyers_by_urgency(df_workload: pd.DataFrame, urgent_required: bool) -> pd.DataFrame:
    """
    Si urgent_required=True, solo deja buyers con Urgent_enabled == 'Yes'.
    Si urgent_required=False, deja todos los buyers (independientemente de urgencias).
    """
    df = df_workload.copy()
    df["Urgent_enabled"] = df["Urgent_enabled"].astype(str).str.strip()

    if urgent_required:
        df = df[df["Urgent_enabled"].str.upper() == "YES"].copy()

    return df.reset_index(drop=True)


# ------------------ ASIGNACIÓN EQUITATIVA (ROUND-ROBIN) ------------------ #

def round_robin_assign(
    tasks: pd.DataFrame,
    buyers: List[str],
    buyer_column_name: str = "BUYER",
) -> pd.DataFrame:
    """
    Asigna buyers a las filas de tasks de forma equitativa (round-robin).
    - tasks: df con las PR pendientes.
    - buyers: lista de Buyer Alias disponibles.
    - buyer_column_name: nombre de la columna donde se escribirá el buyer asignado.
    """
    if not buyers:
        raise ValueError("No hay buyers disponibles para asignar.")

    tasks = tasks.copy()

    # Asegurar que la columna existe y es string
    if buyer_column_name not in tasks.columns:
        tasks[buyer_column_name] = pd.Series([None] * len(tasks), dtype="object")
    else:
        tasks[buyer_column_name] = tasks[buyer_column_name].astype("object")

    num_buyers = len(buyers)

    for i in range(len(tasks)):
        buyer = buyers[i % num_buyers]
        tasks.iat[i, tasks.columns.get_loc(buyer_column_name)] = buyer

    return tasks

# ------------------ LÓGICA PRINCIPAL DE ASIGNACIÓN ------------------ #

def assign_buyers_for_region(
    df_resultados_region: pd.DataFrame,
    df_workload_region: pd.DataFrame,
    execution_time: time,
) -> pd.DataFrame:
    """
    Asigna buyers para una región (NAM o LAM):
    - Separa URGENT=1 y URGENT=0.
    - URGENT=1 -> solo buyers con Urgent_enabled == 'Yes'.
    - Respeta shift: solo buyers activos en la hora de ejecución.
    - Reparte equitativamente (round-robin) entre los buyers disponibles.
    """
    df = df_resultados_region.copy()
    df.columns = [c.strip() for c in df.columns]

    if "URGENT" not in df.columns:
        raise KeyError("df_resultados_region debe contener la columna 'URGENT'.")
    if "BUYER" not in df.columns:
        df["BUYER"] = None

    # Filtrar buyers por shift actual
    df_workload_shift = filter_buyers_by_shift(df_workload_region, execution_time)

    if df_workload_shift.empty:
        raise ValueError("No hay buyers disponibles en el shift actual para esta región.")

    # Separar URGENT=1 y URGENT=0
    urgent_mask = df["URGENT"] == 1
    df_urgent = df[urgent_mask].copy()
    df_non_urgent = df[~urgent_mask].copy()

    # Buyers habilitados para urgencias (Yes)
    df_workload_urgent = filter_buyers_by_urgency(df_workload_shift, urgent_required=False)
    buyers_urgent = df_workload_urgent["Buyer Alias"].dropna().tolist()

    # Si hay urgentes pero nadie habilitado, puedes:
    #  - O lanzar error y forzar corrección de datos
    #  - O degradar esos casos a no urgentes (se asignan con todos los buyers)
    if not df_urgent.empty and not buyers_urgent:
        raise ValueError(
            "Hay solicitudes URGENT=1 pero ningún buyer habilitado para urgencias en el shift actual."
        )

    # Asignar URGENT=1
    if not df_urgent.empty:
        df_urgent_assigned = round_robin_assign(
            tasks=df_urgent,
            buyers=buyers_urgent,
            buyer_column_name="BUYER",
        )
    else:
        df_urgent_assigned = df_urgent

    # Asignar URGENT=0 -> todos los buyers del shift
    buyers_all = df_workload_shift["Buyer Alias"].dropna().tolist()

    if not df_non_urgent.empty:
        df_non_urgent_assigned = round_robin_assign(
            tasks=df_non_urgent,
            buyers=buyers_all,
            buyer_column_name="BUYER",
        )
    else:
        df_non_urgent_assigned = df_non_urgent

    # Unir y ordenar
    df_assigned = pd.concat([df_urgent_assigned, df_non_urgent_assigned], ignore_index=True)

    if "ID" in df_assigned.columns:
        df_assigned = df_assigned.sort_values("ID").reset_index(drop=True)

    return df_assigned


# ------------------ MAIN ------------------ #

def main() -> None:
    ensure_output_dir("./data/final")

    now = parse_today_fixed()
    execution_time = now.time()
    print(f"Asignación ejecutada para fecha: {now.date()}, hora: {execution_time}")

    # Cargar workloads generados previamente
    df_workload_lam, df_workload_nam = load_workloads()

    # Cargar resultados base
    df_resultados = load_resultados("./data/Resultado.xlsx")

    # Separar PR especiales -V2/-V3
    df_resultados_normal, df_special = split_special_prs(df_resultados)
    special_path = "./data/final/ps_special.csv"
    df_special.to_csv(special_path, index=False, encoding="utf-8")
    print(f"Exportado especiales: {special_path}")

    df_resultados_normal.columns = [c.strip() for c in df_resultados_normal.columns]

    if "Assignment Group" not in df_resultados_normal.columns:
        raise KeyError("df_resultados debe contener la columna 'Assignment Group'.")

    # Separar por CAMERON NAM / CAMERON LAM
    group_col = df_resultados_normal["Assignment Group"].astype(str).str.strip().str.upper()
    mask_nam = group_col == "CAMERON NAM"
    mask_lam = group_col == "CAMERON LAM"

    df_nam = df_resultados_normal[mask_nam].copy()
    df_lam = df_resultados_normal[mask_lam].copy()

    # Asignar para NAM
    df_nam_assigned = assign_buyers_for_region(
        df_resultados_region=df_nam,
        df_workload_region=df_workload_nam,
        execution_time=execution_time,
    )

    # Asignar para LAM
    df_lam_assigned = assign_buyers_for_region(
        df_resultados_region=df_lam,
        df_workload_region=df_workload_lam,
        execution_time=execution_time,
    )

    # Exportar resultados
    nam_out = "./data/final/assignments_cam_ind_nam.csv"
    lam_out = "./data/final/assignments_cam_ind_lam.csv"

    df_nam_assigned.to_csv(nam_out, index=False, encoding="utf-8")
    df_lam_assigned.to_csv(lam_out, index=False, encoding="utf-8")

    print(f"Exportado assignments NAM: {nam_out}")
    print(f"Exportado assignments LAM: {lam_out}")


if __name__ == "__main__":
    main()
