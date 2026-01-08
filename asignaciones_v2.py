#!/usr/bin/env python3
"""
asignaciones.py - Versión con Weighted Round-Robin
Considera 'Workload / Availability' para distribución proporcional.
"""

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
    Usa la columna ID.
    """
    df = df.copy()
    df["ID"] = df["ID"].astype(str).str.strip()
    mask_special = df["ID"].str.contains(r"-V", case=False, na=False)
    df_special = df[mask_special].copy()
    df_normal = df[~mask_special].copy()
    return df_normal, df_special

# ------------------ SHIFTS Y URGENCIAS ------------------

def parse_shift_to_range(shift_str: str) -> Tuple[time, time]:
    """
    Convierte un string tipo '7 to 5', '9 to 7', '6 to 4' en (hora_inicio, hora_fin).
    Se interpreta como turno dentro del mismo día:
    '6 to 4' -> 06:00–16:00
    '7 to 5' -> 07:00–17:00
    '9 to 7' -> 09:00–19:00
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

# ------------------ WEIGHTED ROUND ROBIN (NUEVO) ------------------

def extract_capacity_weight(workload_str: str, default: float = 1.0) -> float:
    """
    Extrae peso desde 'Workload / Availability'.
    Ejemplos:
    - "100%" -> 1.0
    - "50%"  -> 0.5
    - "25%"  -> 0.25
    - "0%"   -> 0.0 (excluido)
    - ""     -> 1.0 (default)
    """
    if not isinstance(workload_str, str):
        return default
    
    workload_str = workload_str.strip()
    match = re.search(r'(\d+)%', workload_str)
    
    if match:
        percentage = int(match.group(1))
        return percentage / 100.0
    
    return default


def weighted_round_robin_assign(
    tasks: pd.DataFrame,
    buyers_df: pd.DataFrame,
    buyer_column_name: str = "BUYER",
) -> pd.DataFrame:
    """
    Asigna buyers proporcionalmente según 'Workload / Availability'.
    
    NUEVO ALGORITMO:
    1. Calcular PRs por buyer basado en peso relativo
    2. Asignar exactamente esa cantidad a cada buyer
    3. Distribuir residuo equitativamente
    """
    if buyers_df.empty:
        raise ValueError("No hay buyers disponibles para asignar.")
    
    tasks = tasks.copy()
    
    # Asegurar columna BUYER
    if buyer_column_name not in tasks.columns:
        tasks[buyer_column_name] = pd.Series([None] * len(tasks), dtype="object")
    else:
        tasks[buyer_column_name] = tasks[buyer_column_name].astype("object")
    
    # Extraer capacidades
    buyers_df = buyers_df.copy()
    buyers_df['capacity_weight'] = buyers_df['Workload / Availability'].apply(
        lambda x: extract_capacity_weight(x, default=1.0)
    )
    
    # Filtrar buyers con capacidad > 0
    buyers_df = buyers_df[buyers_df['capacity_weight'] > 0].copy()
    
    if buyers_df.empty:
        raise ValueError("Todos los buyers tienen capacidad 0%.")
    
    # Calcular distribución proporcional EXACTA
    total_prs = len(tasks)
    total_weight = buyers_df['capacity_weight'].sum()
    
    buyers_df['prs_allocated'] = (
        (buyers_df['capacity_weight'] / total_weight) * total_prs
    ).round().astype(int)
    
    # Ajustar residuo (por redondeo)
    diff = total_prs - buyers_df['prs_allocated'].sum()
    
    if diff > 0:
        # Asignar PRs extra a buyers con mayor peso
        buyers_sorted = buyers_df.sort_values('capacity_weight', ascending=False)
        for i in range(diff):
            idx = buyers_sorted.index[i % len(buyers_sorted)]
            buyers_df.loc[idx, 'prs_allocated'] += 1
    elif diff < 0:
        # Quitar PRs de buyers con menor peso
        buyers_sorted = buyers_df.sort_values('capacity_weight', ascending=True)
        for i in range(abs(diff)):
            idx = buyers_sorted.index[i % len(buyers_sorted)]
            buyers_df.loc[idx, 'prs_allocated'] = max(0, buyers_df.loc[idx, 'prs_allocated'] - 1)
    
    # Asignar PRs secuencialmente
    pr_index = 0
    for _, buyer_row in buyers_df.iterrows():
        buyer_alias = buyer_row['Buyer Alias']
        num_prs = int(buyer_row['prs_allocated'])
        
        for _ in range(num_prs):
            if pr_index < len(tasks):
                tasks.iat[pr_index, tasks.columns.get_loc(buyer_column_name)] = buyer_alias
                pr_index += 1
    
    return tasks



# ------------------ LÓGICA PRINCIPAL DE ASIGNACIÓN ------------------

def assign_buyers_for_region(
    df_resultados_region: pd.DataFrame,
    df_workload_region: pd.DataFrame,
    execution_time: time,
    df_buyers_full: pd.DataFrame,  # ← NUEVO: CSV completo
) -> pd.DataFrame:
    """
    Asigna buyers para una región (NAM o LAM):
    - Separa URGENT=1 y URGENT=0.
    - URGENT=1 -> solo buyers con Available For Urgencies == 'Yes'.
    - Respeta shift: solo buyers activos en la hora de ejecución.
    - Reparte proporcionalmente según 'Workload / Availability' (weighted round-robin).
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
    
    # Merge con df_buyers_full para obtener "Workload / Availability"
    buyers_aliases = df_workload_shift["Buyer Alias"].tolist()
    df_buyers_region = df_buyers_full[
        df_buyers_full["Buyer Alias"].isin(buyers_aliases)
    ].copy()
    
    # Separar URGENT=1 y URGENT=0
    urgent_mask = df["URGENT"] == 1
    df_urgent = df[urgent_mask].copy()
    df_non_urgent = df[~urgent_mask].copy()
    
    # Buyers habilitados para urgencias (Yes)
    df_buyers_urgent = df_buyers_region[
        df_buyers_region["Available For Urgencies"].str.strip().str.upper() == "YES"
    ].copy()
    
    # Si hay urgentes pero nadie habilitado
    if not df_urgent.empty and df_buyers_urgent.empty:
        raise ValueError(
            "Hay solicitudes URGENT=1 pero ningún buyer habilitado para urgencias en el shift actual."
        )
    
    # Asignar URGENT=1 (CON PESOS)
    if not df_urgent.empty:
        df_urgent_assigned = weighted_round_robin_assign(
            tasks=df_urgent,
            buyers_df=df_buyers_urgent,
            buyer_column_name="BUYER",
        )
    else:
        df_urgent_assigned = df_urgent
    
    # Asignar URGENT=0 (CON PESOS) -> todos los buyers del shift
    if not df_non_urgent.empty:
        df_non_urgent_assigned = weighted_round_robin_assign(
            tasks=df_non_urgent,
            buyers_df=df_buyers_region,
            buyer_column_name="BUYER",
        )
    else:
        df_non_urgent_assigned = df_non_urgent
    
    # Unir y ordenar
    df_assigned = pd.concat([df_urgent_assigned, df_non_urgent_assigned], ignore_index=True)
    
    if "ID" in df_assigned.columns:
        df_assigned = df_assigned.sort_values("ID").reset_index(drop=True)
    
    return df_assigned

# ------------------ MAIN ------------------

def main() -> None:
    ensure_output_dir("./data/final")
    
    now = parse_today_fixed()
    execution_time = now.time()
    
    print(f"Asignación ejecutada para fecha: {now.date()}, hora: {execution_time}")
    
    # Cargar workloads generados previamente
    df_workload_lam, df_workload_nam = load_workloads()
    
    # ← NUEVO: Cargar CSV completo de buyers
    df_buyers_full = pd.read_csv("./data/sharepoint/sap_buyers.csv", encoding='utf-8-sig')
    df_buyers_full.columns = [c.strip() for c in df_buyers_full.columns]
    
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
    
    # Asignar para NAM (CON PESOS)
    df_nam_assigned = assign_buyers_for_region(
        df_resultados_region=df_nam,
        df_workload_region=df_workload_nam,
        execution_time=execution_time,
        df_buyers_full=df_buyers_full,
    )
    
    # Asignar para LAM (CON PESOS)
    df_lam_assigned = assign_buyers_for_region(
        df_resultados_region=df_lam,
        df_workload_region=df_workload_lam,
        execution_time=execution_time,
        df_buyers_full=df_buyers_full,
    )
    
    # Exportar resultados
    nam_out = "./data/final/assignments_cam_ind_nam.csv"
    lam_out = "./data/final/assignments_cam_ind_lam.csv"
    
    df_nam_assigned.to_csv(nam_out, index=False, encoding="utf-8")
    df_lam_assigned.to_csv(lam_out, index=False, encoding="utf-8")
    
    print(f"✅ Exportado assignments NAM: {nam_out}")
    print(f"✅ Exportado assignments LAM: {lam_out}")
    
    # Validar distribución
    print("\n=== DISTRIBUCIÓN NAM ===")
    if not df_nam_assigned.empty:
        print(df_nam_assigned['BUYER'].value_counts().sort_index())
    else:
        print("(Sin PRs para NAM)")
    
    print("\n=== DISTRIBUCIÓN LAM ===")
    if not df_lam_assigned.empty:
        print(df_lam_assigned['BUYER'].value_counts().sort_index())
    else:
        print("(Sin PRs para LAM)")
    
    # Verificar que no hay PRs duplicados
    print("\n=== VALIDACIÓN UNICIDAD ===")
    total_assigned = len(df_nam_assigned) + len(df_lam_assigned)
    total_original = len(df_nam) + len(df_lam)
    print(f"PRs originales: {total_original}")
    print(f"PRs asignados: {total_assigned}")
    print(f"Duplicados: {total_assigned - total_original} (debe ser 0)")

if __name__ == "__main__":
    main()
