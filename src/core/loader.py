"""
Carga de datos desde archivos Excel/ODS
"""
import pandas as pd
from .config import CONFIG
from .parser import normalizar_dia, parse_range_cell, parse_prioridad


def cargar_monitores_desde_excel(ruta):
    """
    Carga monitores desde Excel (.xlsx, .xls) o LibreOffice (.ods)
    
    Args:
        ruta: Path al archivo
        
    Returns:
        Lista de diccionarios con datos de monitores
    """
    cfg = CONFIG["monitores"]
    
    # Detectar tipo de archivo y leer apropiadamente
    if ruta.endswith('.ods'):
        df_raw = pd.read_excel(ruta, sheet_name=0, header=None, engine='odf')
    else:
        df_raw = pd.read_excel(ruta, sheet_name=0, header=None)
    
    # Leer estructura del archivo
    dias_row = df_raw.iloc[3]
    jornadas_row = df_raw.iloc[cfg["header_row"]]
    col_mapping = {}
    current_dia = None
    col_nombre_idx = None
    col_prioridad_idx = None
    
    # Mapear columnas
    for idx, val in enumerate(jornadas_row):
        val_str = str(val).strip()
        if val_str == cfg["col_nombre"]:
            col_nombre_idx = idx
        if val_str == cfg["col_prioridad"]:
            col_prioridad_idx = idx
        
        dia_val = dias_row[idx] if idx < len(dias_row) else None
        if pd.notna(dia_val) and str(dia_val).strip():
            current_dia = normalizar_dia(dia_val)
        
        val_lower = val_str.lower()
        if val_lower in ['mañana', 'manana', 'tarde', 'noche']:
            if current_dia:
                key = f"{current_dia}_{val_lower}"
                col_mapping[key] = idx
    
    if col_nombre_idx is None:
        raise ValueError(f"No se encuentra la columna '{cfg['col_nombre']}'")
    
    # Extraer monitores
    monitores = []
    for row_idx in range(cfg["data_start_row"], len(df_raw)):
        row = df_raw.iloc[row_idx]
        nombre = row[col_nombre_idx]
        
        if pd.isna(nombre) or str(nombre).strip() == "":
            continue
        
        prioridad = 3
        if col_prioridad_idx is not None:
            prioridad = parse_prioridad(row[col_prioridad_idx])
        
        mon = {
            "id": row_idx - cfg["data_start_row"],
            "nombre": str(nombre).strip(),
            "prioridad": prioridad,
            "min": cfg["horas_min_default"],
            "max": cfg["horas_max_default"],
            "horas": 0,
            "disp": {},
            "asignaciones": []
        }
        
        # Extraer disponibilidad
        dias_unicos = set(k.split('_')[0] for k in col_mapping.keys())
        for dia in dias_unicos:
            mon["disp"][dia] = []
            for jornada in ['mañana', 'manana', 'tarde', 'noche']:
                key = f"{dia}_{jornada}"
                if key in col_mapping:
                    col_idx = col_mapping[key]
                    ranges = parse_range_cell(row[col_idx])
                    mon["disp"][dia].extend(ranges)
        
        monitores.append(mon)
    
    return monitores


def cargar_espacios_desde_excel(ruta):
    """
    Carga espacios desde Excel (.xlsx, .xls) o LibreOffice (.ods)
    
    Args:
        ruta: Path al archivo
        
    Returns:
        DataFrame con datos de espacios
    """
    cfg = CONFIG["espacios"]
    
    # Detectar tipo de archivo
    if ruta.endswith('.ods'):
        df = pd.read_excel(ruta, sheet_name=0, engine='odf')
    else:
        df = pd.read_excel(ruta, sheet_name=0)
    
    # Validar columnas requeridas
    columnas_req = [cfg["col_sala"], cfg["col_dia"], cfg["col_hora_inicio"], 
                    cfg["col_hora_fin"], cfg["col_curso"]]
    faltantes = [col for col in columnas_req if col not in df.columns]
    if faltantes:
        raise ValueError(f"Columnas no encontradas: {faltantes}")
    
    # Normalizar y calcular duración
    df['DIA_NORM'] = df[cfg["col_dia"]].apply(normalizar_dia)
    df['DURACION'] = df[cfg["col_hora_fin"]] - df[cfg["col_hora_inicio"]]
    
    return df