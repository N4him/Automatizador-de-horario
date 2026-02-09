"""
Funciones de parsing y normalización
"""
import re
import pandas as pd


def parse_time_str(time_str):
    """Convierte string de hora a entero (formato 24h)"""
    if pd.isna(time_str):
        return None
    s = str(time_str).strip().lower()
    match = re.search(r'(\d{1,2})(?::(\d{2}))?\s*(am|pm)', s)
    if match:
        hour = int(match.group(1))
        meridiem = match.group(3)
        if meridiem == 'pm' and hour != 12:
            hour += 12
        elif meridiem == 'am' and hour == 12:
            hour = 0
        return hour
    match = re.search(r'\d+', s)
    if match:
        return int(match.group(0))
    return None


def parse_range_cell(cell_value):
    """Extrae rangos horarios de una celda"""
    if pd.isna(cell_value):
        return []
    s = str(cell_value).strip().lower()
    
    # Casos especiales
    if s in ["libre", "disponible", "todo el día", "todo el dia"]:
        return [(7, 22)]
    if s in ["no disponible", "no", "n/a", "", "nan"]:
        return []
    
    # Extrae rangos con regex
    ranges = []
    pattern = r'(\d{1,2}(?::\d{2})?\s*(?:am|pm)?)\s*-\s*(\d{1,2}(?::\d{2})?\s*(?:am|pm)?)'
    matches = re.findall(pattern, s)
    for match in matches:
        start = parse_time_str(match[0])
        end = parse_time_str(match[1])
        if start is not None and end is not None:
            ranges.append((start, end))
    return ranges


def normalizar_dia(dia):
    """Normaliza nombres de días de la semana"""
    if pd.isna(dia):
        return None
    d = str(dia).strip().lower()
    
    # Eliminar acentos
    d = d.replace('á', 'a').replace('é', 'e').replace('í', 'i')
    d = d.replace('ó', 'o').replace('ú', 'u')
    d = re.sub(r'[^a-z]', '', d)
    
    # Mapeo de abreviaturas
    mapeo = {
        'lun': 'lunes', 'mar': 'martes', 'mie': 'miercoles',
        'jue': 'jueves', 'vie': 'viernes', 'sab': 'sabado', 'dom': 'domingo'
    }
    
    for abrev, completo in mapeo.items():
        if d.startswith(abrev):
            return completo
    
    if d in mapeo.values():
        return d
    
    return d if len(d) >= 3 else None


def parse_prioridad(valor):
    """Convierte valor de prioridad a entero 1-5"""
    if pd.isna(valor):
        return 3
    try:
        prioridad = int(valor)
        if 1 <= prioridad <= 5:
            return prioridad
        return 3
    except:
        return 3