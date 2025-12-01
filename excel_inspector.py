import pandas as pd
import re
import os

# ========================================================
# CONFIGURACIÓN
# ========================================================

CONFIG = {
    "monitores": {
        "archivo": "DISPONIBILIDAD HORARIA MONITORES DE SALAS 2025-II.xlsx",
        "hoja": "Hoja 1",
        "header_row": 4,  # Fila donde están los encabezados
        "data_start_row": 5,  # Fila donde empiezan los datos
        
        # Columnas base
        "col_nombre": "Nombre completo",
        "col_cedula": "Cédula",
        "col_codigo": "Código",
        
        # Horas mínimas y máximas (ajustar si existen en tu Excel)
        "horas_min_default": 8,
        "horas_max_default": 20
    },
    
    "cursos": {
        "archivo": "2025-II-Espacios.xlsx",
        "hoja": "espaciosSalonesAud",
        "header_row": 2,  # Fila donde están los días
        "data_start_row": 40,  # Fila donde empiezan los horarios de salas
        "data_end_row": 56  # Fila donde terminan los horarios
    }
}

# ========================================================
# FUNCIONES DE PARSEO DE HORARIOS
# ========================================================

def parse_time_str(time_str):
    """Convierte '7:00am' -> 7, '2:00pm' -> 14"""
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
    
    # Buscar solo números
    match = re.search(r'\d+', s)
    if match:
        return int(match.group(0))
    
    return None


def parse_range_cell(cell_value):
    """Convierte '7:00am-1:00pm' -> [(7, 13)]"""
    if pd.isna(cell_value):
        return []
    
    s = str(cell_value).strip().lower()
    
    if s in ["libre", "disponible", "todo el día"]:
        return [(7, 22)]
    
    if s in ["no disponible", "no", "n/a", "", "nan"]:
        return []
    
    ranges = []
    pattern = r'(\d{1,2}(?::\d{2})?\s*(?:am|pm)?)\s*-\s*(\d{1,2}(?::\d{2})?\s*(?:am|pm)?)'
    
    matches = re.findall(pattern, s)
    for match in matches:
        start = parse_time_str(match[0])
        end = parse_time_str(match[1])
        
        if start is not None and end is not None:
            ranges.append((start, end))
    
    return ranges


# ========================================================
# CARGA DE MONITORES
# ========================================================

def cargar_monitores():
    """Lee el Excel de monitores con estructura multi-nivel"""
    cfg = CONFIG["monitores"]
    
    print(f"\n📂 Cargando monitores desde: {cfg['archivo']}")
    
    # Leer archivo completo sin header
    df_raw = pd.read_excel(cfg["archivo"], sheet_name=cfg["hoja"], header=None)
    
    # Extraer fila de días (fila 3)
    dias_row = df_raw.iloc[3]
    
    # Extraer fila de jornadas (fila 4)
    jornadas_row = df_raw.iloc[cfg["header_row"]]
    
    # Construir mapeo de columnas
    col_mapping = {}
    current_dia = None
    
    for idx, val in enumerate(jornadas_row):
        val_str = str(val).strip().lower()
        
        # Detectar nombre de día en fila superior
        dia_val = dias_row[idx] if idx < len(dias_row) else None
        if pd.notna(dia_val) and str(dia_val).strip():
            current_dia = str(dia_val).strip().lower()
            # Normalizar acentos
            current_dia = current_dia.replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
        
        # Detectar jornada
        if val_str in ['mañana', 'manana', 'tarde', 'noche']:
            if current_dia:
                key = f"{current_dia}_{val_str}"
                col_mapping[key] = idx
    
    print(f"✓ Estructura detectada:")
    for key in sorted(col_mapping.keys()):
        print(f"   {key} -> Columna {col_mapping[key]}")
    
    # Leer datos de monitores
    monitores = []
    
    for row_idx in range(cfg["data_start_row"], len(df_raw)):
        row = df_raw.iloc[row_idx]
        
        nombre = row[jornadas_row.tolist().index(cfg["col_nombre"])]
        if pd.isna(nombre) or str(nombre).strip() == "":
            continue
        
        mon = {
            "id": row_idx - cfg["data_start_row"],
            "nombre": str(nombre).strip(),
            "min": cfg["horas_min_default"],
            "max": cfg["horas_max_default"],
            "horas": 0,
            "disp": {},
            "asignaciones": []
        }
        
        # Organizar disponibilidad por día
        dias_unicos = set(k.split('_')[0] for k in col_mapping.keys())
        
        for dia in dias_unicos:
            mon["disp"][dia] = []
            
            # Recopilar rangos de todas las jornadas del día
            for jornada in ['mañana', 'manana', 'tarde', 'noche']:
                key = f"{dia}_{jornada}"
                if key in col_mapping:
                    col_idx = col_mapping[key]
                    cell_value = row[col_idx]
                    ranges = parse_range_cell(cell_value)
                    mon["disp"][dia].extend(ranges)
        
        monitores.append(mon)
    
    print(f"✅ Cargados {len(monitores)} monitores")
    
    # Mostrar ejemplo
    if monitores:
        print(f"\n📋 Ejemplo - {monitores[0]['nombre']}:")
        for dia, rangos in monitores[0]['disp'].items():
            if rangos:
                print(f"   {dia.capitalize()}: {rangos}")
    
    return monitores


# ========================================================
# CARGA DE CURSOS (desde matriz de horarios)
# ========================================================

def detectar_bloques_salas(df_raw, fila_titulos=1):
    """
    Detecta los bloques de columnas donde está cada sala.
    Cada sala tiene un título en la fila 1 (ej: "SALA 1 (40)")
    """
    fila_salas = df_raw.iloc[fila_titulos]
    
    bloques = []
    sala_actual = None
    inicio_col = None
    
    for col_idx, val in enumerate(fila_salas):
        if pd.notna(val) and str(val).strip():
            # Detectar nombre de sala
            texto = str(val).strip()
            match = re.search(r'(SALA|AUDITORIO|LAB|LABORATORIO)\s*\d+', texto, re.IGNORECASE)
            
            if match:
                # Si ya había una sala anterior, guardarla
                if sala_actual and inicio_col is not None:
                    bloques.append({
                        'nombre': sala_actual,
                        'col_inicio': inicio_col,
                        'col_fin': col_idx - 1
                    })
                
                # Nueva sala detectada
                sala_actual = match.group(0)
                inicio_col = col_idx
    
    # Agregar la última sala
    if sala_actual and inicio_col is not None:
        bloques.append({
            'nombre': sala_actual,
            'col_inicio': inicio_col,
            'col_fin': len(fila_salas) - 1
        })
    
    return bloques


def cargar_cursos_de_bloque(df_raw, sala_info, fila_dias=2):
    """
    Lee los cursos de un bloque específico de columnas (una sala).
    
    Args:
        df_raw: DataFrame completo
        sala_info: dict con 'nombre', 'col_inicio', 'col_fin'
        fila_dias: fila donde están los nombres de días
    """
    cfg = CONFIG["cursos"]
    sala_nombre = sala_info['nombre']
    col_inicio = sala_info['col_inicio']
    col_fin = sala_info['col_fin']
    
    # Extraer días del bloque de esta sala
    dias_row = df_raw.iloc[fila_dias, col_inicio:col_fin+1]
    
    dia_mapping = {}
    for idx_local, val in enumerate(dias_row):
        col_idx_global = col_inicio + idx_local
        if pd.notna(val):
            dia = str(val).strip().lower()
            # Normalizar
            dia = dia.replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
            # Limpiar
            dia = re.sub(r'[^a-z]', '', dia)
            if dia and len(dia) >= 3:  # Al menos 3 letras
                dia_mapping[col_idx_global] = dia
    
    # Extraer cursos
    cursos = []
    
    for row_idx in range(cfg["data_start_row"], cfg["data_end_row"] + 1):
        if row_idx >= len(df_raw):
            break
            
        row = df_raw.iloc[row_idx]
        
        # Extraer hora de la columna 0 (puede estar en cualquier formato)
        hora_str = str(row[0]).strip() if pd.notna(row[0]) else ""
        
        # Intentar múltiples formatos de hora
        inicio = fin = None
        
        # Formato "7 - 8" o "7-8"
        match = re.search(r'(\d+)\s*-\s*(\d+)', hora_str)
        if match:
            inicio = int(match.group(1))
            fin = int(match.group(2))
        else:
            # Formato "7:00" (asumir 1 hora de duración)
            match = re.search(r'(\d+)', hora_str)
            if match:
                inicio = int(match.group(1))
                fin = inicio + 1
        
        if inicio is None or fin is None:
            continue
        
        # Revisar cada día en este bloque
        for col_idx, dia in dia_mapping.items():
            cell_value = row[col_idx]
            
            if pd.notna(cell_value) and str(cell_value).strip():
                curso_texto = str(cell_value).strip()
                curso_texto = re.sub(r'\s+', ' ', curso_texto)
                
                # Filtrar valores no deseados
                if len(curso_texto) > 3 and curso_texto.lower() not in ['nan', 'none']:
                    cursos.append({
                        "curso": curso_texto,
                        "sala": sala_nombre,
                        "dia": dia,
                        "inicio": inicio,
                        "fin": fin
                    })
    
    return cursos


def cargar_cursos():
    """Lee TODAS las salas del Excel (organizadas horizontalmente)"""
    cfg = CONFIG["cursos"]
    
    print(f"\n📂 Cargando cursos desde: {cfg['archivo']}")
    print(f"   Hoja: {cfg['hoja']}")
    print(f"   Filas de datos: {cfg['data_start_row']} a {cfg['data_end_row']}")
    
    # Leer archivo completo
    df_raw = pd.read_excel(cfg["archivo"], sheet_name=cfg["hoja"], header=None)
    
    # Detectar bloques de salas
    bloques_salas = detectar_bloques_salas(df_raw, fila_titulos=1)
    
    print(f"\n✓ Salas detectadas: {len(bloques_salas)}")
    for bloque in bloques_salas:
        print(f"   {bloque['nombre']:20} | Columnas {bloque['col_inicio']:2d}-{bloque['col_fin']:2d}")
    
    # Cargar cursos de cada sala
    todos_cursos = []
    
    for bloque in bloques_salas:
        try:
            cursos_sala = cargar_cursos_de_bloque(df_raw, bloque, fila_dias=cfg["header_row"])
            todos_cursos.extend(cursos_sala)
            print(f"   ✅ {bloque['nombre']:20} | {len(cursos_sala):3d} horarios")
        except Exception as e:
            print(f"   ⚠️ Error en {bloque['nombre']}: {e}")
    
    print(f"\n✅ Total: {len(todos_cursos)} horarios cargados")
    
    if todos_cursos:
        print(f"\n📋 Ejemplo:")
        print(f"   {todos_cursos[0]['curso'][:50]}")
        print(f"   {todos_cursos[0]['sala']} - {todos_cursos[0]['dia'].capitalize()} {todos_cursos[0]['inicio']}:00-{todos_cursos[0]['fin']}:00")
    
    return todos_cursos


# ========================================================
# LÓGICA DE ASIGNACIÓN
# ========================================================

def esta_disponible(mon, dia, inicio, fin):
    """Verifica si el monitor está disponible"""
    if dia not in mon["disp"]:
        return False
    
    for r_inicio, r_fin in mon["disp"][dia]:
        if inicio >= r_inicio and fin <= r_fin:
            return True
    
    return False


def asignar_monitores(monitores, cursos):
    """Asigna monitores a cursos según disponibilidad"""
    asignaciones = []
    sin_monitor = []
    
    for c in cursos:
        inicio = c["inicio"]
        fin = c["fin"]
        horas = fin - inicio
        
        # Buscar monitores disponibles
        candidatos = [
            m for m in monitores
            if m["horas"] + horas <= m["max"] and esta_disponible(m, c["dia"], inicio, fin)
        ]
        
        if not candidatos:
            sin_monitor.append(c)
            asignaciones.append({**c, "monitor": "SIN MONITOR", "estado": "❌"})
            continue
        
        # Asignar al monitor con menos carga
        candidatos.sort(key=lambda x: x["horas"])
        elegido = candidatos[0]
        
        elegido["horas"] += horas
        elegido["asignaciones"].append(c)
        
        asignaciones.append({
            **c,
            "monitor": elegido["nombre"],
            "estado": "✅",
            "horas": horas
        })
    
    return asignaciones, sin_monitor


# ========================================================
# REPORTES
# ========================================================

def generar_reporte(monitores, asignaciones, sin_monitor):
    """Genera estadísticas de asignación"""
    print("\n" + "="*70)
    print("📊 REPORTE DE ASIGNACIÓN")
    print("="*70)
    
    exitosos = len([a for a in asignaciones if a["estado"] == "✅"])
    total = len(asignaciones)
    
    print(f"\n🎯 Cursos asignados: {exitosos}/{total} ({exitosos*100/total:.1f}%)")
    print(f"❌ Sin monitor: {len(sin_monitor)}")
    
    # Estadísticas por sala
    salas = {}
    for a in asignaciones:
        sala = a["sala"]
        if sala not in salas:
            salas[sala] = {"total": 0, "asignados": 0}
        salas[sala]["total"] += 1
        if a["estado"] == "✅":
            salas[sala]["asignados"] += 1
    
    print("\n🏢 RESUMEN POR SALA:")
    print("-" * 70)
    for sala, stats in sorted(salas.items()):
        porcentaje = (stats["asignados"] / stats["total"]) * 100 if stats["total"] > 0 else 0
        print(f"{sala:30} | {stats['asignados']:3}/{stats['total']:3} ({porcentaje:5.1f}%)")
    
    print("\n👥 CARGA DE TRABAJO:")
    print("-" * 70)
    
    monitores_activos = [m for m in monitores if m["horas"] > 0]
    for m in sorted(monitores_activos, key=lambda x: x["horas"], reverse=True):
        porcentaje = (m["horas"] / m["max"]) * 100
        barra = "█" * int(porcentaje / 5)
        
        status = "✅"
        if m["horas"] < m["min"]:
            status = f"⚠️ Bajo mínimo ({m['min']}h)"
        elif m["horas"] > m["max"]:
            status = f"❌ Excede máximo ({m['max']}h)"
        
        print(f"{m['nombre'][:35]:35} | {m['horas']:2}h {barra:20} {status}")
    
    # Monitores sin asignar
    monitores_sin_carga = [m for m in monitores if m["horas"] == 0]
    if monitores_sin_carga:
        print(f"\n⚠️ MONITORES SIN ASIGNACIONES ({len(monitores_sin_carga)}):")
        for m in monitores_sin_carga[:10]:
            print(f"  • {m['nombre']}")
    
    if sin_monitor:
        print(f"\n❌ CURSOS SIN MONITOR ({len(sin_monitor)}):")
        print("-" * 70)
        # Agrupar por sala
        sin_monitor_por_sala = {}
        for c in sin_monitor:
            sala = c["sala"]
            if sala not in sin_monitor_por_sala:
                sin_monitor_por_sala[sala] = []
            sin_monitor_por_sala[sala].append(c)
        
        for sala, cursos in sorted(sin_monitor_por_sala.items()):
            print(f"\n  {sala} ({len(cursos)}):")
            for c in cursos[:5]:
                print(f"    • {c['curso'][:45]:45} | {c['dia'].capitalize()} {c['inicio']}-{c['fin']}")
            if len(cursos) > 5:
                print(f"    ... y {len(cursos)-5} más")


def exportar_resultados(asignaciones, monitores):
    """Exporta resultados a Excel"""
    df_asignaciones = pd.DataFrame(asignaciones)
    
    # Ordenar por sala, día y hora
    df_asignaciones = df_asignaciones.sort_values(['sala', 'dia', 'inicio'])
    
    resumen = []
    for m in monitores:
        resumen.append({
            "Monitor": m["nombre"],
            "Horas": m["horas"],
            "Mínimo": m["min"],
            "Máximo": m["max"],
            "Cursos": len(m["asignaciones"]),
            "Estado": "✅" if m["min"] <= m["horas"] <= m["max"] else "⚠️"
        })
    df_resumen = pd.DataFrame(resumen)
    df_resumen = df_resumen.sort_values('Horas', ascending=False)
    
    # Resumen por sala
    salas_stats = []
    for sala in df_asignaciones['sala'].unique():
        sala_df = df_asignaciones[df_asignaciones['sala'] == sala]
        asignados = len(sala_df[sala_df['estado'] == '✅'])
        total = len(sala_df)
        salas_stats.append({
            'Sala': sala,
            'Total Horarios': total,
            'Con Monitor': asignados,
            'Sin Monitor': total - asignados,
            '% Cobertura': f"{(asignados/total*100):.1f}%"
        })
    df_salas = pd.DataFrame(salas_stats)
    
    archivo = "asignacion_monitores_resultado.xlsx"
    with pd.ExcelWriter(archivo, engine='openpyxl') as writer:
        df_asignaciones.to_excel(writer, sheet_name='Asignaciones', index=False)
        df_resumen.to_excel(writer, sheet_name='Resumen Monitores', index=False)
        df_salas.to_excel(writer, sheet_name='Resumen Salas', index=False)
    
    print(f"\n✅ Archivo generado: {archivo}")
    print(f"   📄 Hoja 1: Asignaciones completas")
    print(f"   📄 Hoja 2: Resumen de monitores")
    print(f"   📄 Hoja 3: Resumen por sala")


# ========================================================
# MAIN
# ========================================================

if __name__ == "__main__":
    print("🚀 Sistema de Asignación de Monitores")
    print("="*70)
    
    # Verificar archivos
    print(f"\n📁 Directorio: {os.getcwd()}")
    
    monitores = cargar_monitores()
    cursos = cargar_cursos()
    
    if not monitores or not cursos:
        print("\n❌ Error al cargar datos")
        exit(1)
    
    print("\n🔄 Procesando asignaciones...")
    asignaciones, sin_monitor = asignar_monitores(monitores, cursos)
    
    generar_reporte(monitores, asignaciones, sin_monitor)
    exportar_resultados(asignaciones, monitores)
    
    print("\n✨ ¡Proceso completado!")