"""
Lógica de asignación de monitores
"""
import pandas as pd
from .config import CONFIG


def esta_disponible(monitor, dia, hora_inicio, hora_fin):
    """Verifica si un monitor está disponible en el horario dado"""
    if dia not in monitor["disp"]:
        return False
    for r_inicio, r_fin in monitor["disp"][dia]:
        if hora_inicio >= r_inicio and hora_fin <= r_fin:
            return True
    return False


def verificar_restricciones(monitor, dia, hora_inicio, hora_fin):
    """Verifica restricciones de horas continuas"""
    cfg = CONFIG["asignacion"]
    if not cfg.get("max_horas_seguidas"):
        return True
    
    for asig in monitor["asignaciones"]:
        if asig["dia"] == dia:
            if (hora_inicio <= asig["fin"] and hora_fin >= asig["inicio"]):
                duracion_total = max(hora_fin, asig["fin"]) - min(hora_inicio, asig["inicio"])
                if duracion_total > cfg["max_horas_seguidas"]:
                    return False
    return True


def tiene_conflicto_horario(monitor, dia, hora_inicio, hora_fin):
    """Verifica si el monitor ya está asignado a otro grupo en ese horario"""
    for asig in monitor["asignaciones"]:
        if asig["dia"] == dia:
            # Verificar si hay solapamiento de horarios
            if not (hora_fin <= asig["inicio"] or hora_inicio >= asig["fin"]):
                return True  # Hay conflicto
    return False  # No hay conflicto


def asignar_monitores(monitores, df_espacios):
    """
    Asigna monitores a espacios según disponibilidad y configuración
    
    Args:
        monitores: Lista de monitores
        df_espacios: DataFrame con espacios a asignar
        
    Returns:
        Tuple (asignaciones, sin_monitor, monitores_actualizados)
    """
    cfg_asig = CONFIG["asignacion"]
    cfg_esp = CONFIG["espacios"]
    asignaciones = []
    sin_monitor = []
    espacios = df_espacios.to_dict('records')
    
    # FASE 1: Priorizar monitores que no alcanzan mínimo
    if cfg_asig.get("usar_prioridad") and cfg_asig.get("priorizar_minimo"):
        for espacio in espacios:
            dia = espacio['DIA_NORM']
            if pd.isna(dia):
                continue
            
            inicio = espacio[cfg_esp["col_hora_inicio"]]
            fin = espacio[cfg_esp["col_hora_fin"]]
            duracion = espacio['DURACION']
            
            candidatos = [
                m for m in monitores
                if m["horas"] < m["min"]
                and m["horas"] + duracion <= m["max"]
                and esta_disponible(m, dia, inicio, fin)
                and verificar_restricciones(m, dia, inicio, fin)
                and not tiene_conflicto_horario(m, dia, inicio, fin)  # ← NUEVO
            ]
            
            if candidatos:
                candidatos.sort(key=lambda x: (x["prioridad"], x["min"] - x["horas"]), reverse=True)
                elegido = candidatos[0]
                elegido["horas"] += duracion
                elegido["asignaciones"].append({
                    "dia": dia, 
                    "inicio": inicio, 
                    "fin": fin,
                    "sala": espacio.get(cfg_esp["col_sala"], ""),
                    "curso": espacio.get(cfg_esp["col_curso"], "")
                })
                asignaciones.append({
                    **espacio, 
                    "MONITOR": elegido["nombre"], 
                    "PRIORIDAD": elegido["prioridad"], 
                    "ESTADO": "✅"
                })
    
    # FASE 2: Asignar espacios restantes
    for espacio in espacios:
        dia = espacio['DIA_NORM']
        
        # Validar día
        if pd.isna(dia):
            sin_monitor.append(espacio)
            asignaciones.append({
                **espacio, 
                "MONITOR": "DÍA INVÁLIDO", 
                "PRIORIDAD": "-", 
                "ESTADO": "❌"
            })
            continue
        
        # Verificar si ya fue asignado
        ya_asignado = any(
            a.get(cfg_esp["col_sala"]) == espacio[cfg_esp["col_sala"]] and
            a.get('DIA_NORM') == espacio['DIA_NORM'] and
            a.get(cfg_esp["col_hora_inicio"]) == espacio[cfg_esp["col_hora_inicio"]] and
            a.get(cfg_esp["col_curso"]) == espacio.get(cfg_esp["col_curso"])  # ← Verificar curso también
            for a in asignaciones
        )
        if ya_asignado:
            continue
        
        inicio = espacio[cfg_esp["col_hora_inicio"]]
        fin = espacio[cfg_esp["col_hora_fin"]]
        duracion = espacio['DURACION']
        
        # Buscar candidatos
        candidatos = [
            m for m in monitores
            if m["horas"] + duracion <= m["max"]
            and esta_disponible(m, dia, inicio, fin)
            and verificar_restricciones(m, dia, inicio, fin)
            and not tiene_conflicto_horario(m, dia, inicio, fin)  # ← NUEVO: Evitar conflictos
        ]
        
        if not candidatos:
            sin_monitor.append(espacio)
            asignaciones.append({
                **espacio, 
                "MONITOR": "SIN MONITOR", 
                "PRIORIDAD": "-", 
                "ESTADO": "❌"
            })
            continue
        
        # Ordenar candidatos según configuración
        if cfg_asig.get("usar_prioridad"):
            candidatos.sort(key=lambda x: (x["prioridad"], -x["horas"]), reverse=True)
        elif cfg_asig.get("balancear_carga"):
            candidatos.sort(key=lambda x: x["horas"])
        
        elegido = candidatos[0]
        elegido["horas"] += duracion
        elegido["asignaciones"].append({
            "dia": dia, 
            "inicio": inicio, 
            "fin": fin,
            "sala": espacio.get(cfg_esp["col_sala"], ""),
            "curso": espacio.get(cfg_esp["col_curso"], "")
        })
        asignaciones.append({
            **espacio, 
            "MONITOR": elegido["nombre"], 
            "PRIORIDAD": elegido["prioridad"], 
            "ESTADO": "✅"
        })
    
    return asignaciones, sin_monitor, monitores