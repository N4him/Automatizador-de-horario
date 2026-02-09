"""
Configuración central del sistema
"""
import json
import os
from pathlib import Path

# Ruta del archivo de configuración
CONFIG_FILE = Path.home() / ".asignacion_monitores" / "config.json"

# Configuración por defecto
DEFAULT_CONFIG = {
    "monitores": {
        "header_row": 4,
        "data_start_row": 5,
        "col_nombre": "Nombre completo",
        "col_prioridad": "Prioridad",
        "col_min": None,
        "col_max": None,
        "horas_min_default": 8,
        "horas_max_default": 20
    },
    "espacios": {
        "col_sala": "SALA",
        "col_dia": "DIA",
        "col_hora_inicio": "HORA_INICIO",
        "col_hora_fin": "HORA_FIN",
        "col_curso": "CURSO"
    },
    "asignacion": {
        "balancear_carga": True,
        "priorizar_minimo": True,
        "usar_prioridad": True,
        "max_horas_seguidas": 4,
        "descanso_minimo": 1,
        "permitir_sobrepasar_max": False
    }
}

def load_config():
    """Carga la configuración desde archivo JSON"""
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                loaded_config = json.load(f)
                # Merge con valores por defecto para agregar nuevos campos
                merged_config = DEFAULT_CONFIG.copy()
                for key in loaded_config:
                    if key in merged_config:
                        merged_config[key].update(loaded_config[key])
                return merged_config
        except Exception as e:
            print(f"Error al cargar configuración: {e}")
            return DEFAULT_CONFIG.copy()
    return DEFAULT_CONFIG.copy()

def save_config(config):
    """Guarda la configuración en archivo JSON"""
    try:
        # Crear directorio si no existe
        CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
        
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"Error al guardar configuración: {e}")
        return False

# Cargar configuración al importar el módulo
CONFIG = load_config()