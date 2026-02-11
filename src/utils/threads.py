"""
Threads para operaciones en segundo plano
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import pandas as pd
from PySide6.QtCore import QThread, Signal

from core.assignment import asignar_monitores
from core.config import CONFIG


class AsignacionThread(QThread):
    """Thread para ejecutar la asignación sin bloquear la UI"""
    finished = Signal(pd.DataFrame, list, str)
    error = Signal(str)
    progress = Signal(str)
    
    def __init__(self, monitores, df_espacios):
        super().__init__()
        self.monitores = monitores
        self.df_espacios = df_espacios
    
    def run(self):
        try:
            modo = "con prioridad" if CONFIG["asignacion"]["usar_prioridad"] else "sin prioridad"
            self.progress.emit(f"🔄 Iniciando asignación {modo}...")
            
            asignaciones, sin_monitor, monitores = asignar_monitores(
                self.monitores, 
                self.df_espacios
            )
            
            df_result = pd.DataFrame(asignaciones)
            exitosos = len([a for a in asignaciones if a["ESTADO"] == "✅"])
            total = len(asignaciones)
            
            modo_titulo = "CON PRIORIDAD" if CONFIG["asignacion"]["usar_prioridad"] else "SIN PRIORIDAD"
            reporte = f"📊 REPORTE DE ASIGNACIÓN {modo_titulo}\n{'='*60}\n"
            reporte += f"\n🎯 Resumen:\n   Total horarios: {total}\n"
            reporte += f"   Asignados: {exitosos} ({exitosos*100/total:.1f}%)\n"
            reporte += f"   Sin monitor: {len(sin_monitor)} ({len(sin_monitor)*100/total:.1f}%)\n"
            
            self.progress.emit("✅ Asignación completada")
            self.finished.emit(df_result, monitores, reporte)
            
        except Exception as e:
            self.error.emit(str(e))