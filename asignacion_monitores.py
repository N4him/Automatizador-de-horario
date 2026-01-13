import sys
import pandas as pd
import re
from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, 
    QHBoxLayout, QTableView, QFileDialog, QLabel, QMessageBox,
    QProgressBar, QTextEdit, QCheckBox, QFrame, QDialog, QSizePolicy, QTabWidget
)
from PySide6.QtCore import Qt, QAbstractTableModel, QThread, Signal
from PySide6.QtGui import QFont


# ========================================================
# CONFIGURACIÓN
# ========================================================
CONFIG = {
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


# ========================================================
# FUNCIONES DE LÓGICA
# ========================================================

def parse_time_str(time_str):
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
    if pd.isna(cell_value):
        return []
    s = str(cell_value).strip().lower()
    if s in ["libre", "disponible", "todo el día", "todo el dia"]:
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


def normalizar_dia(dia):
    if pd.isna(dia):
        return None
    d = str(dia).strip().lower()
    d = d.replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
    d = re.sub(r'[^a-z]', '', d)
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
    if pd.isna(valor):
        return 3
    try:
        prioridad = int(valor)
        if 1 <= prioridad <= 5:
            return prioridad
        return 3
    except:
        return 3


def cargar_monitores_desde_excel(ruta):
    """
    Carga monitores desde Excel (.xlsx, .xls) o LibreOffice (.ods)
    """
    cfg = CONFIG["monitores"]
    
    # Detectar tipo de archivo y leer apropiadamente
    if ruta.endswith('.ods'):
        # Usar engine específico para ODS
        df_raw = pd.read_excel(ruta, sheet_name=0, header=None, engine='odf')
    else:
        # Usar engine por defecto para Excel
        df_raw = pd.read_excel(ruta, sheet_name=0, header=None)
    
    dias_row = df_raw.iloc[3]
    jornadas_row = df_raw.iloc[cfg["header_row"]]
    col_mapping = {}
    current_dia = None
    col_nombre_idx = None
    col_prioridad_idx = None
    
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
    """
    cfg = CONFIG["espacios"]
    
    # Detectar tipo de archivo y leer apropiadamente
    if ruta.endswith('.ods'):
        df = pd.read_excel(ruta, sheet_name=0, engine='odf')
    else:
        df = pd.read_excel(ruta, sheet_name=0)
    
    columnas_req = [cfg["col_sala"], cfg["col_dia"], cfg["col_hora_inicio"], 
                    cfg["col_hora_fin"], cfg["col_curso"]]
    faltantes = [col for col in columnas_req if col not in df.columns]
    if faltantes:
        raise ValueError(f"Columnas no encontradas: {faltantes}")
    df['DIA_NORM'] = df[cfg["col_dia"]].apply(normalizar_dia)
    df['DURACION'] = df[cfg["col_hora_fin"]] - df[cfg["col_hora_inicio"]]
    return df


def esta_disponible(monitor, dia, hora_inicio, hora_fin):
    if dia not in monitor["disp"]:
        return False
    for r_inicio, r_fin in monitor["disp"][dia]:
        if hora_inicio >= r_inicio and hora_fin <= r_fin:
            return True
    return False


def verificar_restricciones(monitor, dia, hora_inicio, hora_fin):
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


def asignar_monitores(monitores, df_espacios):
    cfg_asig = CONFIG["asignacion"]
    cfg_esp = CONFIG["espacios"]
    asignaciones = []
    sin_monitor = []
    espacios = df_espacios.to_dict('records')
    
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
            ]
            if candidatos:
                candidatos.sort(key=lambda x: (x["prioridad"], x["min"] - x["horas"]), reverse=True)
                elegido = candidatos[0]
                elegido["horas"] += duracion
                elegido["asignaciones"].append({"dia": dia, "inicio": inicio, "fin": fin})
                asignaciones.append({**espacio, "MONITOR": elegido["nombre"], "PRIORIDAD": elegido["prioridad"], "ESTADO": "✅"})
    
    for espacio in espacios:
        dia = espacio['DIA_NORM']
        if pd.isna(dia):
            sin_monitor.append(espacio)
            asignaciones.append({**espacio, "MONITOR": "DÍA INVÁLIDO", "PRIORIDAD": "-", "ESTADO": "❌"})
            continue
        ya_asignado = any(
            a.get(cfg_esp["col_sala"]) == espacio[cfg_esp["col_sala"]] and
            a.get('DIA_NORM') == espacio['DIA_NORM'] and
            a.get(cfg_esp["col_hora_inicio"]) == espacio[cfg_esp["col_hora_inicio"]]
            for a in asignaciones
        )
        if ya_asignado:
            continue
        inicio = espacio[cfg_esp["col_hora_inicio"]]
        fin = espacio[cfg_esp["col_hora_fin"]]
        duracion = espacio['DURACION']
        candidatos = [
            m for m in monitores
            if m["horas"] + duracion <= m["max"]
            and esta_disponible(m, dia, inicio, fin)
            and verificar_restricciones(m, dia, inicio, fin)
        ]
        if not candidatos:
            sin_monitor.append(espacio)
            asignaciones.append({**espacio, "MONITOR": "SIN MONITOR", "PRIORIDAD": "-", "ESTADO": "❌"})
            continue
        if cfg_asig.get("usar_prioridad"):
            candidatos.sort(key=lambda x: (x["prioridad"], -x["horas"]), reverse=True)
        elif cfg_asig.get("balancear_carga"):
            candidatos.sort(key=lambda x: x["horas"])
        elegido = candidatos[0]
        elegido["horas"] += duracion
        elegido["asignaciones"].append({"dia": dia, "inicio": inicio, "fin": fin})
        asignaciones.append({**espacio, "MONITOR": elegido["nombre"], "PRIORIDAD": elegido["prioridad"], "ESTADO": "✅"})
    return asignaciones, sin_monitor, monitores


# ========================================================
# FUNCIONES DE EXPORTACIÓN CON HORARIOS VISUALES
# ========================================================

def crear_horario_consolidado(writer, df_asig, salas, cfg_esp):
    """Crea horario visual con todas las salas HORIZONTALMENTE"""
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    
    workbook = writer.book
    worksheet = workbook.create_sheet('Horarios Salas', 0)
    
    # Estilos
    color_naranja = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
    color_verde = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
    color_azul = PatternFill(start_color="87CEEB", end_color="87CEEB", fill_type="solid")
    color_vacio = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    color_sin_monitor = PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid")
    color_header = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    color_titulo_sala = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    
    fuente_negra = Font(bold=False, size=8, color="000000")
    fuente_header = Font(bold=True, size=9, color="000000")
    fuente_titulo = Font(bold=True, size=11, color="FFFFFF")
    
    alineacion_centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    borde = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    
    dias = ['Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado']
    
    hora_min = int(df_asig[cfg_esp["col_hora_inicio"]].min())
    hora_max = int(df_asig[cfg_esp["col_hora_fin"]].max())
    
    columna_actual = 1
    
    # Fila 1: Títulos de salas
    for sala in salas:
        worksheet.merge_cells(
            start_row=1, 
            start_column=columna_actual, 
            end_row=1, 
            end_column=columna_actual + 6
        )
        cell_titulo = worksheet.cell(row=1, column=columna_actual)
        cell_titulo.value = sala
        cell_titulo.fill = color_titulo_sala
        cell_titulo.font = fuente_titulo
        cell_titulo.alignment = alineacion_centro
        cell_titulo.border = borde
        
        # Fila 2: Encabezados de días
        cell_hora_header = worksheet.cell(row=2, column=columna_actual)
        cell_hora_header.value = "Hora"
        cell_hora_header.fill = color_header
        cell_hora_header.font = fuente_header
        cell_hora_header.alignment = alineacion_centro
        cell_hora_header.border = borde
        
        for idx_dia, dia in enumerate(dias, start=1):
            cell = worksheet.cell(row=2, column=columna_actual + idx_dia)
            cell.value = dia
            cell.fill = color_header
            cell.font = fuente_header
            cell.alignment = alineacion_centro
            cell.border = borde
        
        columna_actual += 7
    
    # Filas 3+: Horas y datos
    for fila_hora, hora in enumerate(range(hora_min, hora_max), start=3):
        hora_str = f"{hora}:00-{hora+1}:00"
        
        columna_actual = 1
        
        for sala in salas:
            df_sala = df_asig[df_asig[cfg_esp["col_sala"]] == sala]
            
            cell_hora = worksheet.cell(row=fila_hora, column=columna_actual)
            cell_hora.value = hora_str
            cell_hora.fill = color_header
            cell_hora.font = fuente_header
            cell_hora.alignment = alineacion_centro
            cell_hora.border = borde
            
            for idx_dia, dia in enumerate(dias, start=1):
                cell = worksheet.cell(row=fila_hora, column=columna_actual + idx_dia)
                cell.border = borde
                cell.alignment = alineacion_centro
                
                dia_norm = dia.lower()
                asignaciones_celda = df_sala[
                    (df_sala['DIA_NORM'] == dia_norm) &
                    (df_sala[cfg_esp["col_hora_inicio"]] <= hora) &
                    (df_sala[cfg_esp["col_hora_fin"]] > hora)
                ]
                
                if len(asignaciones_celda) > 0:
                    asig = asignaciones_celda.iloc[0]
                    curso = asig[cfg_esp["col_curso"]]
                    monitor = asig['MONITOR']
                    
                    if monitor == "SIN MONITOR":
                        cell.value = f"{curso}\n❌ SIN MONITOR"
                        cell.fill = color_sin_monitor
                        cell.font = Font(size=7, color="FF0000", bold=True)
                    else:
                        nombre_corto = monitor.split()[0] if monitor else ""
                        cell.value = f"{curso}\n{nombre_corto}"
                        cell.font = fuente_negra
                        
                        hash_val = hash(monitor) % 3
                        if hash_val == 0:
                            cell.fill = color_naranja
                        elif hash_val == 1:
                            cell.fill = color_verde
                        else:
                            cell.fill = color_azul
                else:
                    cell.value = ""
                    cell.fill = color_vacio
            
            columna_actual += 7
    
    # Ajustar anchos
    for col_num in range(1, columna_actual):
        col_letter = get_column_letter(col_num)
        if (col_num - 1) % 7 == 0:
            worksheet.column_dimensions[col_letter].width = 11
        else:
            worksheet.column_dimensions[col_letter].width = 18
    
    worksheet.row_dimensions[1].height = 25
    worksheet.row_dimensions[2].height = 20
    for row in range(3, fila_hora + 1):
        worksheet.row_dimensions[row].height = 35


def crear_horario_monitores(writer, monitores, df_asig, cfg_esp):
    """Crea horario detallado de cada monitor"""
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    
    workbook = writer.book
    worksheet = workbook.create_sheet('Horarios Monitores', 1)
    
    color_monitor = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    color_header = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    color_clase = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
    color_vacio = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    fuente_monitor = Font(bold=True, size=12, color="FFFFFF")
    fuente_header = Font(bold=True, size=10, color="000000")
    fuente_normal = Font(size=8, color="000000")
    
    alineacion_centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    borde = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    
    dias = ['Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado']
    
    hora_min = int(df_asig[cfg_esp["col_hora_inicio"]].min())
    hora_max = int(df_asig[cfg_esp["col_hora_fin"]].max())
    
    fila_actual = 1
    
    monitores_activos = [m for m in monitores if m["horas"] > 0]
    monitores_activos.sort(key=lambda x: x["nombre"])
    
    for monitor in monitores_activos:
        # Título del monitor
        worksheet.merge_cells(start_row=fila_actual, start_column=1, 
                             end_row=fila_actual, end_column=7)
        cell_titulo = worksheet.cell(row=fila_actual, column=1)
        cell_titulo.value = f"{monitor['nombre']} - {monitor['horas']} horas"
        cell_titulo.fill = color_monitor
        cell_titulo.font = fuente_monitor
        cell_titulo.alignment = alineacion_centro
        cell_titulo.border = borde
        fila_actual += 1
        
        # Encabezados
        worksheet.cell(row=fila_actual, column=1).value = "Hora"
        for idx, dia in enumerate(dias, start=2):
            cell = worksheet.cell(row=fila_actual, column=idx)
            cell.value = dia
            cell.fill = color_header
            cell.font = fuente_header
            cell.alignment = alineacion_centro
            cell.border = borde
        
        cell_hora_header = worksheet.cell(row=fila_actual, column=1)
        cell_hora_header.fill = color_header
        cell_hora_header.font = fuente_header
        cell_hora_header.alignment = alineacion_centro
        cell_hora_header.border = borde
        fila_actual += 1
        
        asignaciones_monitor = df_asig[df_asig['MONITOR'] == monitor['nombre']]
        
        for hora in range(hora_min, hora_max):
            hora_str = f"{hora}-{hora+1}"
            
            cell_hora = worksheet.cell(row=fila_actual, column=1)
            cell_hora.value = hora_str
            cell_hora.fill = color_header
            cell_hora.font = fuente_header
            cell_hora.alignment = alineacion_centro
            cell_hora.border = borde
            
            for idx_dia, dia in enumerate(dias, start=2):
                cell = worksheet.cell(row=fila_actual, column=idx_dia)
                cell.border = borde
                cell.alignment = alineacion_centro
                
                dia_norm = dia.lower()
                asig_celda = asignaciones_monitor[
                    (asignaciones_monitor['DIA_NORM'] == dia_norm) &
                    (asignaciones_monitor[cfg_esp["col_hora_inicio"]] <= hora) &
                    (asignaciones_monitor[cfg_esp["col_hora_fin"]] > hora)
                ]
                
                if len(asig_celda) > 0:
                    asig = asig_celda.iloc[0]
                    curso = asig[cfg_esp["col_curso"]]
                    sala = asig[cfg_esp["col_sala"]]
                    
                    cell.value = f"{sala}\n{curso}"
                    cell.fill = color_clase
                    cell.font = fuente_normal
                else:
                    cell.value = ""
                    cell.fill = color_vacio
            
            fila_actual += 1
        
        fila_actual += 2
    
    worksheet.column_dimensions['A'].width = 10
    for col in range(2, 8):
        worksheet.column_dimensions[get_column_letter(col)].width = 22
    
    for row in range(1, fila_actual):
        worksheet.row_dimensions[row].height = 35


# ========================================================
# MODELOS Y THREADS
# ========================================================

class PandasModel(QAbstractTableModel):
    def __init__(self, df=pd.DataFrame()):
        super().__init__()
        self._df = df

    def rowCount(self, parent=None):
        return self._df.shape[0]

    def columnCount(self, parent=None):
        return self._df.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            return str(self._df.iat[index.row(), index.column()])

    def headerData(self, section, orientation, role):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return self._df.columns[section]
            return section


class AsignacionThread(QThread):
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
            asignaciones, sin_monitor, monitores = asignar_monitores(self.monitores, self.df_espacios)
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


class DownloadDialog(QDialog):
    def __init__(self, df_resultado, monitores_asignados, parent=None):
        super().__init__(parent)
        self.df_resultado = df_resultado
        self.monitores_asignados = monitores_asignados
        self.archivo_guardado = None
        self.setWindowTitle("Asignación Completada")
        self.setMinimumSize(500, 350)
        self.setModal(True)
        self.setStyleSheet("""
            QDialog {background: white; border-radius: 16px;}
            QLabel {color: #1F2937; background: transparent;}
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #667EEA, stop:1 #564FEE);
                color: white; border: none; padding: 12px 24px; border-radius: 10px;
                font-weight: 600; font-size: 13px; min-height: 40px;
            }
            QPushButton:hover {background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #7C8FF5, stop:1 #6B64F8);}
            QPushButton#btnSecondary {background: #F3F4F6; color: #374151;}
            QPushButton#btnSecondary:hover {background: #E5E7EB;}
        """)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(32, 32, 32, 32)
        layout.setSpacing(20)
        icon = QLabel("✅")
        icon.setFont(QFont("Inter", 64))
        icon.setAlignment(Qt.AlignCenter)
        icon.setStyleSheet("color: #10B981;")
        layout.addWidget(icon)
        title = QLabel("¡Asignación Completada!")
        title.setFont(QFont("Inter", 20, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        exitosos = len([a for a in df_resultado.to_dict('records') if a.get("ESTADO") == "✅"])
        total = len(df_resultado)
        porcentaje = f"{exitosos*100/total:.1f}%" if total > 0 else "0%"
        stats_frame = QFrame()
        stats_frame.setStyleSheet("QFrame {background: #F9FAFB; border-radius: 12px; padding: 16px;}")
        stats_layout = QVBoxLayout(stats_frame)
        stats_text = f"<div style='text-align: center;'><p style='font-size: 14px; color: #6B7280;'>Se asignaron correctamente</p><p style='font-size: 32px; font-weight: bold; color: #667EEA;'>{exitosos}/{total}</p><p style='font-size: 14px; color: #6B7280;'>horarios ({porcentaje})</p></div>"
        stats_label = QLabel(stats_text)
        stats_label.setAlignment(Qt.AlignCenter)
        stats_layout.addWidget(stats_label)
        layout.addWidget(stats_frame)
        message = QLabel("¿Deseas descargar el archivo con horarios visuales ahora?")
        message.setFont(QFont("Inter", 13))
        message.setAlignment(Qt.AlignCenter)
        message.setStyleSheet("color: #6B7280;")
        layout.addWidget(message)
        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(12)
        self.btn_descargar = QPushButton("💾 Descargar Ahora")
        self.btn_descargar.clicked.connect(self.descargar)
        self.btn_despues = QPushButton("Descargar Después")
        self.btn_despues.setObjectName("btnSecondary")
        self.btn_despues.clicked.connect(self.reject)
        buttons_layout.addWidget(self.btn_despues)
        buttons_layout.addWidget(self.btn_descargar)
        layout.addLayout(buttons_layout)
    
    def descargar(self):
        ruta, _ = QFileDialog.getSaveFileName(self, "Guardar archivo de resultados", "Asignacion_Monitores.xlsx", "Archivos Excel (*.xlsx)")
        if ruta:
            try:
                cfg_esp = CONFIG["espacios"]
                
                # Preparar datos
                df_mon = pd.DataFrame([{
                    'Monitor': m['nombre'], 'Prioridad': m['prioridad'], 'Horas': m['horas'],
                    'Min': m['min'], 'Max': m['max'], 'Horarios': len(m['asignaciones']),
                    'Estado': '✅' if m['min'] <= m['horas'] <= m['max'] else '⚠️'
                } for m in self.monitores_asignados]).sort_values('Horas', ascending=False)
                
                salas = sorted(self.df_resultado[cfg_esp["col_sala"]].unique())
                
                # Crear archivo Excel con horarios visuales
                with pd.ExcelWriter(ruta, engine='openpyxl') as writer:
                    # HOJA 1: Horarios visuales de todas las salas
                    crear_horario_consolidado(writer, self.df_resultado, salas, cfg_esp)
                    
                    # HOJA 2: Horarios individuales por monitor
                    crear_horario_monitores(writer, self.monitores_asignados, self.df_resultado, cfg_esp)
                    
                    # HOJA 3: Lista de asignaciones
                    self.df_resultado.to_excel(writer, sheet_name='Lista Asignaciones', index=False)
                    
                    # HOJA 4: Resumen de monitores
                    df_mon.to_excel(writer, sheet_name='Resumen Monitores', index=False)
                
                self.archivo_guardado = ruta
                success_msg = QMessageBox(self)
                success_msg.setIcon(QMessageBox.Information)
                success_msg.setWindowTitle("Descarga Exitosa")
                success_msg.setText("✅ Archivo guardado exitosamente")
                success_msg.setInformativeText(f"📁 {ruta}\n\n📄 4 hojas generadas:\n• Horarios Salas (visual)\n• Horarios Monitores (individual)\n• Lista Asignaciones\n• Resumen Monitores")
                success_msg.exec()
                self.accept()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error al guardar el archivo:\n{str(e)}")


class ModernCard(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("ModernCard {background: white; border-radius: 16px; border: 1px solid #E5E7EB;}")


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sistema de Asignación de Monitores")
        self.setMinimumSize(1400, 900)
        self.setStyleSheet("""
            QWidget {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #F0F4F8, stop:1 #E5E9F0);
                font-family: 'Inter', 'Segoe UI', sans-serif;
            }
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #667EEA, stop:1 #564FEE);
                color: white; border: none; padding: 12px 24px; border-radius: 10px;
                font-weight: 600; font-size: 13px; min-height: 40px;
            }
            QPushButton:hover {background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #7C8FF5, stop:1 #6B64F8);}
            QPushButton:disabled {background: #D1D5DB; color: #9CA3AF;}
        """)

        main_layout = QVBoxLayout()
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(24, 24, 24, 24)

        # STATS CARDS
        stats_layout = QHBoxLayout()
        stats_layout.setSpacing(16)
        self.card_monitores = self.create_stat_card("👥", "Monitores", "0", "#8B5CF6")
        self.card_espacios = self.create_stat_card("📅", "Horarios", "0", "#3B82F6")
        self.card_asignados = self.create_stat_card("✅", "Asignados", "0%", "#10B981")
        stats_layout.addWidget(self.card_monitores)
        stats_layout.addWidget(self.card_espacios)
        stats_layout.addWidget(self.card_asignados)
        main_layout.addLayout(stats_layout)

        # ROADMAP
        roadmap_card = ModernCard()
        roadmap_layout = QVBoxLayout(roadmap_card)
        roadmap_layout.setContentsMargins(20, 16, 20, 16)
        roadmap_layout.setSpacing(12)
        
        header_roadmap = QHBoxLayout()
        roadmap_title = QLabel("🗺️ Proceso de Asignación")
        roadmap_title.setFont(QFont("Inter", 13, QFont.Bold))
        roadmap_title.setStyleSheet("color: #111827; background: transparent;")
        header_roadmap.addWidget(roadmap_title)
        header_roadmap.addStretch()
        
        self.btn_config = QPushButton("⚙️")
        self.btn_config.setFixedSize(50, 50)
        self.btn_config.setToolTip("Configuración del sistema")
        self.btn_config.clicked.connect(self.toggle_config_panel)
        self.btn_config.setStyleSheet("""
            QPushButton {background: #F3F4F6; color: #374151; font-size: 22px; padding: 2px; border-radius: 10px;}
            QPushButton:hover {background: #E5E7EB;}
        """)
        header_roadmap.addWidget(self.btn_config)
        roadmap_layout.addLayout(header_roadmap)
        
        steps_container = QHBoxLayout()
        steps_container.setSpacing(0)
        steps_container.setContentsMargins(0, 0, 0, 0)
        
        self.step1_frame = self.create_path_step("1", "Cargar Monitores", "👥", "pending")
        self.step1_frame.mousePressEvent = lambda e: self.cargar_monitores()
        self.step1_frame.setCursor(Qt.PointingHandCursor)
        steps_container.addWidget(self.step1_frame)
        
        self.connector1 = self.create_path_connector("pending")
        steps_container.addWidget(self.connector1)
        
        self.step2_frame = self.create_path_step("2", "Cargar Espacios", "📅", "pending")
        self.step2_frame.mousePressEvent = lambda e: self.cargar_espacios()
        self.step2_frame.setCursor(Qt.PointingHandCursor)
        steps_container.addWidget(self.step2_frame)
        
        self.connector2 = self.create_path_connector("pending")
        steps_container.addWidget(self.connector2)
        
        self.step3_frame = self.create_path_step("3", "Asignar Monitores", "⚡", "pending")
        self.step3_frame.mousePressEvent = lambda e: self.iniciar_asignacion() if self.step1_frame.status == "success" and self.step2_frame.status == "success" else None
        self.step3_frame.setCursor(Qt.PointingHandCursor)
        steps_container.addWidget(self.step3_frame)
        
        self.connector3 = self.create_path_connector("pending")
        steps_container.addWidget(self.connector3)
        
        self.step4_frame = self.create_path_step("4", "Exportar Resultados", "💾", "pending")
        self.step4_frame.mousePressEvent = lambda e: self.exportar() if not self.df_resultado.empty else None
        self.step4_frame.setCursor(Qt.PointingHandCursor)
        steps_container.addWidget(self.step4_frame)
        
        roadmap_layout.addLayout(steps_container)
        main_layout.addWidget(roadmap_card)

        # CONFIG OVERLAY
        self.config_overlay = QFrame(self)
        self.config_overlay.setVisible(False)
        self.config_overlay.setStyleSheet("QFrame {background: rgba(0, 0, 0, 0.5);}")
        self.config_overlay.mousePressEvent = lambda e: self.toggle_config_panel()
        overlay_layout = QVBoxLayout(self.config_overlay)
        overlay_layout.setContentsMargins(0, 0, 0, 0)
        overlay_layout.setAlignment(Qt.AlignCenter)
        
        self.config_card = ModernCard()
        self.config_card.setMaximumWidth(900)
        self.config_card.setMaximumHeight(280)
        self.config_card.mousePressEvent = lambda e: e.accept()
        config_layout = QVBoxLayout(self.config_card)
        config_layout.setContentsMargins(32, 24, 32, 24)
        config_layout.setSpacing(20)
        
        config_header = QHBoxLayout()
        config_icon = QLabel("⚙️")
        config_icon.setFont(QFont("Inter", 20))
        config_icon.setStyleSheet("background: transparent;")
        config_header.addWidget(config_icon)
        config_title = QLabel("Configuración del Sistema")
        config_title.setFont(QFont("Inter", 16, QFont.Bold))
        config_title.setStyleSheet("color: #111827; background: transparent;")
        config_header.addWidget(config_title)
        config_header.addStretch()
        btn_close = QPushButton("✕ Cerrar")
        btn_close.setFixedHeight(36)
        btn_close.clicked.connect(self.toggle_config_panel)
        btn_close.setStyleSheet("QPushButton {background: #F3F4F6; color: #6B7280; padding: 0px 16px; border-radius: 8px;} QPushButton:hover {background: #E5E7EB;}")
        config_header.addWidget(btn_close)
        config_layout.addLayout(config_header)
        
        options_layout = QHBoxLayout()
        options_layout.setSpacing(16)
        
        self.chk_usar_prioridad = QCheckBox("🎖️  Usar prioridades")
        self.chk_usar_prioridad.setChecked(True)
        self.chk_usar_prioridad.stateChanged.connect(self.toggle_prioridad)
        self.chk_balancear = QCheckBox("⚖️  Balancear carga")
        self.chk_balancear.setChecked(True)
        self.chk_priorizar_min = QCheckBox("🎯  Priorizar mínimo")
        self.chk_priorizar_min.setChecked(True)
        options_layout.addWidget(self.chk_usar_prioridad)
        options_layout.addWidget(self.chk_balancear)
        options_layout.addWidget(self.chk_priorizar_min)
        config_layout.addLayout(options_layout)
        overlay_layout.addWidget(self.config_card)

        # PROGRESS BAR
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setTextVisible(False)
        self.progress.setMaximumHeight(6)
        main_layout.addWidget(self.progress)

        # STATUS
        self.lbl_estado = QLabel("📋 Esperando archivos...")
        self.lbl_estado.setFont(QFont("Inter", 12))
        self.lbl_estado.setStyleSheet("color: #6B7280; background: white; padding: 10px 16px; border-radius: 8px; border: 1px solid #E5E7EB;")
        main_layout.addWidget(self.lbl_estado)

        # CONTENT AREA CON PESTAÑAS
        content_layout = QHBoxLayout()
        content_layout.setSpacing(12)
        table_card = ModernCard()
        table_layout = QVBoxLayout(table_card)
        table_layout.setContentsMargins(0, 0, 0, 0)
        table_header = QLabel("📊 Visualización de Datos")
        table_header.setFont(QFont("Inter", 13, QFont.Bold))
        table_header.setStyleSheet("color: #1F2937; padding: 14px 16px; background: #F9FAFB; border-radius: 12px 12px 0 0;")
        table_layout.addWidget(table_header)
        
        self.tabs = QTabWidget()
        self.table_monitores = QTableView()
        self.table_espacios = QTableView()
        self.table_asignaciones = QTableView()
        self.table_resumen = QTableView()
        self.tabs.addTab(self.table_monitores, "👥  Monitores")
        self.tabs.addTab(self.table_espacios, "📅  Horarios")
        self.tabs.addTab(self.table_asignaciones, "✅  Asignaciones")
        self.tabs.addTab(self.table_resumen, "📈  Resumen")
        table_layout.addWidget(self.tabs)
        content_layout.addWidget(table_card, stretch=6)
        
        report_card = ModernCard()
        report_layout = QVBoxLayout(report_card)
        report_layout.setContentsMargins(0, 0, 0, 0)
        report_header = QLabel("📝 Reporte")
        report_header.setFont(QFont("Inter", 13, QFont.Bold))
        report_header.setStyleSheet("color: #1F2937; padding: 14px 16px; background: #F9FAFB; border-radius: 12px 12px 0 0;")
        report_layout.addWidget(report_header)
        self.text_reporte = QTextEdit()
        self.text_reporte.setReadOnly(True)
        self.text_reporte.setPlaceholderText("El reporte detallado aparecerá aquí...")
        report_layout.addWidget(self.text_reporte)
        content_layout.addWidget(report_card, stretch=4)
        main_layout.addLayout(content_layout, stretch=1)
        self.setLayout(main_layout)

        self.monitores = []
        self.df_espacios = pd.DataFrame()
        self.df_resultado = pd.DataFrame()
        self.monitores_asignados = []

    def create_path_step(self, number, title, icon, status="pending"):
        frame = QFrame()
        frame.setMinimumHeight(90)
        frame.setMaximumHeight(90)
        frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        
        if status == "pending":
            bg_gradient = "stop:0 #F9FAFB, stop:1 #F3F4F6"
            border_color = "#E5E7EB"
            text_color = "#9CA3AF"
            number_bg = "#E5E7EB"
            number_color = "#6B7280"
        elif status == "success":
            bg_gradient = "stop:0 #D1FAE5, stop:1 #A7F3D0"
            border_color = "#10B981"
            text_color = "#065F46"
            number_bg = "#10B981"
            number_color = "white"
        else:
            bg_gradient = "stop:0 #FEE2E2, stop:1 #FECACA"
            border_color = "#EF4444"
            text_color = "#991B1B"
            number_bg = "#EF4444"
            number_color = "white"
        
        frame.setStyleSheet(f"""
            QFrame {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, {bg_gradient});
                border: 2px solid {border_color};
                border-radius: 12px;
            }}
        """)
        
        layout = QHBoxLayout(frame)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
        layout.setAlignment(Qt.AlignCenter)
        
        number_label = QLabel(number)
        number_label.setAlignment(Qt.AlignCenter)
        number_label.setFixedSize(36, 36)
        number_label.setFont(QFont("Inter", 16, QFont.Bold))
        number_label.setStyleSheet(f"QLabel {{background: {number_bg}; color: {number_color}; border-radius: 18px; border: none;}}")
        layout.addWidget(number_label)
        
        icon_label = QLabel(icon)
        icon_label.setFont(QFont("Inter", 32))
        icon_label.setAlignment(Qt.AlignCenter)
        icon_label.setStyleSheet(f"color: {text_color}; background: transparent; border: none;")
        layout.addWidget(icon_label)
        
        title_label = QLabel(title)
        title_label.setFont(QFont("Inter", 12, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setWordWrap(True)
        title_label.setStyleSheet(f"color: {text_color}; background: transparent; border: none;")
        layout.addWidget(title_label)
        
        frame.number_label = number_label
        frame.icon_label = icon_label
        frame.title_label = title_label
        frame.status = status
        return frame
    
    def create_path_connector(self, status="pending"):
        connector = QFrame()
        connector.setFixedSize(40, 90)
        connector.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        
        if status == "pending":
            bg_color = "#E5E7EB"
        elif status == "success":
            bg_color = "#10B981"
        else:
            bg_color = "#EF4444"
        
        connector.setStyleSheet(f"""
            QFrame {{
                background: {bg_color};
                border: none;
            }}
        """)
        connector.status = status
        return connector
    
    def update_step_status(self, step_frame, status):
        if status == "pending":
            bg_gradient = "stop:0 #F9FAFB, stop:1 #F3F4F6"
            border_color = "#E5E7EB"
            text_color = "#9CA3AF"
            number_bg = "#E5E7EB"
            number_color = "#6B7280"
        elif status == "success":
            bg_gradient = "stop:0 #D1FAE5, stop:1 #A7F3D0"
            border_color = "#10B981"
            text_color = "#065F46"
            number_bg = "#10B981"
            number_color = "white"
        else:
            bg_gradient = "stop:0 #FEE2E2, stop:1 #FECACA"
            border_color = "#EF4444"
            text_color = "#991B1B"
            number_bg = "#EF4444"
            number_color = "white"
        
        step_frame.setStyleSheet(f"""
            QFrame {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, {bg_gradient});
                border: 2px solid {border_color};
                border-radius: 12px;
            }}
        """)
        step_frame.number_label.setStyleSheet(f"QLabel {{background: {number_bg}; color: {number_color}; border-radius: 14px;}}")
        step_frame.icon_label.setStyleSheet(f"color: {text_color}; background: transparent;")
        step_frame.title_label.setStyleSheet(f"color: {text_color}; background: transparent;")
        step_frame.status = status
        self.update_connectors()
    
    def update_connectors(self):
        if hasattr(self, 'step1_frame') and self.step1_frame.status == "success":
            self.connector1.setStyleSheet("QFrame {background: #10B981;}")
        else:
            self.connector1.setStyleSheet("QFrame {background: #E5E7EB;}")
        
        if hasattr(self, 'step2_frame') and self.step2_frame.status == "success":
            self.connector2.setStyleSheet("QFrame {background: #10B981;}")
        else:
            self.connector2.setStyleSheet("QFrame {background: #E5E7EB;}")
        
        if hasattr(self, 'step3_frame') and self.step3_frame.status == "success":
            self.connector3.setStyleSheet("QFrame {background: #10B981;}")
        else:
            self.connector3.setStyleSheet("QFrame {background: #E5E7EB;}")

    def create_stat_card(self, icon, title, value, color):
        card = QFrame()
        card.setStyleSheet(f"QFrame {{background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {color}, stop:1 {self.darken_color(color)}); border-radius: 14px;}}")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(18, 14, 18, 14)
        header_layout = QHBoxLayout()
        icon_label = QLabel(icon)
        icon_label.setFont(QFont("Inter", 18))
        icon_label.setStyleSheet("color: white; background: transparent;")
        header_layout.addWidget(icon_label)
        title_label = QLabel(title.upper())
        title_label.setFont(QFont("Inter", 9, QFont.Bold))
        title_label.setStyleSheet("color: rgba(255, 255, 255, 0.9); background: transparent;")
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        layout.addLayout(header_layout)
        value_label = QLabel(value)
        value_label.setFont(QFont("Inter", 24, QFont.Bold))
        value_label.setStyleSheet("color: white; background: transparent;")
        layout.addWidget(value_label)
        card.value_label = value_label
        return card
    
    def darken_color(self, hex_color):
        hex_color = hex_color.lstrip('#')
        r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        factor = 0.8
        r, g, b = int(r * factor), int(g * factor), int(b * factor)
        return f'#{r:02x}{g:02x}{b:02x}'

    def cargar_monitores(self):
        ruta, _ = QFileDialog.getOpenFileName(
            self, 
            "Seleccionar archivo de monitores", 
            "", 
            "Hojas de Cálculo (*.xlsx *.xls *.ods);;Excel (*.xlsx *.xls);;LibreOffice (*.ods)"
        )
        if ruta:
            try:
                self.monitores = cargar_monitores_desde_excel(ruta)
                df_preview = pd.DataFrame([{'Nombre': m['nombre'], 'Prioridad': f"⭐{m['prioridad']}", 'Min': m['min'], 'Max': m['max']} for m in self.monitores])
                self.table_monitores.setModel(PandasModel(df_preview))
                self.tabs.setCurrentIndex(0)
                self.card_monitores.value_label.setText(str(len(self.monitores)))
                self.lbl_estado.setText(f"✅ {len(self.monitores)} monitores cargados exitosamente")
                self.text_reporte.setPlainText(f"📂 Monitores cargados: {len(self.monitores)}")
                self.verificar_listo()
                self.update_step_status(self.step1_frame, "success")
            except Exception as e:
                self.update_step_status(self.step1_frame, "error")
                QMessageBox.critical(self, "Error", f"Error al cargar monitores:\n{str(e)}")

    def cargar_espacios(self):
        # CAMBIAR LA LÍNEA DEL FILTRO DE ARCHIVOS
        ruta, _ = QFileDialog.getOpenFileName(
            self, 
            "Seleccionar archivo de espacios", 
            "", 
            "Hojas de Cálculo (*.xlsx *.xls *.ods);;Excel (*.xlsx *.xls);;LibreOffice (*.ods)"
        )
        if ruta:
            try:
                self.df_espacios = cargar_espacios_desde_excel(ruta)
                self.table_espacios.setModel(PandasModel(self.df_espacios.head(100)))
                self.tabs.setCurrentIndex(1)
                self.card_espacios.value_label.setText(str(len(self.df_espacios)))
                self.lbl_estado.setText(f"✅ {len(self.df_espacios)} horarios cargados exitosamente")
                self.verificar_listo()
                self.update_step_status(self.step2_frame, "success")
            except Exception as e:
                self.update_step_status(self.step2_frame, "error")
                QMessageBox.critical(self, "Error", f"Error al cargar espacios:\n{str(e)}")

    def verificar_listo(self):
        if len(self.monitores) > 0 and len(self.df_espacios) > 0:
            self.lbl_estado.setText("✅ Todo listo! Haz click en 'Asignar' para comenzar")

    def iniciar_asignacion(self):
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)
        CONFIG["asignacion"]["usar_prioridad"] = self.chk_usar_prioridad.isChecked()
        CONFIG["asignacion"]["balancear_carga"] = self.chk_balancear.isChecked()
        CONFIG["asignacion"]["priorizar_minimo"] = self.chk_priorizar_min.isChecked()
        import copy
        monitores_copy = copy.deepcopy(self.monitores)
        self.thread = AsignacionThread(monitores_copy, self.df_espacios)
        self.thread.finished.connect(self.asignacion_completada)
        self.thread.error.connect(self.asignacion_error)
        self.thread.progress.connect(self.actualizar_progreso)
        self.thread.start()
    
    def toggle_prioridad(self, state):
        if state == Qt.Checked:
            self.lbl_estado.setText("✅ Sistema de prioridades ACTIVADO")
        else:
            self.lbl_estado.setText("⚠️ Sistema de prioridades DESACTIVADO")

    def toggle_config_panel(self):
        if self.config_overlay.isVisible():
            self.config_overlay.setVisible(False)
        else:
            self.config_overlay.setVisible(True)
            self.config_overlay.raise_()

    def actualizar_progreso(self, mensaje):
        self.lbl_estado.setText(mensaje)

    def asignacion_completada(self, df_resultado, monitores, reporte):
        self.df_resultado = df_resultado
        self.monitores_asignados = monitores
        self.table_asignaciones.setModel(PandasModel(df_resultado))
        self.tabs.setCurrentIndex(2)
        df_mon = pd.DataFrame([{'Monitor': m['nombre'], 'Prioridad': m['prioridad'], 'Horas': m['horas'], 'Min': m['min'], 'Max': m['max'], 'Horarios': len(m['asignaciones']), 'Estado': '✅' if m['min'] <= m['horas'] <= m['max'] else '⚠️'} for m in monitores if m['horas'] > 0]).sort_values('Horas', ascending=False)
        self.table_resumen.setModel(PandasModel(df_mon))
        self.text_reporte.setPlainText(reporte)
        exitosos = len([a for a in df_resultado.to_dict('records') if a.get("ESTADO") == "✅"])
        total = len(df_resultado)
        self.card_asignados.value_label.setText(f"{exitosos}/{total}")
        self.progress.setVisible(False)
        self.lbl_estado.setText(f"✅ Asignación completada exitosamente!")
        if exitosos > 0:
            self.update_step_status(self.step3_frame, "success")
        else:
            self.update_step_status(self.step3_frame, "error")
        dialog = DownloadDialog(df_resultado, monitores, self)
        resultado = dialog.exec()
        if resultado == QDialog.Accepted and dialog.archivo_guardado:
            self.lbl_estado.setText(f"✅ Archivo exportado: {dialog.archivo_guardado}")
            self.update_step_status(self.step4_frame, "success")

    def asignacion_error(self, error):
        self.progress.setVisible(False)
        self.update_step_status(self.step3_frame, "error")
        QMessageBox.critical(self, "Error", f"Error en la asignación:\n{error}")
        self.lbl_estado.setText("❌ Error en la asignación")

    def exportar(self):
        if self.df_resultado.empty:
            QMessageBox.warning(self, "Advertencia", "No hay resultados para exportar")
            return
        ruta, _ = QFileDialog.getSaveFileName(self, "Guardar archivo", "Asignacion_Monitores.xlsx", "Archivos Excel (*.xlsx)")
        if ruta:
            try:
                cfg_esp = CONFIG["espacios"]
                
                df_mon = pd.DataFrame([{
                    'Monitor': m['nombre'], 'Prioridad': m['prioridad'], 'Horas': m['horas'],
                    'Min': m['min'], 'Max': m['max'], 'Horarios': len(m['asignaciones']),
                    'Estado': '✅' if m['min'] <= m['horas'] <= m['max'] else '⚠️'
                } for m in self.monitores_asignados]).sort_values('Horas', ascending=False)
                
                salas = sorted(self.df_resultado[cfg_esp["col_sala"]].unique())
                
                with pd.ExcelWriter(ruta, engine='openpyxl') as writer:
                    # HOJA 1: Horarios visuales de todas las salas
                    crear_horario_consolidado(writer, self.df_resultado, salas, cfg_esp)
                    
                    # HOJA 2: Horarios individuales por monitor
                    crear_horario_monitores(writer, self.monitores_asignados, self.df_resultado, cfg_esp)
                    
                    # HOJA 3: Lista de asignaciones
                    self.df_resultado.to_excel(writer, sheet_name='Lista Asignaciones', index=False)
                    
                    # HOJA 4: Resumen de monitores
                    df_mon.to_excel(writer, sheet_name='Resumen Monitores', index=False)
                
                QMessageBox.information(self, "Exportado", f"✅ Archivo guardado exitosamente:\n{ruta}\n\n📄 4 hojas generadas:\n• Horarios Salas (visual)\n• Horarios Monitores (individual)\n• Lista Asignaciones\n• Resumen Monitores")
                self.lbl_estado.setText(f"✅ Exportado: {ruta}")
                self.update_step_status(self.step4_frame, "success")
            except Exception as e:
                self.update_step_status(self.step4_frame, "error")
                QMessageBox.critical(self, "Error", f"Error al exportar:\n{str(e)}")

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, 'config_overlay'):
            self.config_overlay.setGeometry(0, 0, self.width(), self.height())


if __name__ == "__main__":
    app = QApplication(sys.argv)
    font = QFont("Inter")
    if not font.exactMatch():
        font = QFont("Segoe UI")
    app.setFont(font)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())