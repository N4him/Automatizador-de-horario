import sys
import pandas as pd
import re
from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, 
    QHBoxLayout, QTableView, QFileDialog, QLabel, QMessageBox,
    QProgressBar, QTextEdit
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
        "max_horas_seguidas": 4,
        "descanso_minimo": 1,
        "permitir_sobrepasar_max": False
    }
}


# ========================================================
# FUNCIONES DE LÓGICA DE ASIGNACIÓN
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
    
    match = re.search(r'\d+', s)
    if match:
        return int(match.group(0))
    
    return None


def parse_range_cell(cell_value):
    """Convierte '7:00am-1:00pm' -> [(7, 13)]"""
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
    """Normaliza nombres de días"""
    if pd.isna(dia):
        return None
    
    d = str(dia).strip().lower()
    d = d.replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
    d = re.sub(r'[^a-z]', '', d)
    
    mapeo = {
        'lun': 'lunes',
        'mar': 'martes',
        'mie': 'miercoles',
        'jue': 'jueves',
        'vie': 'viernes',
        'sab': 'sabado',
        'dom': 'domingo'
    }
    
    for abrev, completo in mapeo.items():
        if d.startswith(abrev):
            return completo
    
    if d in mapeo.values():
        return d
    
    return d if len(d) >= 3 else None


def cargar_monitores_desde_excel(ruta):
    """Carga monitores desde Excel"""
    cfg = CONFIG["monitores"]
    
    df_raw = pd.read_excel(ruta, sheet_name=0, header=None)
    
    dias_row = df_raw.iloc[3]
    jornadas_row = df_raw.iloc[cfg["header_row"]]
    
    col_mapping = {}
    current_dia = None
    col_nombre_idx = None
    
    for idx, val in enumerate(jornadas_row):
        val_str = str(val).strip()
        
        if val_str == cfg["col_nombre"]:
            col_nombre_idx = idx
        
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
        
        mon = {
            "id": row_idx - cfg["data_start_row"],
            "nombre": str(nombre).strip(),
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
    """Carga espacios desde Excel"""
    cfg = CONFIG["espacios"]
    
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
    """Verifica disponibilidad del monitor"""
    if dia not in monitor["disp"]:
        return False
    
    for r_inicio, r_fin in monitor["disp"][dia]:
        if hora_inicio >= r_inicio and hora_fin <= r_fin:
            return True
    
    return False


def verificar_restricciones(monitor, dia, hora_inicio, hora_fin):
    """Verifica restricciones adicionales"""
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
    """Algoritmo principal de asignación"""
    cfg_asig = CONFIG["asignacion"]
    cfg_esp = CONFIG["espacios"]
    
    asignaciones = []
    sin_monitor = []
    
    espacios = df_espacios.to_dict('records')
    
    # Fase 1: Priorizar mínimo
    if cfg_asig.get("priorizar_minimo"):
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
                candidatos.sort(key=lambda x: x["min"] - x["horas"], reverse=True)
                elegido = candidatos[0]
                
                elegido["horas"] += duracion
                elegido["asignaciones"].append({
                    "dia": dia,
                    "inicio": inicio,
                    "fin": fin
                })
                
                asignaciones.append({
                    **espacio,
                    "MONITOR": elegido["nombre"],
                    "ESTADO": "✅"
                })
    
    # Fase 2: Asignar restantes
    for espacio in espacios:
        dia = espacio['DIA_NORM']
        if pd.isna(dia):
            sin_monitor.append(espacio)
            asignaciones.append({
                **espacio,
                "MONITOR": "DÍA INVÁLIDO",
                "ESTADO": "❌"
            })
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
            asignaciones.append({
                **espacio,
                "MONITOR": "SIN MONITOR",
                "ESTADO": "❌"
            })
            continue
        
        if cfg_asig.get("balancear_carga"):
            candidatos.sort(key=lambda x: x["horas"])
        
        elegido = candidatos[0]
        elegido["horas"] += duracion
        elegido["asignaciones"].append({
            "dia": dia,
            "inicio": inicio,
            "fin": fin
        })
        
        asignaciones.append({
            **espacio,
            "MONITOR": elegido["nombre"],
            "ESTADO": "✅"
        })
    
    return asignaciones, sin_monitor, monitores


# ========================================================
# MODELO PARA TABLA
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


# ========================================================
# HILO PARA PROCESAMIENTO
# ========================================================
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
            self.progress.emit("🔄 Iniciando asignación...")
            
            asignaciones, sin_monitor, monitores = asignar_monitores(
                self.monitores, 
                self.df_espacios
            )
            
            df_result = pd.DataFrame(asignaciones)
            
            # Generar reporte
            exitosos = len([a for a in asignaciones if a["ESTADO"] == "✅"])
            total = len(asignaciones)
            
            reporte = f"""
📊 REPORTE DE ASIGNACIÓN
{'='*50}

🎯 Resumen:
   Total horarios: {total}
   Asignados: {exitosos} ({exitosos*100/total:.1f}%)
   Sin monitor: {len(sin_monitor)} ({len(sin_monitor)*100/total:.1f}%)

👥 Monitores:
"""
            
            for m in sorted(monitores, key=lambda x: x["horas"], reverse=True):
                if m["horas"] > 0:
                    status = "✅"
                    if m["horas"] < m["min"]:
                        status = f"⚠️ <{m['min']}h"
                    elif m["horas"] > m["max"]:
                        status = f"❌ >{m['max']}h"
                    
                    reporte += f"\n   {m['nombre'][:30]:30} | {m['horas']:2}h {status}"
            
            self.progress.emit("✅ Asignación completada")
            self.finished.emit(df_result, monitores, reporte)
            
        except Exception as e:
            self.error.emit(str(e))


# ========================================================
# VENTANA PRINCIPAL
# ========================================================
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Gestor de Monitores – Sistema Completo")
        self.setMinimumSize(1100, 700)

        self.setStyleSheet("""
            QWidget {
                background-color: #F4F6F9;
                font-family: 'Segoe UI';
                font-size: 14px;
            }
            QPushButton {
                background-color: #0078D4;
                color: white;
                padding: 12px;
                border-radius: 8px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #005A9E;
            }
            QPushButton:disabled {
                background-color: #CCCCCC;
                color: #666666;
            }
            QTableView {
                background: white;
                border-radius: 8px;
                border: 1px solid #DDD;
            }
            QTextEdit {
                background: white;
                border-radius: 8px;
                border: 1px solid #DDD;
                padding: 10px;
                font-family: 'Consolas', monospace;
                font-size: 12px;
            }
            QLabel {
                color: #333;
            }
        """)

        layout = QVBoxLayout()

        # Título
        title = QLabel("🎯 Sistema de Asignación de Monitores")
        title.setFont(QFont("Segoe UI", 22, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # Subtítulo con instrucciones
        subtitle = QLabel("1️⃣ Carga Monitores → 2️⃣ Carga Espacios → 3️⃣ Asignar → 4️⃣ Exportar")
        subtitle.setFont(QFont("Segoe UI", 11))
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color: #666; margin-bottom: 10px;")
        layout.addWidget(subtitle)

        # Botones superiores
        btn_layout = QHBoxLayout()

        self.btn_monitores = QPushButton("📁 Cargar Monitores")
        self.btn_espacios = QPushButton("📁 Cargar Espacios")
        self.btn_asignar = QPushButton("⚡ Asignar Automáticamente")
        self.btn_exportar = QPushButton("💾 Exportar Resultados")

        self.btn_asignar.setEnabled(False)
        self.btn_exportar.setEnabled(False)

        btn_layout.addWidget(self.btn_monitores)
        btn_layout.addWidget(self.btn_espacios)
        btn_layout.addWidget(self.btn_asignar)
        btn_layout.addWidget(self.btn_exportar)

        layout.addLayout(btn_layout)

        # Barra de progreso
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setStyleSheet("""
            QProgressBar {
                border: 2px solid #DDD;
                border-radius: 5px;
                text-align: center;
                height: 25px;
            }
            QProgressBar::chunk {
                background-color: #0078D4;
            }
        """)
        layout.addWidget(self.progress)

        # Etiqueta de estado
        self.lbl_estado = QLabel("📋 Esperando archivos...")
        self.lbl_estado.setFont(QFont("Segoe UI", 10))
        self.lbl_estado.setStyleSheet("color: #666; padding: 5px;")
        layout.addWidget(self.lbl_estado)

        # Tabla de resultados
        self.table = QTableView()
        layout.addWidget(self.table, stretch=3)

        # Área de texto para reporte
        self.text_reporte = QTextEdit()
        self.text_reporte.setReadOnly(True)
        self.text_reporte.setPlaceholderText("El reporte de asignación aparecerá aquí...")
        layout.addWidget(self.text_reporte, stretch=2)

        self.setLayout(layout)

        # Conectar funciones
        self.btn_monitores.clicked.connect(self.cargar_monitores)
        self.btn_espacios.clicked.connect(self.cargar_espacios)
        self.btn_asignar.clicked.connect(self.iniciar_asignacion)
        self.btn_exportar.clicked.connect(self.exportar)

        # Variables de datos
        self.monitores = []
        self.df_espacios = pd.DataFrame()
        self.df_resultado = pd.DataFrame()
        self.monitores_asignados = []

    def cargar_monitores(self):
        ruta, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar archivo de monitores", "", 
            "Archivos Excel (*.xlsx *.xls)"
        )
        
        if ruta:
            try:
                self.monitores = cargar_monitores_desde_excel(ruta)
                
                df_preview = pd.DataFrame([{
                    'Nombre': m['nombre'],
                    'Min': m['min'],
                    'Max': m['max']
                } for m in self.monitores])
                
                self.table.setModel(PandasModel(df_preview))
                self.lbl_estado.setText(f"✅ {len(self.monitores)} monitores cargados")
                self.text_reporte.setPlainText(f"📂 Monitores cargados: {len(self.monitores)}")
                
                self.verificar_listo()
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error al cargar monitores:\n{str(e)}")

    def cargar_espacios(self):
        ruta, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar archivo de espacios", "", 
            "Archivos Excel (*.xlsx *.xls)"
        )
        
        if ruta:
            try:
                self.df_espacios = cargar_espacios_desde_excel(ruta)
                
                self.table.setModel(PandasModel(self.df_espacios.head(50)))
                self.lbl_estado.setText(f"✅ {len(self.df_espacios)} horarios cargados")
                self.text_reporte.setPlainText(
                    f"📂 Espacios cargados: {len(self.df_espacios)} horarios\n"
                    f"🏢 Salas: {self.df_espacios['SALA'].nunique()}\n"
                    f"⏱️  Total horas: {self.df_espacios['DURACION'].sum()}"
                )
                
                self.verificar_listo()
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error al cargar espacios:\n{str(e)}")

    def verificar_listo(self):
        if len(self.monitores) > 0 and len(self.df_espacios) > 0:
            self.btn_asignar.setEnabled(True)
            self.lbl_estado.setText("✅ Listo para asignar")

    def iniciar_asignacion(self):
        self.btn_asignar.setEnabled(False)
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)  # Indeterminado
        
        # Crear copia de monitores para el thread
        import copy
        monitores_copy = copy.deepcopy(self.monitores)
        
        self.thread = AsignacionThread(monitores_copy, self.df_espacios)
        self.thread.finished.connect(self.asignacion_completada)
        self.thread.error.connect(self.asignacion_error)
        self.thread.progress.connect(self.actualizar_progreso)
        self.thread.start()

    def actualizar_progreso(self, mensaje):
        self.lbl_estado.setText(mensaje)

    def asignacion_completada(self, df_resultado, monitores, reporte):
        self.df_resultado = df_resultado
        self.monitores_asignados = monitores
        
        self.table.setModel(PandasModel(df_resultado))
        self.text_reporte.setPlainText(reporte)
        
        self.progress.setVisible(False)
        self.btn_asignar.setEnabled(True)
        self.btn_exportar.setEnabled(True)
        
        self.lbl_estado.setText("✅ Asignación completada exitosamente")
        
        QMessageBox.information(
            self, 
            "Completado", 
            "✅ Asignación completada\n\nRevisa los resultados en la tabla y el reporte."
        )

    def asignacion_error(self, error):
        self.progress.setVisible(False)
        self.btn_asignar.setEnabled(True)
        
        QMessageBox.critical(self, "Error", f"Error en la asignación:\n{error}")
        self.lbl_estado.setText("❌ Error en la asignación")

    def exportar(self):
        if self.df_resultado.empty:
            QMessageBox.warning(self, "Advertencia", "No hay resultados para exportar")
            return
        
        ruta, _ = QFileDialog.getSaveFileName(
            self, "Guardar archivo", "Asignacion_Monitores.xlsx",
            "Archivos Excel (*.xlsx)"
        )
        
        if ruta:
            try:
                df_mon = pd.DataFrame([{
                    'Monitor': m['nombre'],
                    'Horas': m['horas'],
                    'Min': m['min'],
                    'Max': m['max'],
                    'Horarios': len(m['asignaciones']),
                    'Estado': '✅' if m['min'] <= m['horas'] <= m['max'] else '⚠️'
                } for m in self.monitores_asignados]).sort_values('Horas', ascending=False)
                
                with pd.ExcelWriter(ruta, engine='openpyxl') as writer:
                    self.df_resultado.to_excel(writer, sheet_name='Asignaciones', index=False)
                    df_mon.to_excel(writer, sheet_name='Resumen Monitores', index=False)
                
                QMessageBox.information(self, "Exportado", f"✅ Archivo guardado:\n{ruta}")
                self.lbl_estado.setText(f"✅ Exportado: {ruta}")
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error al exportar:\n{str(e)}")


# ========================================================
# EJECUTAR APLICACIÓN
# ========================================================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())