"""
Ventana principal de la aplicación
"""
import sys
import copy
import pandas as pd
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTableView, QProgressBar, QTextEdit, QCheckBox, QFileDialog,
    QMessageBox, QTabWidget, QDialog, QFrame
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

# IMPORTS CORREGIDOS - Agregar estos imports relativos
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.config import CONFIG
from core.loader import cargar_monitores_desde_excel, cargar_espacios_desde_excel
from export.excel_writer import exportar_resultados
from utils.threads import AsignacionThread
from gui.models.pandas_model import PandasModel
from gui.widgets.modern_card import ModernCard
from gui.widgets.stat_card import StatCard
from gui.widgets.path_step import PathStep, PathConnector
from gui.dialogs.download_dialog import DownloadDialog
from gui.styles.app_styles import MAIN_STYLESHEET


class MainWindow(QWidget):
    """Ventana principal del sistema de asignación"""
    
    def __init__(self):
        super().__init__()
        
        # Datos
        self.monitores = []
        self.df_espacios = pd.DataFrame()
        self.df_resultado = pd.DataFrame()
        self.monitores_asignados = []
        
        self._setup_ui()
    
    def _setup_ui(self):
        """Configura la interfaz principal"""
        self.setWindowTitle("Sistema de Asignación de Monitores")
        self.showMaximized()
        self.setStyleSheet(MAIN_STYLESHEET)
        
        main_layout = QVBoxLayout()
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(24, 24, 24, 24)
        
        # Stats cards
        self._create_stats_section(main_layout)
        
        # Roadmap
        self._create_roadmap_section(main_layout)
        
        # Progress bar
        self._create_progress_section(main_layout)
        
        # Status label
        self._create_status_section(main_layout)
        
        # Content area (tables + report)
        self._create_content_section(main_layout)
        
        self.setLayout(main_layout)
    
    def _create_stats_section(self, parent_layout):
        """Crea la sección de estadísticas"""
        stats_layout = QHBoxLayout()
        stats_layout.setSpacing(16)
        
        self.card_monitores = StatCard("👥", "Monitores", "0", "#8B5CF6")
        self.card_espacios = StatCard("📅", "Horarios", "0", "#3B82F6")
        self.card_asignados = StatCard("✅", "Asignados", "0%", "#10B981")
        
        stats_layout.addWidget(self.card_monitores)
        stats_layout.addWidget(self.card_espacios)
        stats_layout.addWidget(self.card_asignados)
        
        parent_layout.addLayout(stats_layout)
    
    def _create_roadmap_section(self, parent_layout):
        """Crea la sección del roadmap"""
        roadmap_card = ModernCard()
        roadmap_layout = QVBoxLayout(roadmap_card)
        roadmap_layout.setContentsMargins(20, 16, 20, 16)
        roadmap_layout.setSpacing(12)
        
        # Header
        header_layout = QHBoxLayout()
        
        roadmap_title = QLabel("🗺️ Proceso de Asignación")
        roadmap_title.setFont(QFont("Inter", 13, QFont.Bold))
        roadmap_title.setStyleSheet("color: #111827; background: transparent;")
        header_layout.addWidget(roadmap_title)
        header_layout.addStretch()
        
        # Botón recarga
        self.btn_reload = QPushButton("🔄")
        self.btn_reload.setFixedSize(50, 50)
        self.btn_reload.setToolTip("Recargar - Limpiar todos los datos cargados")
        self.btn_reload.setStyleSheet("""
            QPushButton {
                background: #F3F4F6; 
                color: #991B1B; 
                font-size: 22px; 
                padding: 2px; 
                border-radius: 10px;
            }
            QPushButton:hover {background: #E5E7EB;}
        """)
        header_layout.addWidget(self.btn_reload)
        self.btn_reload.clicked.connect(self.recargar_sistema)
        
        # Botón configuración
        self.btn_config = QPushButton("⚙️")
        self.btn_config.setFixedSize(50, 50)
        self.btn_config.setToolTip("Configuración del sistema")
        self.btn_config.setStyleSheet("""
            QPushButton {
                background: #F3F4F6; 
                color: #374151; 
                font-size: 22px; 
                padding: 2px; 
                border-radius: 10px;
            }
            QPushButton:hover {background: #E5E7EB;}
        """)
        header_layout.addWidget(self.btn_config)
        self.btn_config.clicked.connect(self.open_config_dialog)
        
        roadmap_layout.addLayout(header_layout)
        
        # Steps container
        steps_container = QHBoxLayout()
        steps_container.setSpacing(0)
        steps_container.setContentsMargins(0, 0, 0, 0)
        
        # Step 1
        self.step1_frame = PathStep("1", "Cargar Monitores", "👥", "pending")
        self.step1_frame.mousePressEvent = lambda e: self.cargar_monitores()
        self.step1_frame.setCursor(Qt.PointingHandCursor)
        steps_container.addWidget(self.step1_frame)
        
        self.connector1 = PathConnector("pending")
        steps_container.addWidget(self.connector1)
        
        # Step 2
        self.step2_frame = PathStep("2", "Cargar Espacios", "📅", "pending")
        self.step2_frame.mousePressEvent = lambda e: self.cargar_espacios()
        self.step2_frame.setCursor(Qt.PointingHandCursor)
        steps_container.addWidget(self.step2_frame)
        
        self.connector2 = PathConnector("pending")
        steps_container.addWidget(self.connector2)
        
        # Step 3
        self.step3_frame = PathStep("3", "Asignar Monitores", "⚡", "pending")
        self.step3_frame.mousePressEvent = lambda e: self._handle_step3_click()
        self.step3_frame.setCursor(Qt.PointingHandCursor)
        steps_container.addWidget(self.step3_frame)
        
        self.connector3 = PathConnector("pending")
        steps_container.addWidget(self.connector3)
        
        # Step 4
        self.step4_frame = PathStep("4", "Exportar Resultados", "💾", "pending")
        self.step4_frame.mousePressEvent = lambda e: self._handle_step4_click()
        self.step4_frame.setCursor(Qt.PointingHandCursor)
        steps_container.addWidget(self.step4_frame)
        
        roadmap_layout.addLayout(steps_container)
        parent_layout.addWidget(roadmap_card)
        
        # Config overlay
    
    
    
    def _create_progress_section(self, parent_layout):
        """Crea la sección de progreso"""
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setTextVisible(False)
        self.progress.setMaximumHeight(6)
        parent_layout.addWidget(self.progress)
    
    def _create_status_section(self, parent_layout):
        """Crea la sección de estado"""
        self.lbl_estado = QLabel("📋 Esperando archivos...")
        self.lbl_estado.setFont(QFont("Inter", 12))
        self.lbl_estado.setStyleSheet(
            "color: #6B7280; background: white; padding: 10px 16px; "
            "border-radius: 8px; border: 1px solid #E5E7EB;"
        )
        parent_layout.addWidget(self.lbl_estado)
    
    def _create_content_section(self, parent_layout):
        """Crea la sección de contenido (tablas + reporte)"""
        content_layout = QHBoxLayout()
        content_layout.setSpacing(12)
        
        # Tables
        table_card = ModernCard()
        table_layout = QVBoxLayout(table_card)
        table_layout.setContentsMargins(0, 0, 0, 0)
        
        table_header = QLabel("📊 Visualización de Datos")
        table_header.setFont(QFont("Inter", 13, QFont.Bold))
        table_header.setStyleSheet(
            "color: #1F2937; padding: 14px 16px; background: #F9FAFB; "
            "border-radius: 12px 12px 0 0;"
        )
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
        
        # Report
        report_card = ModernCard()
        report_layout = QVBoxLayout(report_card)
        report_layout.setContentsMargins(0, 0, 0, 0)
        
        report_header = QLabel("📝 Reporte")
        report_header.setFont(QFont("Inter", 13, QFont.Bold))
        report_header.setStyleSheet(
            "color: #1F2937; padding: 14px 16px; background: #F9FAFB; "
            "border-radius: 12px 12px 0 0;"
        )
        report_layout.addWidget(report_header)
        
        self.text_reporte = QTextEdit()
        self.text_reporte.setReadOnly(True)
        self.text_reporte.setPlaceholderText("El reporte detallado aparecerá aquí...")
        report_layout.addWidget(self.text_reporte)
        
        content_layout.addWidget(report_card, stretch=4)
        parent_layout.addLayout(content_layout, stretch=1)
    
    def _handle_step3_click(self):
        """Maneja el click en el paso 3"""
        if (self.step1_frame.status == "success" and 
            self.step2_frame.status == "success"):
            self.iniciar_asignacion()
    
    def _handle_step4_click(self):
        """Maneja el click en el paso 4"""
        if not self.df_resultado.empty:
            self.exportar()
    
    def update_connectors(self):
        """Actualiza el estado de los conectores"""
        if self.step1_frame.status == "success":
            self.connector1.update_status("success")
        else:
            self.connector1.update_status("pending")
        
        if self.step2_frame.status == "success":
            self.connector2.update_status("success")
        else:
            self.connector2.update_status("pending")
        
        if self.step3_frame.status == "success":
            self.connector3.update_status("success")
        else:
            self.connector3.update_status("pending")
    
    def recargar_sistema(self):
        """Reinicia todo el sistema a su estado inicial"""
        # Confirmar acción
        respuesta = QMessageBox.question(
            self,
            "Confirmar Recarga",
            "¿Estás seguro de que deseas reiniciar el sistema?\n\n"
            "Esto borrará:\n"
            "• Todos los archivos cargados\n"
            "• Las asignaciones realizadas\n"
            "• Los resultados generados",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if respuesta == QMessageBox.Yes:
            # Limpiar datos
            self.monitores = []
            self.df_espacios = pd.DataFrame()
            self.df_resultado = pd.DataFrame()
            self.monitores_asignados = []
            
            # Limpiar tablas
            self.table_monitores.setModel(None)
            self.table_espacios.setModel(None)
            self.table_asignaciones.setModel(None)
            self.table_resumen.setModel(None)
            
            # Resetear stats
            self.card_monitores.update_value("0")
            self.card_espacios.update_value("0")
            self.card_asignados.update_value("0%")
            
            # Limpiar reporte
            self.text_reporte.clear()
            
            # Resetear estado
            self.lbl_estado.setText("📋 Esperando archivos...")
            
            # Resetear pasos
            self.step1_frame.update_status("pending")
            self.step2_frame.update_status("pending")
            self.step3_frame.update_status("pending")
            self.step4_frame.update_status("pending")
            self.update_connectors()
            
            # Ocultar progress bar
            self.progress.setVisible(False)
            
            # Volver al primer tab
            self.tabs.setCurrentIndex(0)
            
            QMessageBox.information(
                self,
                "Sistema Reiniciado",
                "✅ El sistema ha sido reiniciado correctamente.\n"
                "Puedes comenzar de nuevo cargando los archivos."
            )
    
    def cargar_monitores(self):
        """Carga el archivo de monitores"""
        ruta, _ = QFileDialog.getOpenFileName(
            self, 
            "Seleccionar archivo de monitores", 
            "", 
            "Hojas de Cálculo (*.xlsx *.xls *.ods);;"
            "Excel (*.xlsx *.xls);;LibreOffice (*.ods)"
        )
        
        if ruta:
            try:
                self.monitores = cargar_monitores_desde_excel(ruta)
                
                df_preview = pd.DataFrame([{
                    'Nombre': m['nombre'], 
                    'Prioridad': f"⭐{m['prioridad']}", 
                    'Min': m['min'], 
                    'Max': m['max']
                } for m in self.monitores])
                
                self.table_monitores.setModel(PandasModel(df_preview))
                self.tabs.setCurrentIndex(0)
                
                self.card_monitores.update_value(str(len(self.monitores)))
                self.lbl_estado.setText(
                    f"✅ {len(self.monitores)} monitores cargados exitosamente"
                )
                self.text_reporte.setPlainText(
                    f"📂 Monitores cargados: {len(self.monitores)}"
                )
                
                self.verificar_listo()
                self.step1_frame.update_status("success")
                self.update_connectors()
                
            except Exception as e:
                self.step1_frame.update_status("error")
                self.update_connectors()
                QMessageBox.critical(
                    self, 
                    "Error", 
                    f"Error al cargar monitores:\n{str(e)}"
                )
    
    def cargar_espacios(self):
        """Carga el archivo de espacios"""
        ruta, _ = QFileDialog.getOpenFileName(
            self, 
            "Seleccionar archivo de espacios", 
            "", 
            "Hojas de Cálculo (*.xlsx *.xls *.ods);;"
            "Excel (*.xlsx *.xls);;LibreOffice (*.ods)"
        )
        
        if ruta:
            try:
                self.df_espacios = cargar_espacios_desde_excel(ruta)
                
                self.table_espacios.setModel(
                    PandasModel(self.df_espacios.head(100))
                )
                self.tabs.setCurrentIndex(1)
                
                self.card_espacios.update_value(str(len(self.df_espacios)))
                self.lbl_estado.setText(
                    f"✅ {len(self.df_espacios)} horarios cargados exitosamente"
                )
                
                self.verificar_listo()
                self.step2_frame.update_status("success")
                self.update_connectors()
                
            except Exception as e:
                self.step2_frame.update_status("error")
                self.update_connectors()
                QMessageBox.critical(
                    self, 
                    "Error", 
                    f"Error al cargar espacios:\n{str(e)}"
                )
    
    def verificar_listo(self):
        """Verifica si el sistema está listo para asignar"""
        if len(self.monitores) > 0 and len(self.df_espacios) > 0:
            self.lbl_estado.setText(
                "✅ Todo listo! Haz click en 'Asignar' para comenzar"
            )
    def iniciar_asignacion(self):
        """Inicia el proceso de asignación"""
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)
        
        # Actualizar configuración

        
        # Copiar monitores para no modificar originales
        monitores_copy = copy.deepcopy(self.monitores)
        
        # Crear y lanzar thread
        self.thread = AsignacionThread(monitores_copy, self.df_espacios)
        self.thread.finished.connect(self.asignacion_completada)
        self.thread.error.connect(self.asignacion_error)
        self.thread.progress.connect(self.actualizar_progreso)
        self.thread.start()
    


    
    def actualizar_progreso(self, mensaje):
        """Actualiza el mensaje de progreso"""
        self.lbl_estado.setText(mensaje)
    
    def asignacion_completada(self, df_resultado, monitores, reporte):
        """Maneja la finalización de la asignación"""
        self.df_resultado = df_resultado
        self.monitores_asignados = monitores
        
        # Actualizar tabla de asignaciones
        self.table_asignaciones.setModel(PandasModel(df_resultado))
        self.tabs.setCurrentIndex(2)
        
        # Actualizar resumen de monitores
        df_mon = pd.DataFrame([{
            'Monitor': m['nombre'], 
            'Prioridad': m['prioridad'], 
            'Horas': m['horas'], 
            'Min': m['min'], 
            'Max': m['max'], 
            'Horarios': len(m['asignaciones']), 
            'Estado': '✅' if m['min'] <= m['horas'] <= m['max'] else '⚠️'
        } for m in monitores if m['horas'] > 0]).sort_values('Horas', ascending=False)
        
        self.table_resumen.setModel(PandasModel(df_mon))
        
        # Actualizar reporte
        self.text_reporte.setPlainText(reporte)
        
        # Actualizar stats
        exitosos = len([a for a in df_resultado.to_dict('records') 
                       if a.get("ESTADO") == "✅"])
        total = len(df_resultado)
        self.card_asignados.update_value(f"{exitosos}/{total}")
        
        # Ocultar progress
        self.progress.setVisible(False)
        self.lbl_estado.setText("✅ Asignación completada exitosamente!")
        
        # Actualizar step 3
        if exitosos > 0:
            self.step3_frame.update_status("success")
        else:
            self.step3_frame.update_status("error")
        self.update_connectors()
        
        # Mostrar diálogo de descarga
        dialog = DownloadDialog(df_resultado, monitores, self)
        resultado = dialog.exec()
        
        if resultado == QDialog.Accepted and dialog.archivo_guardado:
            self.lbl_estado.setText(
                f"✅ Archivo exportado: {dialog.archivo_guardado}"
            )
            self.step4_frame.update_status("success")
            self.update_connectors()
    
    def asignacion_error(self, error):
        """Maneja errores en la asignación"""
        self.progress.setVisible(False)
        self.step3_frame.update_status("error")
        self.update_connectors()
        
        QMessageBox.critical(
            self, 
            "Error", 
            f"Error en la asignación:\n{error}"
        )
        self.lbl_estado.setText("❌ Error en la asignación")
    
    def exportar(self):
        """Exporta los resultados a Excel"""
        if self.df_resultado.empty:
            QMessageBox.warning(
                self, 
                "Advertencia", 
                "No hay resultados para exportar"
            )
            return
        
        ruta, _ = QFileDialog.getSaveFileName(
            self, 
            "Guardar archivo", 
            "Asignacion_Monitores.xlsx", 
            "Archivos Excel (*.xlsx)"
        )
        
        if ruta:
            try:
                cfg_esp = CONFIG["espacios"]
                exportar_resultados(
                    ruta, 
                    self.df_resultado, 
                    self.monitores_asignados, 
                    cfg_esp
                )
                
                QMessageBox.information(
                    self, 
                    "Exportado", 
                    f"✅ Archivo guardado exitosamente:\n{ruta}\n\n"
                    f"📄 4 hojas generadas:\n"
                    f"• Horarios Salas (visual)\n"
                    f"• Horarios Monitores (individual)\n"
                    f"• Lista Asignaciones\n"
                    f"• Resumen Monitores"
                )
                
                self.lbl_estado.setText(f"✅ Exportado: {ruta}")
                self.step4_frame.update_status("success")
                self.update_connectors()
                
            except Exception as e:
                self.step4_frame.update_status("error")
                self.update_connectors()
                QMessageBox.critical(
                    self, 
                    "Error", 
                    f"Error al exportar:\n{str(e)}"
                )
    def open_config_dialog(self):
            """Abre el diálogo de configuración moderno"""
            from gui.dialogs.config_dialog import ConfigDialog
            dialog = ConfigDialog(self)
            if dialog.exec():
                self.lbl_estado.setText("✅ Configuración guardada correctamente")