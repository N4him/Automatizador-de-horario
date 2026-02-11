"""
Diálogo de descarga de resultados
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
    QPushButton, QFrame, QMessageBox, QFileDialog
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
import pandas as pd

from export.excel_writer import exportar_resultados
from core.config import CONFIG


class DownloadDialog(QDialog):
    """Diálogo para descargar resultados de asignación"""
    
    def __init__(self, df_resultado, monitores_asignados, parent=None):
        super().__init__(parent)
        self.df_resultado = df_resultado
        self.monitores_asignados = monitores_asignados
        self.archivo_guardado = None
        
        self._setup_ui()
    
    def _setup_ui(self):
        """Configura la interfaz del diálogo"""
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
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #7C8FF5, stop:1 #6B64F8);
            }
            QPushButton#btnSecondary {background: #F3F4F6; color: #374151;}
            QPushButton#btnSecondary:hover {background: #E5E7EB;}
        """)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(32, 32, 32, 32)
        layout.setSpacing(20)
        
        # Icono
        icon = QLabel("✅")
        icon.setFont(QFont("Inter", 64))
        icon.setAlignment(Qt.AlignCenter)
        icon.setStyleSheet("color: #10B981;")
        layout.addWidget(icon)
        
        # Título
        title = QLabel("¡Asignación Completada!")
        title.setFont(QFont("Inter", 20, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        # Estadísticas
        exitosos = len([a for a in self.df_resultado.to_dict('records') 
                       if a.get("ESTADO") == "✅"])
        total = len(self.df_resultado)
        porcentaje = f"{exitosos*100/total:.1f}%" if total > 0 else "0%"
        
        stats_frame = QFrame()
        stats_frame.setStyleSheet(
            "QFrame {background: #F9FAFB; border-radius: 12px; padding: 16px;}"
        )
        stats_layout = QVBoxLayout(stats_frame)
        
        stats_text = (
            f"<div style='text-align: center;'>"
            f"<p style='font-size: 14px; color: #6B7280;'>Se asignaron correctamente</p>"
            f"<p style='font-size: 32px; font-weight: bold; color: #667EEA;'>"
            f"{exitosos}/{total}</p>"
            f"<p style='font-size: 14px; color: #6B7280;'>horarios ({porcentaje})</p>"
            f"</div>"
        )
        stats_label = QLabel(stats_text)
        stats_label.setAlignment(Qt.AlignCenter)
        stats_layout.addWidget(stats_label)
        layout.addWidget(stats_frame)
        
        # Mensaje
        message = QLabel("¿Deseas descargar el archivo con horarios visuales ahora?")
        message.setFont(QFont("Inter", 13))
        message.setAlignment(Qt.AlignCenter)
        message.setStyleSheet("color: #6B7280;")
        layout.addWidget(message)
        
        # Botones
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
        """Maneja la descarga del archivo"""
        ruta, _ = QFileDialog.getSaveFileName(
            self, 
            "Guardar archivo de resultados", 
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
                
                self.archivo_guardado = ruta
                
                success_msg = QMessageBox(self)
                success_msg.setIcon(QMessageBox.Information)
                success_msg.setWindowTitle("Descarga Exitosa")
                success_msg.setText("✅ Archivo guardado exitosamente")
                success_msg.setInformativeText(
                    f"📁 {ruta}\n\n"
                    f"📄 4 hojas generadas:\n"
                    f"• Horarios Salas (visual)\n"
                    f"• Horarios Monitores (individual)\n"
                    f"• Lista Asignaciones\n"
                    f"• Resumen Monitores"
                )
                success_msg.exec()
                self.accept()
                
            except Exception as e:
                QMessageBox.critical(
                    self, 
                    "Error", 
                    f"Error al guardar el archivo:\n{str(e)}"
                )