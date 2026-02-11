"""
Diálogo moderno de configuración del sistema
"""
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QSpinBox, QCheckBox, QFrame, QGridLayout, QScrollArea, QWidget
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from core.config import CONFIG, save_config


class ConfigDialog(QDialog):
    """Diálogo de configuración con diseño moderno"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("⚙️ Configuración del Sistema")
        self.setModal(True)
        self.setMinimumSize(700, 500)
        self.setStyleSheet("""
            QDialog {
                background: white;
            }
        """)
        
        self.setup_ui()
        self.load_config()
    
    def setup_ui(self):
        """Configura la interfaz"""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Header con gradiente
        self.create_header(main_layout)
        
        # Área de scroll para configuraciones
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("QScrollArea {background: white; border: none;}")
        
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_layout.setContentsMargins(32, 24, 32, 24)
        scroll_layout.setSpacing(24)
        
        # Secciones de configuración
        self.create_monitor_section(scroll_layout)
        self.create_assignment_section(scroll_layout)
        
        scroll_layout.addStretch()
        scroll.setWidget(scroll_content)
        main_layout.addWidget(scroll)
        
        # Footer con botones
        self.create_footer(main_layout)
    
    def create_header(self, parent_layout):
        """Crea el header con gradiente"""
        header = QFrame()
        header.setFixedHeight(100)
        header.setStyleSheet("""
            QFrame {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:1,
                    stop:0 #3B82F6,
                    stop:1 #8B5CF6
                );
                border-radius: 0px;
            }
        """)
        
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(32, 20, 32, 20)
        
        # Icono y título
        icon_label = QLabel("⚙️")
        icon_label.setFont(QFont("Inter", 32))
        icon_label.setStyleSheet("background: transparent; color: white;")
        header_layout.addWidget(icon_label)
        
        title_layout = QVBoxLayout()
        title_layout.setSpacing(4)
        
        title = QLabel("Configuración del Sistema")
        title.setFont(QFont("Inter", 20, QFont.Bold))
        title.setStyleSheet("color: white; background: transparent;")
        title_layout.addWidget(title)
        
        subtitle = QLabel("Personaliza el comportamiento de las asignaciones")
        subtitle.setFont(QFont("Inter", 11))
        subtitle.setStyleSheet("color: rgba(255, 255, 255, 0.9); background: transparent;")
        title_layout.addWidget(subtitle)
        
        header_layout.addLayout(title_layout)
        header_layout.addStretch()
        
        parent_layout.addWidget(header)
    
    def create_monitor_section(self, parent_layout):
        """Sección de configuración de monitores"""
        section = self.create_section_card(
            "👥 Configuración de Monitores",
            "Define las horas por defecto para nuevos monitores"
        )
        section_layout = QGridLayout()
        section_layout.setSpacing(16)
        section_layout.setContentsMargins(20, 16, 20, 20)
        
        # Horas mínimas
        min_container = self.create_input_container(
            "Horas Mínimas",
            "Mínimo de horas que debe trabajar cada monitor"
        )
        self.spin_horas_min = self.create_spinbox(0, 100, " hrs")
        min_container.layout().addWidget(self.spin_horas_min)
        section_layout.addWidget(min_container, 0, 0)
        
        # Horas máximas
        max_container = self.create_input_container(
            "Horas Máximas",
            "Máximo de horas que puede trabajar cada monitor"
        )
        self.spin_horas_max = self.create_spinbox(0, 100, " hrs")
        max_container.layout().addWidget(self.spin_horas_max)
        section_layout.addWidget(max_container, 0, 1)
        
        section.layout().addLayout(section_layout)
        parent_layout.addWidget(section)
    
    def create_assignment_section(self, parent_layout):
        """Sección de configuración de asignación"""
        section = self.create_section_card(
            "🎯 Opciones de Asignación",
            "Configura cómo se asignan los monitores a los horarios"
        )
        
        options_layout = QVBoxLayout()
        options_layout.setSpacing(12)
        options_layout.setContentsMargins(20, 16, 20, 20)
        
        # Crear checkboxes directamente sin contenedor
        self.chk_usar_prioridad = QCheckBox("🎖️ Usar sistema de prioridades")
        self.chk_usar_prioridad.setStyleSheet("""
            QCheckBox {
                color: #111827;
                background: #F9FAFB;
                padding: 12px;
                border-radius: 8px;
                font-size: 11pt;
                font-weight: 600;
            }
            QCheckBox:hover {
                background: #F3F4F6;
            }
            QCheckBox::indicator {
                width: 20px;
                height: 20px;
                border-radius: 4px;
                border: 2px solid #D1D5DB;
            }
            QCheckBox::indicator:checked {
                background: #3B82F6;
                border-color: #3B82F6;
            }
        """)
        options_layout.addWidget(self.chk_usar_prioridad)
        
        self.chk_balancear = QCheckBox("⚖️ Balancear carga de trabajo")
        self.chk_balancear.setStyleSheet("""
            QCheckBox {
                color: #111827;
                background: #F9FAFB;
                padding: 12px;
                border-radius: 8px;
                font-size: 11pt;
                font-weight: 600;
            }
            QCheckBox:hover {
                background: #F3F4F6;
            }
            QCheckBox::indicator {
                width: 20px;
                height: 20px;
                border-radius: 4px;
                border: 2px solid #D1D5DB;
            }
            QCheckBox::indicator:checked {
                background: #3B82F6;
                border-color: #3B82F6;
            }
        """)
        options_layout.addWidget(self.chk_balancear)
        
        self.chk_priorizar_min = QCheckBox("🎯 Priorizar cumplir mínimos")
        self.chk_priorizar_min.setStyleSheet("""
            QCheckBox {
                color: #111827;
                background: #F9FAFB;
                padding: 12px;
                border-radius: 8px;
                font-size: 11pt;
                font-weight: 600;
            }
            QCheckBox:hover {
                background: #F3F4F6;
            }
            QCheckBox::indicator {
                width: 20px;
                height: 20px;
                border-radius: 4px;
                border: 2px solid #D1D5DB;
            }
            QCheckBox::indicator:checked {
                background: #3B82F6;
                border-color: #3B82F6;
            }
        """)
        options_layout.addWidget(self.chk_priorizar_min)
        
        section.layout().addLayout(options_layout)
        parent_layout.addWidget(section)
    
    def create_footer(self, parent_layout):
        """Crea el footer con botones"""
        footer = QFrame()
        footer.setStyleSheet("""
            QFrame {
                background: #F9FAFB;
                border-top: 1px solid #E5E7EB;
            }
        """)
        footer.setFixedHeight(80)
        
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(32, 16, 32, 16)
        
        # Info de guardado
        info_label = QLabel("💡 Los cambios se guardarán automáticamente")
        info_label.setStyleSheet("color: #6B7280; background: transparent;")
        info_label.setFont(QFont("Inter", 10))
        footer_layout.addWidget(info_label)
        
        footer_layout.addStretch()
        
        # Botón cancelar
        btn_cancel = QPushButton("✕ Cancelar")
        btn_cancel.setFixedHeight(44)
        btn_cancel.setFixedWidth(120)
        btn_cancel.clicked.connect(self.reject)
        btn_cancel.setStyleSheet("""
            QPushButton {
                background: white;
                color: #6B7280;
                border: 1px solid #D1D5DB;
                border-radius: 8px;
                font-size: 13px;
                font-weight: 600;
            }
            QPushButton:hover {
                background: #F9FAFB;
                border-color: #9CA3AF;
            }
        """)
        footer_layout.addWidget(btn_cancel)
        
        # Botón guardar
        btn_save = QPushButton("💾 Guardar Cambios")
        btn_save.setFixedHeight(44)
        btn_save.setFixedWidth(160)
        btn_save.clicked.connect(self.save_and_close)
        btn_save.setStyleSheet("""
            QPushButton {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 #3B82F6,
                    stop:1 #2563EB
                );
                color: white;
                border: none;
                border-radius: 8px;
                font-size: 13px;
                font-weight: 600;
            }
            QPushButton:hover {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 #2563EB,
                    stop:1 #1D4ED8
                );
            }
        """)
        footer_layout.addWidget(btn_save)
        
        parent_layout.addWidget(footer)
    
    # Funciones helper para crear elementos
    
    def create_section_card(self, title, description):
        """Crea una tarjeta de sección"""
        card = QFrame()
        card.setStyleSheet("""
            QFrame {
                background: white;
                border: 1px solid #E5E7EB;
                border-radius: 12px;
            }
        """)
        
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(0, 0, 0, 0)
        card_layout.setSpacing(0)
        
        # Header de la sección
        header = QFrame()
        header.setStyleSheet("""
            QFrame {
                background: #F9FAFB;
                border-radius: 12px 12px 0 0;
                border-bottom: 1px solid #E5E7EB;
            }
        """)
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(20, 16, 20, 16)
        header_layout.setSpacing(4)
        
        title_label = QLabel(title)
        title_label.setFont(QFont("Inter", 13, QFont.Bold))
        title_label.setStyleSheet("color: #111827; background: transparent; border: none;")
        header_layout.addWidget(title_label)
        
        desc_label = QLabel(description)
        desc_label.setFont(QFont("Inter", 10))
        desc_label.setStyleSheet("color: #6B7280; background: transparent; border: none;")
        header_layout.addWidget(desc_label)
        
        card_layout.addWidget(header)
        
        return card
    
    def create_input_container(self, label, hint):
        """Crea un contenedor para input"""
        container = QFrame()
        container.setStyleSheet("QFrame {background: transparent; border: none;}")
        layout = QVBoxLayout(container)
        layout.setSpacing(6)
        layout.setContentsMargins(0, 0, 0, 0)
        
        label_widget = QLabel(label)
        label_widget.setFont(QFont("Inter", 11, QFont.Bold))
        label_widget.setStyleSheet("color: #374151; background: transparent;")
        layout.addWidget(label_widget)
        
        hint_widget = QLabel(hint)
        hint_widget.setFont(QFont("Inter", 9))
        hint_widget.setStyleSheet("color: #9CA3AF; background: transparent;")
        hint_widget.setWordWrap(True)
        layout.addWidget(hint_widget)
        
        return container
    
    def create_spinbox(self, min_val, max_val, suffix):
        """Crea un spinbox estilizado"""
        spinbox = QSpinBox()
        spinbox.setRange(min_val, max_val)
        spinbox.setSuffix(suffix)
        spinbox.setFixedHeight(40)
        spinbox.setStyleSheet("""
            QSpinBox {
                background: white;
                border: 2px solid #E5E7EB;
                border-radius: 8px;
                padding: 8px 12px;
                color: #111827;
                font-size: 13px;
                font-weight: 600;
            }
            QSpinBox:focus {
                border: 2px solid #3B82F6;
            }
            QSpinBox::up-button, QSpinBox::down-button {
                width: 20px;
                border-radius: 4px;
            }
            QSpinBox::up-button:hover, QSpinBox::down-button:hover {
                background: #F3F4F6;
            }
        """)
        return spinbox
    
    def load_config(self):
        """Carga la configuración actual"""
        self.spin_horas_min.setValue(CONFIG["monitores"]["horas_min_default"])
        self.spin_horas_max.setValue(CONFIG["monitores"]["horas_max_default"])
        self.chk_usar_prioridad.setChecked(CONFIG["asignacion"]["usar_prioridad"])
        self.chk_balancear.setChecked(CONFIG["asignacion"]["balancear_carga"])
        self.chk_priorizar_min.setChecked(CONFIG["asignacion"]["priorizar_minimo"])
    
    def save_and_close(self):
        """Guarda y cierra el diálogo"""
        CONFIG["monitores"]["horas_min_default"] = self.spin_horas_min.value()
        CONFIG["monitores"]["horas_max_default"] = self.spin_horas_max.value()
        CONFIG["asignacion"]["usar_prioridad"] = self.chk_usar_prioridad.isChecked()
        CONFIG["asignacion"]["balancear_carga"] = self.chk_balancear.isChecked()
        CONFIG["asignacion"]["priorizar_minimo"] = self.chk_priorizar_min.isChecked()
        
        if save_config(CONFIG):
            self.accept()
        else:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.warning(self, "Error", "No se pudo guardar la configuración")