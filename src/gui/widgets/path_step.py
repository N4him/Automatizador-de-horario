"""
Widget de paso en el roadmap
"""
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QSizePolicy
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont


class PathStep(QFrame):
    """Paso visual en el proceso de asignación"""
    
    def __init__(self, number, title, icon, status="pending", parent=None):
        super().__init__(parent)
        self.number = number
        self.title = title
        self.icon = icon
        self.status = status
        self._setup_ui()
    
    def _setup_ui(self):
        """Configura la interfaz del paso"""
        self.setMinimumHeight(90)
        self.setMaximumHeight(90)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        
        self._update_style()
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
        layout.setAlignment(Qt.AlignCenter)
        
        # Número
        self.number_label = QLabel(self.number)
        self.number_label.setAlignment(Qt.AlignCenter)
        self.number_label.setFixedSize(36, 36)
        self.number_label.setFont(QFont("Inter", 16, QFont.Bold))
        layout.addWidget(self.number_label)
        
        # Icono
        self.icon_label = QLabel(self.icon)
        self.icon_label.setFont(QFont("Inter", 32))
        self.icon_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.icon_label)
        
        # Título
        self.title_label = QLabel(self.title)
        self.title_label.setFont(QFont("Inter", 12, QFont.Bold))
        self.title_label.setAlignment(Qt.AlignCenter)
        self.title_label.setWordWrap(True)
        layout.addWidget(self.title_label)
        
        self._update_style()
    
    def update_status(self, status):
        """Actualiza el estado del paso"""
        self.status = status
        self._update_style()
    
    def _update_style(self):
        """Actualiza los estilos según el estado"""
        if self.status == "pending":
            bg_gradient = "stop:0 #F9FAFB, stop:1 #F3F4F6"
            border_color = "#E5E7EB"
            text_color = "#9CA3AF"
            number_bg = "#E5E7EB"
            number_color = "#6B7280"
        elif self.status == "success":
            bg_gradient = "stop:0 #D1FAE5, stop:1 #A7F3D0"
            border_color = "#10B981"
            text_color = "#065F46"
            number_bg = "#10B981"
            number_color = "white"
        else:  # error
            bg_gradient = "stop:0 #FEE2E2, stop:1 #FECACA"
            border_color = "#EF4444"
            text_color = "#991B1B"
            number_bg = "#EF4444"
            number_color = "white"
        
        self.setStyleSheet(
            f"QFrame {{"
            f"    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, {bg_gradient});"
            f"    border: 2px solid {border_color};"
            f"    border-radius: 12px;"
            f"}}"
        )
        
        if hasattr(self, 'number_label'):
            self.number_label.setStyleSheet(
                f"QLabel {{"
                f"    background: {number_bg};"
                f"    color: {number_color};"
                f"    border-radius: 18px;"
                f"    border: none;"
                f"}}"
            )
            self.icon_label.setStyleSheet(
                f"color: {text_color}; background: transparent; border: none;"
            )
            self.title_label.setStyleSheet(
                f"color: {text_color}; background: transparent; border: none;"
            )


class PathConnector(QFrame):
    """Conector entre pasos"""
    
    def __init__(self, status="pending", parent=None):
        super().__init__(parent)
        self.status = status
        self.setFixedSize(40, 90)
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self._update_style()
    
    def update_status(self, status):
        """Actualiza el estado del conector"""
        self.status = status
        self._update_style()
    
    def _update_style(self):
        """Actualiza el estilo según el estado"""
        if self.status == "pending":
            bg_color = "#E5E7EB"
        elif self.status == "success":
            bg_color = "#10B981"
        else:
            bg_color = "#EF4444"
        
        self.setStyleSheet(f"QFrame {{background: {bg_color}; border: none;}}")