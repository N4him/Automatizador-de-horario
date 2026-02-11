"""
Tarjetas de estadísticas
"""
from PySide6.QtWidgets import QFrame, QVBoxLayout, QHBoxLayout, QLabel
from PySide6.QtGui import QFont


class StatCard(QFrame):
    """Tarjeta para mostrar estadísticas"""
    
    def __init__(self, icon, title, value, color, parent=None):
        super().__init__(parent)
        self.color = color
        self._setup_ui(icon, title, value)
    
    def _setup_ui(self, icon, title, value):
        """Configura la interfaz de la tarjeta"""
        self.setStyleSheet(
            f"QFrame {{"
            f"    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, "
            f"                                 stop:0 {self.color}, "
            f"                                 stop:1 {self._darken_color(self.color)});"
            f"    border-radius: 14px;"
            f"}}"
        )
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 14, 18, 14)
        
        # Header
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
        
        # Value
        self.value_label = QLabel(value)
        self.value_label.setFont(QFont("Inter", 24, QFont.Bold))
        self.value_label.setStyleSheet("color: white; background: transparent;")
        layout.addWidget(self.value_label)
    
    def update_value(self, value):
        """Actualiza el valor mostrado"""
        self.value_label.setText(str(value))
    
    @staticmethod
    def _darken_color(hex_color):
        """Oscurece un color hex"""
        hex_color = hex_color.lstrip('#')
        r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        factor = 0.8
        r, g, b = int(r * factor), int(g * factor), int(b * factor)
        return f'#{r:02x}{g:02x}{b:02x}'