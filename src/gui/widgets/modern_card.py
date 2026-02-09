"""
Widget de tarjeta moderna
"""
from PySide6.QtWidgets import QFrame


class ModernCard(QFrame):
    """Tarjeta con estilo moderno"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet(
            "ModernCard {"
            "    background: white;"
            "    border-radius: 16px;"
            "    border: 1px solid #E5E7EB;"
            "}"
        )