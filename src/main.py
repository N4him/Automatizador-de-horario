"""
Punto de entrada principal de la aplicación
"""
import sys
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QFont
from gui.widgets.main_window import MainWindow

def main():
    """Función principal"""
    app = QApplication(sys.argv)
    
    # Configurar fuente
    font = QFont("Inter")
    if not font.exactMatch():
        font = QFont("Segoe UI")
    app.setFont(font)
    
    # Crear y mostrar ventana
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()