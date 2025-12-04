import sys
import pandas as pd
from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, 
    QHBoxLayout, QTableView, QFileDialog, QLabel
)
from PySide6.QtCore import Qt, QAbstractTableModel
from PySide6.QtGui import QFont


# -------------------------
# Modelo para mostrar pandas en la tabla
# -------------------------
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


# -------------------------
# Ventana principal
# -------------------------
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Gestor de Monitores – PySide6")
        self.setMinimumSize(900, 600)

        # Estilos modernos
        self.setStyleSheet("""
            QWidget {
                background-color: #F4F6F9;
                font-family: 'Segoe UI';
                font-size: 14px;
            }
            QPushButton {
                background-color: #0078D4;
                color: white;
                padding: 10px;
                border-radius: 8px;
            }
            QPushButton:hover {
                background-color: #005A9E;
            }
            QTableView {
                background: white;
                border-radius: 8px;
            }
        """)

        # Layout principal
        layout = QVBoxLayout()

        # Título
        title = QLabel("Gestión de Monitores y Espacios")
        title.setFont(QFont("Segoe UI", 20, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # Botones
        btn_layout = QHBoxLayout()

        self.btn_monitores = QPushButton("Cargar Monitores")
        self.btn_espacios = QPushButton("Cargar Espacios")
        self.btn_asignar = QPushButton("Asignar")

        btn_layout.addWidget(self.btn_monitores)
        btn_layout.addWidget(self.btn_espacios)
        btn_layout.addWidget(self.btn_asignar)

        layout.addLayout(btn_layout)

        # Tabla
        self.table = QTableView()
        layout.addWidget(self.table)

        self.setLayout(layout)

        # Conectar funciones
        self.btn_monitores.clicked.connect(self.cargar_monitores)
        self.btn_espacios.clicked.connect(self.cargar_espacios)
        self.btn_asignar.clicked.connect(self.asignar)

        # DataFrames cargados
        self.df_monitores = pd.DataFrame()
        self.df_espacios = pd.DataFrame()

    # --------------------- FUNCIONES ---------------------

    def cargar_monitores(self):
        ruta, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar archivo de monitores", "", "Archivos Excel (*.xlsx *.xls)"
        )
        if ruta:
            self.df_monitores = pd.read_excel(ruta)
            self.table.setModel(PandasModel(self.df_monitores))

    def cargar_espacios(self):
        ruta, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar archivo de espacios", "", "Archivos Excel (*.xlsx *.xls)"
        )
        if ruta:
            self.df_espacios = pd.read_excel(ruta)
            self.table.setModel(PandasModel(self.df_espacios))

    def asignar(self):
        if self.df_monitores.empty or self.df_espacios.empty:
            return  # podrías mostrar un mensaje de error

        # EJEMPLO de asignación (reemplázalo con tu lógica real)
        asignaciones = pd.DataFrame({
            "Monitor": self.df_monitores.iloc[:, 0],
            "Espacio": self.df_espacios.iloc[:, 0].reindex(range(len(self.df_monitores))).fillna("")
        })

        self.table.setModel(PandasModel(asignaciones))


# -------------------------
# Ejecutar aplicación
# -------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
