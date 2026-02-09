"""
Modelo Qt para visualizar DataFrames de pandas
"""
import pandas as pd
from PySide6.QtCore import Qt, QAbstractTableModel


class PandasModel(QAbstractTableModel):
    """Modelo para mostrar DataFrames en QTableView"""
    
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