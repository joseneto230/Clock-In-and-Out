import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableView, QPushButton, QVBoxLayout, QWidget, QFileDialog
from PyQt5.QtCore import QAbstractTableModel, Qt
from PyQt5.QtGui import QFont

# Load Excel File
def load_excel(file_name):
    return pd.read_excel(file_name)

# Pandas Model for PyQt Table View
class PandasModel(QAbstractTableModel):
    def __init__(self, df):
        super().__init__()
        self.df = df

    def rowCount(self, parent=None):
        return self.df.shape[0]

    def columnCount(self, parent=None):
        return self.df.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            return str(self.df.iloc[index.row(), index.column()])
        return None

    def setData(self, index, value, role=Qt.EditRole):
        if role == Qt.EditRole:
            self.df.iloc[index.row(), index.column()] = value
            return True
        return False

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return self.df.columns[section]
            if orientation == Qt.Vertical:
                return str(section)
        return None

    def flags(self, index):
        return Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable

# GUI Application
class ExcelViewer(QMainWindow):
    def __init__(self, file_name):
        super().__init__()
        self.file_name = file_name
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Excel Viewer & Editor")
        self.setGeometry(200, 200, 1180, 600)

        self.df = load_excel(self.file_name)
        self.model = PandasModel(self.df)

        self.table = QTableView()
        self.table.setModel(self.model)
        self.table.setFont(QFont("Arial", 10))

        self.save_button = QPushButton("Save Changes")
        self.save_button.clicked.connect(self.save_changes)

        self.load_button = QPushButton("Load New File")
        self.load_button.clicked.connect(self.load_new_file)

        layout = QVBoxLayout()
        layout.addWidget(self.table)
        layout.addWidget(self.save_button)
        layout.addWidget(self.load_button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def save_changes(self):
        self.df.to_excel(self.file_name, index=False)
        print("Changes saved successfully!")

    def load_new_file(self):
        file_dialog = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)")
        if file_dialog[0]:
            self.file_name = file_dialog[0]
            self.df = load_excel(self.file_name)
            self.model = PandasModel(self.df)
            self.table.setModel(self.model)
            print("New file loaded successfully!")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    viewer = ExcelViewer("2025_SSA.xlsx")  # Default file
    viewer.show()
    sys.exit(app.exec_())