import sys
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtCore import QSettings
from modules.order import OrderModule


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("数据分析工具集 v1.0")
        self.settings = QSettings("DataTool", "MainWindow")
        saved_geometry = self.settings.value("geometry")
        if saved_geometry:
            self.restoreGeometry(saved_geometry)
        else:
            self.resize(1200, 800)
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f2f5;
            }
            QPushButton {
                background-color: #ffffff;
                border: 1px solid #d9d9d9;
                border-radius: 6px;
                padding: 10px 18px;
                font-size: 14px;
                color: #333333;
            }
            QPushButton:hover {
                border-color: #1890ff;
                color: #1890ff;
            }
            QPushButton:pressed {
                background-color: #e6f7ff;
            }
            QPushButton:disabled {
                background-color: #f5f5f5;
                color: #bfbfbf;
                border-color: #d9d9d9;
            }
            QTableWidget {
                background-color: #ffffff;
                border: 1px solid #e8e8e8;
                border-radius: 6px;
                gridline-color: #f0f0f0;
                font-size: 13px;
            }
            QTableWidget::item {
                padding: 5px 10px;
            }
            QHeaderView::section {
                background-color: #fafafa;
                border: none;
                border-bottom: 2px solid #e8e8e8;
                padding: 8px 10px;
                font-weight: bold;
                color: #555555;
            }
            QDateEdit {
                background-color: #ffffff;
                border: 1px solid #d9d9d9;
                border-radius: 4px;
                padding: 6px 10px;
                font-size: 13px;
            }
            QDateEdit:focus {
                border-color: #1890ff;
            }
            QLabel {
                color: #555555;
                font-size: 13px;
            }
            QScrollArea {
                border: none;
                background: transparent;
            }
        """)

        self.order_module = OrderModule()
        self.setCentralWidget(self.order_module)

    def closeEvent(self, event):
        self.settings.setValue("geometry", self.saveGeometry())
        super().closeEvent(event)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())