from windows import MainWindow
from operation_component import *
from PyQt5.QtWidgets import QApplication, QWidget
import sys

app = QApplication(sys.argv)
point = Point(300, 300)
mw = MainWindow(200, 200, point, "main window")
# 使程序不会立即退出
sys.exit(app.exec_())
