# This is the main program that runs the GUI
import sys
from PyQt5.QtWidgets import QApplication
from mainwindow import MainWindow

# 确保所有类定义完整后再实例化
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())