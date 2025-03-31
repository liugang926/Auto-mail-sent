import sys
from PyQt5.QtWidgets import QApplication
from ui import MainWindow
import resources_rc  # 导入编译后的资源文件

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_()) 