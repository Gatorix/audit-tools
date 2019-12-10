import sys
from PyQt5.QtWidgets import QApplication, QMainWindow
from untitled import Ui_MainWindow


class MyMainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyMainWindow, self).__init__(parent)
        self.setupUi(self)


app = QApplication(sys.argv)
myWin = MyMainWindow()
myWin.show()
sys.exit(app.exec_())
