import sys
from PyQt5.QtWidgets import QApplication


class WinForm(QtWidget):
    def __init__(self, parent=None):
        super(WinForm, self).__init__(parent)
        self.setGeometry(300, 300, 350, 350)
        self.setWindowTitle('点击关闭')
        quit = QPushButton('Close', self)
        quit.setGeometry(10, 10, 60, 35)
        quit.setStyleSheet("background-color:red")
        quit.clicked.connect(self.close)

if __name__=="__main__":
    app=QApplication(sys.argv)
    win=WinForm()
    win.show()
    sys.exit(app.exec_())
# import sys
# import PyQt5.
# import PyQt5.QtCore

# if __name__ == '__main__':

#     app = QApplication(sys.argv)    
#     btn = QPushButton("Hello PyQt5")
#     btn.clicked.connect(QCoreApplication.instance().quit)
#     btn.resize(400,100)
#     btn.move(50,50)
#     btn.show()

#     sys.exit(app.exec_())