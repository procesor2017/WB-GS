import sys
from PySide2.QtWidgets import QMainWindow, QApplication
from window import Ui_Tabulky



class MainWindow(QMainWindow, Ui_Tabulky):
    def __init__(self):
        QMainWindow.__init__(self)
        Ui_Tabulky.__init__(self)
        self.setupUi(self)
        self.retranslateUi(self)
        self.changetext(self)
        self.show()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = MainWindow()
    ret = app.exec_()
    sys.exit(ret)
