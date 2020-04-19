import sys
from PyQt5.QtWidgets import QApplication, QWidget, QDesktopWidget

NAME = '엑셀 자동 작업 프로그램'
SIZE = (500, 350)


class MyApp(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def my_menuBar_File(self):
        menubar = self.menuBar()
        

    def initUI(self):
        self.setWindowTitle(NAME)
        self.resize(SIZE[0], SIZE[1])
        self.center()
        self.show()

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())