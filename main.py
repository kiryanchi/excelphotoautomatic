import sys
import openpyxl
from PyQt5.QtWidgets import QFileDialog, QApplication, QWidget, QLabel, QLayout, QHeaderView, QTableWidget, QVBoxLayout, \
    QHBoxLayout, QAbstractItemView, QTabWidget
from PyQt5.QtGui import QPixmap
from PyQt5 import uic
from PyQt5 import QtCore
from PyQt5.QtCore import QFile, QIODevice, QDataStream, QVariant
import threading
from image.insert import insertinexcel
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SAVE_DIR = BASE_DIR + '\\Save'
form_class = uic.loadUiType(BASE_DIR + '\\UI\\' + "my_form.ui")[0]
FILE_NAME = 'default'


def setLabelText(label, text):
    label.setText(text)
    label.adjustSize()


class TableWidgetPixmap(QPixmap):
    def __init__(self, imgpath):
        super().__init__()
        self.load(imgpath)


class TableWidget(QWidget):
    def __init__(self, imgpath, pixmap=None):
        super().__init__()
        tab_idx = myWindow.sheetlist.currentIndex()
        self.imgpath = imgpath
        self.lbl = QLabel()
        self.img = TableWidgetPixmap(imgpath) if pixmap is None else pixmap
        self.img = self.img.scaled(myWindow.sheetlist.widget(tab_idx).table.width() // 3 - 30, 300 + 10)
        self.lbl.setPixmap(self.img)
        self.widget_layout = QHBoxLayout()
        self.widget_layout.addWidget(self.lbl)
        self.widget_layout.setSizeConstraint(QLayout.SetFixedSize)
        self.setLayout(self.widget_layout)


class MyTable(QTableWidget):
    def __init__(self):
        super().__init__(0, 3)
        self.setAcceptDrops(True)
        self.setSelectionMode(QAbstractItemView.SingleSelection)
        # self.setDragDropMode(QAbstractItemView.InternalMove)
        # self.setDefaultDropAction(QtCore.Qt.CopyAction)
        # 테이블위젯의 헤더 크기를 조정
        self.setAlternatingRowColors(True)
        self.setFocusPolicy(QtCore.Qt.StrongFocus)
        self.setMinimumSize(QtCore.QSize(990, 630))
        self.verticalHeader().setDefaultSectionSize(330)
        self.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.setHorizontalHeaderLabels(['작업 전', '작업 후', '전주 번호'])

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls:
            e.accept()
        else:
            e.ignore()

    def dragMoveEvent(self, e):
        if e.mimeData().hasUrls:
            e.accept()
        else:
            e.ignore()

    def dropEvent(self, e):
        e.setDropAction(QtCore.Qt.CopyAction)
        e.accept()
        if e.mimeData().hasUrls and FILE_NAME != "default":
            url = e.mimeData().urls()[0]
            url = str(url.toLocalFile())
            if url.split('.')[-1] == 'JPG' or url.split('.')[-1] == 'jpg':
                tab_idx = myWindow.sheetlist.currentIndex()
                print(tab_idx)
                current_tab = myWindow.sheetlist.tabText(tab_idx)
                print(url)
                widget = self.CreateTableWidget(url)
                print(widget)
                row = myWindow.sheetlist.widget(tab_idx).table.currentRow()
                print(row)
                col = myWindow.sheetlist.widget(tab_idx).table.currentColumn()
                print(col)
                print('***********')
                myWindow.sheetlist.widget(tab_idx).table.setCellWidget(row, col, widget)

    def CreateTableWidget(self, imgpath, pixmap=None):
        print(imgpath)
        widget = TableWidget(imgpath, pixmap)
        print('done?')
        return widget


class MyTabBar(QWidget):
    def __init__(self):
        super().__init__()
        self.table = MyTable()
        self.table.setEnabled(True)
        vbox = QVBoxLayout()
        vbox.addWidget(self.table)
        self.setLayout(vbox)


class WindowClass(QWidget, form_class):
    table = None
    wb = None
    sheet = None
    tab_bar_list = {}

    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.initUI()

    def initUI(self):
        self.buttonClick()

    def buttonClick(self):
        self.fileopen_btn.clicked.connect(self.openExcel)
        self.save_btn.clicked.connect(self.saveXpa)
        self.load_btn.clicked.connect(self.loadXpa)
        self.reload_btn.clicked.connect(self.reload)
        self.filesave_btn.clicked.connect(self.insert)
        self.delete_btn.clicked.connect(self.delete)
        self.deleteall_btn.clicked.connect(self.deleteall)

    def deleteall(self):
        self.progressOn()
        widget = QWidget()
        for r in range(self.table.rowCount()):
            for c in range(self.table.columnCount()):
                self.table.setCellWidget(r, c, widget)
        self.progressOff()

    def delete(self):
        widget = QWidget()
        r = self.table.currentRow()
        c = self.table.currentColumn()
        self.table.setCellWidget(r, c, widget)

    def insert(self):
        t = threading.Thread(target=self.inserting)
        t.start()

    def inserting(self):
        global FILE_NAME
        self.progressOn()
        column_list = ['A', 'I', 'Q']
        for r in range(self.table.rowCount()):
            for c in range(self.table.columnCount()):
                if self.table.cellWidget(r, c):
                    row = 19 * r + 2
                    imgname = self.table.cellWidget(r, c).imgpath
                    insertinexcel(imgname, column_list[c], row, self.sheet)
        self.wb.save(FILE_NAME)
        setLabelText(self.progress_lbl, '[완료] 엑셀에 사진을 넣었습니다.')
        self.progressOff()

    def reload(self):
        for r in range(self.table.rowCount()):
            for c in range(self.table.columnCount()):
                if self.table.cellWidget(r, c):
                    imgpath = self.table.cellWidget(r, c).imgpath
                    widget = TableWidget(imgpath)
                    self.table.setCellWidget(r, c, widget)

    def loadExcel(self, fname):
        global FILE_NAME
        name = fname.split('/')[-1]
        FILE_NAME = fname
        # 엑셀 파일을 분석한다. openpyxl
        self.sheetlist.clear()
        try:
            self.wb = openpyxl.load_workbook(FILE_NAME)
        except FileNotFoundError:
            setLabelText(self.progress_lbl, '[에러 1] ' + name + ' 파일을 열지 못 했습니다. 엑셀 파일을 확인해주세요.')
            self.progressOff()
        else:
            name_list = self.wb.sheetnames
            for sheet_name in name_list:
                sheet = self.wb[sheet_name]
                row = (sheet.max_row - 1) // 19
                new_tab_bar = MyTabBar()
                new_tab_bar.table.setRowCount(row)
                self.sheetlist.addTab(new_tab_bar, sheet_name)
                self.tab_bar_list[sheet_name] = new_tab_bar
            setLabelText(self.fileopen_lbl, name)
            self.save_btn.setEnabled(True)
            self.reload_btn.setEnabled(True)
            self.filesave_btn.setEnabled(True)
            self.delete_btn.setEnabled(True)
            self.deleteall_btn.setEnabled(True)
            setLabelText(self.progress_lbl, '[완료] 파일을 성공적으로 열었습니다.')
            self.progressOff()
            return True

    def openExcel(self):
        setLabelText(self.progress_lbl, '[진행중] 파일을 여는 중...')
        self.progressOn()
        fname = QFileDialog.getOpenFileName(self, '엑셀 파일 선택', BASE_DIR, "Excel files (*.xlsx)")

        if fname[0]:
            if os.path.isfile(fname[0]):
                self.loadExcel(fname[0])
        else:
            setLabelText(self.progress_lbl, '[에러 0] 파일이 잘못 선택됐습니다. 다시 선택해주세요')
            self.progressOff()

    def saveXpa(self):
        global FILE_NAME
        filters = "xpa File (*.xpa);; xpae File (*.xpae)"
        setLabelText(self.progress_lbl, '[진행중] 작업을 저장중...')
        self.progressOn()
        excel_name = FILE_NAME.split('.')[0]
        save_file_name = QFileDialog.getSaveFileName(self, "파일 저장하기", filter=filters)
        print(save_file_name)
        print(self.sheetlist.count())
        save_file = QFile(save_file_name[0])
        save_file.open(QIODevice.WriteOnly)
        save_file.open(QIODevice.Append)
        stream_out = QDataStream(save_file)
        # 엑셀 파일 이름을 먼저 저장
        file_name = QVariant(excel_name)
        stream_out << file_name
        if save_file_name[0].split('.')[1] == 'xpa':
            # 시트 이름을 저장
            for i in range(self.sheetlist.count()):
                tab_str = QVariant(self.sheetlist.tabText(i))
                stream_out << tab_str
                for r in range(self.sheetlist.widget(i).table.rowCount()):
                    for c in range(self.sheetlist.widget(i).table.columnCount()):
                        if self.sheetlist.widget(i).table.cellWidget(r, c):
                            output_str = QVariant(self.sheetlist.widget(i).table.cellWidget(r, c).imgpath)
                            stream_out << output_str
                        else:
                            output_str = QVariant('Null')
                            stream_out << output_str

        elif save_file_name[0].split('.')[1] == 'xpae':
            # 시트 이름을 저장
            for i in range(self.sheetlist.count()):
                tab_str = QVariant(self.sheetlist.tabText(i))
                stream_out << tab_str
                for r in range(self.sheetlist.widget(i).table.rowCount()):
                    for c in range(self.sheetlist.widget(i).table.columnCount()):
                        if self.sheetlist.widget(i).table.cellWidget(r, c):
                            output_str = QPixmap(self.sheetlist.widget(i).table.cellWidget(r, c).img)
                            stream_out << output_str
                        else:
                            output_str = QVariant('Null')
                            stream_out << output_str
        save_file.close()
        setLabelText(self.progress_lbl, '[완료] 작업을 저장했습니다.')
        self.progressOff()

    def loadXpa(self):
        global FILE_NAME, SAVE_DIR
        self.progressOn()
        save_file_name = FILE_NAME.split('/')[-1]
        save_file_name = save_file_name.split('.')[0]
        save_file_name = SAVE_DIR + '\\' + save_file_name + '.xpa'
        if not os.path.isfile(save_file_name):
            setLabelText(self.progress_lbl, '[에러 2] 작업 파일이 없습니다. %s 폴더를 다시 확인해보세요.' % 'Save')
            self.progressOff()
            return
        try:
            load_file = QFile(save_file_name)
            load_file.open(QIODevice.ReadOnly)
            stream_in = QDataStream(load_file)
            for r in range(self.table.rowCount()):
                for c in range(self.table.columnCount()):
                    input_str = QVariant()
                    stream_in >> input_str
                    if input_str.value() != 'Null':
                        widget = TableWidget(input_str.value())
                        self.table.setCellWidget(r, c, widget)
            load_file.close()
            setLabelText(self.progress_lbl, '[완료] 작업 파일을 성공적으로 불러왔습니다.')
        except:
            setLabelText(self.progress_lbl, '[에러 3] 작업 파일을 불러올 수 없습니다. 다시 확인해보세요.sss')
        self.progressOff()

    def progressOn(self):
        self.progress_bar.setMaximum(0)

    def progressOff(self):
        self.progress_bar.setMaximum(1)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    app.exec_()
