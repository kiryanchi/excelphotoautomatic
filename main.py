import sys
import openpyxl
from PyQt5.QtWidgets import QFileDialog, QApplication, QWidget
from PyQt5.QtGui import QPixmap
from PyQt5 import uic
from image.insert import insertimg
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
form_class = uic.loadUiType("my_form.ui")[0]


def setLabelText(label, text):
    label.setText(text)
    label.adjustSize()


class WindowClass(QWidget, form_class):
    file_name = 'default'
    wb = None
    sheet = None

    def __init__(self):
        """
        Widgets                 tab1                    tab2                    tab3
        fileopen                beforeimg               afterimg                polenum
        fileopen_btn            beforeimg_lbl           afterimg_lbl            polenum_lbl
        fileopen_lbl            beforeimg_btn           afterimg_btn            polenum_btn
        save_btn                beforeimg_list          afterimg_list           polenum_list
        progress_lbl            beforeimg_insert_btn    afterimg_insert_btn     polenum_insert_btn
        alert_lbl               beforeimg_clear_btn     afterimg_clear_btn      polenum_clear_btn
                                beforeimg_remove_btn    afterimg_remove_btn     polenum_remove_btn
        """
        super().__init__()
        self.setupUi(self)
        self.saveOff()
        # 파일 오픈
        self.fileopen_btn.clicked.connect(self.openExcel)

        # 사진 불러오기
        self.beforeimg_btn.clicked.connect(self.openBeforePhoto)
        self.afterimg_btn.clicked.connect(self.openAfterPhoto)
        self.polenum_btn.clicked.connect(self.openPolePhoto)

        # 사진 비우기
        self.beforeimg_clear_btn.clicked.connect(self.clearBeforePhoto)
        self.afterimg_clear_btn.clicked.connect(self.clearAfterPhoto)
        self.polenum_clear_btn.clicked.connect(self.clearPolePhoto)

        # 사진 삭제
        self.beforeimg_remove_btn.clicked.connect(self.removeBeforePhoto)
        self.afterimg_remove_btn.clicked.connect(self.removeAfterPhoto)
        self.polenum_remove_btn.clicked.connect(self.removePolePhoto)

        # 사진 출력
        self.beforeimg_list.itemClicked.connect(self.viewBeforePhoto)
        self.afterimg_list.itemClicked.connect(self.viewAfterPhoto)
        self.polenum_list.itemClicked.connect(self.viewPolePhoto)

        # 엑셀에 사진 삽입하기
        self.beforeimg_insert_btn.clicked.connect(self.insertBeforePhoto)
        self.afterimg_insert_btn.clicked.connect(self.insertAfterPhoto)
        self.polenum_insert_btn.clicked.connect(self.insertPolePhoto)

        # 저장
        self.save_btn.clicked.connect(self.saveExcel)

    # 함수들
    def saveExcel(self):
        self.wb.save(self.file_name)
        setLabelText(self.progress_lbl, '[완료] 저장했습니다.')
        self.saveOff()

    # 수정 용이하게 위에서 작업
    def insertBeforePhoto(self):
        # default 이면 작업 X
        if self.file_name == 'default':
            setLabelText(self.progress_lbl, '[에러 2] 파일을 선택한 후 진행해주세요')
            return
        img_name = self.beforeimg_list.currentItem().text()
        row = self.spinBox.value()
        row = (row - 1) * 19 + 2
        result = insertimg(img_name, 'A', row, self.sheet)
        setLabelText(self.progress_lbl, result)
        self.saveOn()

    def insertAfterPhoto(self):
        # default 이면 작업 X
        if self.file_name == 'default':
            setLabelText(self.progress_lbl, '[에러 2] 파일을 선택한 후 진행해주세요')
            return
        img_name = self.beforeimg_list.currentItem().text()
        row = self.spinBox.value()
        row = (row - 1) * 19 + 2
        result = insertimg(img_name, 'I', row, self.sheet)
        setLabelText(self.progress_lbl, result)
        self.saveOn()

    def insertPolePhoto(self):
        # default 이면 작업 X
        if self.file_name == 'default':
            setLabelText(self.progress_lbl, '[에러 2] 파일을 선택한 후 진행해주세요')
            return
        img_name = self.beforeimg_list.currentItem().text()
        row = self.spinBox.value()
        row = (row - 1) * 19 + 2
        result = insertimg(img_name, 'Q', row, self.sheet)
        setLabelText(self.progress_lbl, result)
        self.saveOn()

    # 엑셀 파일 여는 함수
    def openExcel(self):
        setLabelText(self.progress_lbl, '[진행중] 파일을 여는 중...')
        fname = QFileDialog.getOpenFileName(self, '엑셀 파일 선택', BASE_DIR, "Excel files (*.xlsx)")
        # print(fname)

        if fname[0]:
            if os.path.isfile(fname[0]):
                name = fname[0].split('/')[-1]
                self.file_name = fname[0]
                # 엑셀 파일을 분석한다. openpyxl
                self.wb = openpyxl.load_workbook(self.file_name)
                try:
                    self.sheet = self.wb['작업사진']
                except KeyError:
                    setLabelText(self.progress_lbl, "[에러 1] '작업사진' 워크시트를 열지 못 했습니다. 엑셀파일을 확인해주세요.")
                else:
                    row = (self.sheet.max_row - 1) // 19
                    self.spinBox.setRange(1, row)
                    setLabelText(self.rowlabel, '원하는 행 (1 ~ ' + str(row) + ')')
                    setLabelText(self.fileopen_lbl, name)
                    setLabelText(self.progress_lbl, '[완료] 파일을 성공적으로 열었습니다.')
        else:
            setLabelText(self.progress_lbl, '[에러 0] 파일이 잘못 선택됐습니다. 다시 선택해주세요')

    # 사진 불러오는 함수들
    def openBeforePhoto(self):
        p_list = QFileDialog.getOpenFileNames(self, '사진 열기', BASE_DIR, "Image files (*.JPG *.png)")

        if p_list[0]:
            for img in p_list[0]:
                self.beforeimg_list.addItem(img)

    def openAfterPhoto(self):
        p_list = QFileDialog.getOpenFileNames(self, '사진 열기', BASE_DIR, "Image files (*.JPG *.png)")

        if p_list[0]:
            for img in p_list[0]:
                self.afterimg_list.addItem(img)

    def openPolePhoto(self):
        p_list = QFileDialog.getOpenFileNames(self, '사진 열기', BASE_DIR, "Image files (*.JPG *.png)")

        if p_list[0]:
            for img in p_list[0]:
                self.polenum_list.addItem(img)

    # 추가된 사진을 보는 함수들
    def viewBeforePhoto(self):
        img = QPixmap()
        item = self.beforeimg_list.currentItem().text()
        img.load(item)
        img = img.scaled(self.beforeimg_lbl.width(), self.beforeimg_lbl.height())
        self.beforeimg_lbl.setPixmap(img)

    def viewAfterPhoto(self):
        img = QPixmap()
        item = self.afterimg_list.currentItem().text()
        img.load(item)
        img = img.scaled(self.afterimg_lbl.width(), self.afterimg_lbl.height())
        self.afterimg_lbl.setPixmap(img)

    def viewPolePhoto(self):
        img = QPixmap()
        item = self.polenum_list.currentItem().text()
        img.load(item)
        img = img.scaled(self.polenum_lbl.width(), self.polenum_lbl.height())
        self.polenum_lbl.setPixmap(img)

    # 리스트에 추가된 사진들을 지우는 함수들
    def clearBeforePhoto(self):
        self.beforeimg_list.clear()
        setLabelText(self.beforeimg_lbl, '전 사진을 비웠습니다.')

    def clearAfterPhoto(self):
        self.afterimg_list.clear()
        setLabelText(self.afterimg_lbl, '후 사진을 비웠습니다.')

    def clearPolePhoto(self):
        self.polenum_list.clear()
        setLabelText(self.polenum_lbl, '전주번호 사진을 비웠습니다.')

    def removeBeforePhoto(self):
        self.beforeimg_list.takeItem(self.beforeimg_list.currentRow())
        setLabelText(self.beforeimg_lbl, '선택한 사진을 삭제했습니다.')

    def removeAfterPhoto(self):
        self.afterimg_list.takeItem(self.afterimg_list.currentRow())
        setLabelText(self.afterimg_lbl, '선택한 사진을 삭제했습니다.')

    def removePolePhoto(self):
        self.polenum_list.takeItem(self.polenum_list.currentRow())
        setLabelText(self.polenum_lbl, '선택한 사진을 삭제했습니다.')

    def saveOn(self):
        self.save_btn.setEnabled(True)

    def saveOff(self):
        self.save_btn.setEnabled(False)


if __name__ == "__main__" :
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    app.exec_()