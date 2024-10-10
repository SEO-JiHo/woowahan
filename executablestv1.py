import sys
import os
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtGui import QIcon


import json
import urllib
from urllib.request import urlopen
from urllib.parse import quote
from openpyxl.workbook import Workbook

form_class = uic.loadUiType(r"C:\Users\Park\PycharmProjects\stv1\uiv1.ui")[0]

class WindowClass(QMainWindow, form_class) :

    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle('Searching Tool v1')
        self.setWindowIcon(QIcon(r"C:\Users\Park\PycharmProjects\stv1\cn.png"))

        self.searchInput1.returnPressed.connect(self.printTextFunction)
        self.search1.clicked.connect(self.clickFunction)

    def printTextFunction(self):
        self.searchOutput1.clear()

        global search
        search = self.searchInput1.text()
        url = f"https://m.map.naver.com/search2/searchMore.naver?query={quote(search)}{quote('캠핑용품')}"
        text_data = urllib.request.urlopen(url).read().decode('utf-8')

        global company
        company = json.loads(text_data)
        coreInfo = company['result']['site']['list']
        for k in coreInfo:
            self.searchOutput1.appendPlainText(k['name'])

    def clickFunction(self):
        self.searchOutput1.clear()

        global search
        search = self.searchInput1.text()
        url = f"https://m.map.naver.com/search2/searchMore.naver?query={quote(search)}{quote('캠핑용품')}"
        text_data = urllib.request.urlopen(url).read().decode('utf-8')

        global company
        company = json.loads(text_data)
        coreInfo = company['result']['site']['list']
        for k in coreInfo:
            self.searchOutput1.appendPlainText(k['name'])

    def folderSelectSignal(self):
        self.searchInput1.clear()

        folderPath = QFileDialog.getExistingDirectory()
        folderPath = os.path.realpath(folderPath)
        raw_fPath = r"{}".format(folderPath)
        print(raw_fPath)
        QMessageBox.information(self, "엑셀 파일 저장 완료", f"{search} 캠핑용품.xlsx  ")

        wb = Workbook()
        ws = wb.active
        ws.column_dimensions['A'].width = 5
        ws.column_dimensions['B'].width = 22
        ws.column_dimensions['C'].width = 14
        ws.column_dimensions['D'].width = 50
        ws.column_dimensions['E'].width = 37
        ws.column_dimensions['F'].width = 20

        ws.append(['번호', '업체명', '전화번호', '주소', '홈페이지', '썸네일 이미지 주소', '이미지'])
        for i in company['result']['site']['list']:
            ws.append([i['rank'], i['name'], i['tel'], i['roadAddress'], i['homePage'], i['thumUrl']])
            wb.save(f'{raw_fPath}/{search} 캠핑용품.xlsx')
        # save code
        folderPath2 = folderPath
        return folderPath2

if __name__ == "__main__" :
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    app.exec_()