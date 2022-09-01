# -*- coding: utf-8 -*-
from PyQt5 import QtWidgets, QtTest, QtCore
import os, sys
import ui_main
import configparser
import datetime
import win32com.client


class MainWindow(QtWidgets.QMainWindow, ui_main.Ui_MainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setupUi(self)

        self._generator = None
        self._timerId = None

        self.config_main = configparser.ConfigParser()
        self.config_main.read('setting.ini', "utf8")
        self.default_path = self.config_main['DIR']['default_path']

        self.lineEdit.setText(str(self.default_path))

        self.pushButton.clicked.connect(self.dir_path)
        self.pushButton_2.clicked.connect(self.start)

        self.th_work = Worker(parent=self)
        self.th_work.run_state.connect(self.update)

    def dir_path(self):
        info_dir = QtWidgets.QFileDialog.getExistingDirectory(self, "Select Directory")
        info_dir = str(info_dir)
        if info_dir == '':
            info_dir = self.config_main.get('DIR', 'default_path')
        self.config_main['DIR']['default_path'] = info_dir
        self.lineEdit.setText(info_dir)

        with open('setting.ini', 'w') as configfile:
            self.config_main.write(configfile)

    @QtCore.pyqtSlot()
    def start(self):
        self.pushButton_2.setEnabled(False)
        self.th_work.start()

    @QtCore.pyqtSlot(str)
    def update(self, msg):
        if msg == 'Complete':
            msg = QtWidgets.QMessageBox()
            msg.setText("작업이 완료 되었습니다.")
            msg.setWindowTitle("알림")
            msg.setIcon(QtWidgets.QMessageBox.Information)
            msg.setStyleSheet("")
            msg.exec_()
            self.pushButton_2.setEnabled(True)


class Worker(QtCore.QThread):
    run_state = QtCore.pyqtSignal(str)

    def __init__(self, parent=None):
        super(Worker, self).__init__(parent)
        self.parent = parent
        self.running = True

    def run(self):
        path_dir = self.parent.lineEdit.text()
        print(path_dir)

        path_dir = os.path.abspath(path_dir)

        file_list = os.listdir(path_dir)
        file_list.sort()

        file_list = [file for file in file_list if (file.endswith(".xls") or file.endswith(".xlsx"))]
        file_list = [file for file in file_list if not (file.startswith("~$"))]

        print('Total file No. : ' + str(len(file_list)))
        no = 0
        excel = win32com.client.Dispatch("Excel.Application")

        for item in file_list:
            no = no + 1
            now = datetime.datetime.now()
            nowDatetime = now.strftime('[%Y-%m-%d %H:%M:%S]')
            log_text = str(nowDatetime) + ' Progress : ( ' + str(no) + ' / ' + str(len(file_list)) + ' )'
            self.parent.listWidget.addItem(log_text)
            self.parent.listWidget.scrollToBottom()
            QtTest.QTest.qWait(50)
            excel.Visible = False
            WB_PATH = path_dir + '\\' + (item)
            now = datetime.datetime.now()
            nowDatetime = now.strftime('[%Y-%m-%d %H:%M:%S]')
            log_text = str(nowDatetime) + ' XLSX File : ' + WB_PATH
            self.parent.listWidget.addItem(log_text)
            self.parent.listWidget.scrollToBottom()
            QtTest.QTest.qWait(50)
            wb = excel.Workbooks.Open(WB_PATH)
            ws_index_list = [1]
            wb.WorkSheets(ws_index_list).Select()
            PATH_TO_PDF = item
            PATH_TO_PDF = item[0:-5] + u'.pdf'
            PATH_TO_PDF = path_dir + '\\' + PATH_TO_PDF
            now = datetime.datetime.now()
            nowDatetime = now.strftime('[%Y-%m-%d %H:%M:%S]')
            log_text = str(nowDatetime) + ' PDF File : ' + PATH_TO_PDF
            self.parent.listWidget.addItem(log_text)
            self.parent.listWidget.scrollToBottom()
            QtTest.QTest.qWait(50)
            wb.ActiveSheet.ExportAsFixedFormat(0, PATH_TO_PDF)
            wb.Close(True)
            QtTest.QTest.qWait(50)
        excel.quit()
        self.run_state.emit('Complete')


if __name__=="__main__":
    app = QtWidgets.QApplication(sys.argv)
    form = MainWindow()
    form.show()
    app.exec_()
    print("DONE")