import data_worker as dw
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QApplication, QMainWindow
import sys
import os.path


class MyTread(QtCore.QThread):
    mysignal = QtCore.pyqtSignal(str)
    params = {}
    table = None

    def __init__(self, parent=None):
        QtCore.QThread.__init__(self, parent)

    def run(self):
        error = None
        try:
            match self.params['func']:
                case 'click_btn_upload':
                    if os.path.exists("fns.db") is False:
                        self.mysignal.emit("Отсутствует файл с базой fns.db")
                    elif self.params['file_in'] == '':
                        self.mysignal.emit("Файл не выбран")
                    else:
                        self.mysignal.emit("Читаем файл")
                        rows = dw.read_data(self.params['file_in'], self.mysignal)
                        self.table.setRowCount(len(rows))
                        for i, row in enumerate(rows):
                            for j, col in enumerate(row):
                                self.table.setItem(i, j, QtWidgets.QTableWidgetItem(str(col)))
                        self.mysignal.emit("чтение завершено")  # тут добавить рез загрузки
                case 'click_btn_download':
                    if os.path.exists("fns.db") is False:
                        self.mysignal.emit("Отсутствует файл с базой fns.db")
                    elif self.params['file_out'] == '':
                        self.mysignal.emit("Файл не выбран")
                    else:
                        self.mysignal.emit("Записываем в файл")
                        error = dw.write_data(self.params['file_out'])
                        if error is None:
                            self.mysignal.emit("файл сохранен")
                        else:
                            self.mysignal.emit(error)
                case 'run_process1':
                    dw.process1(self.mysignal)
                case 'run_process2':
                    dw.process2(self.mysignal)
        except Exception as err_all:
            self.mysignal.emit(f"Непредвиденная ошибка: {err_all}")


class Window(QMainWindow):
    def __init__(self):
        self.func = ''
        self.file_in = ''
        super(Window, self).__init__()
        self.setWindowTitle("Сбор чеков")
        self.setGeometry(300, 150, 800, 800)
        self.table = QtWidgets.QTableWidget(self)
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["NN", "Сумма чека", "Дата время чека", "фиск номер", "фиск документ", "фиск признак"])
        self.table.move(20, 200)
        self.table.setMinimumWidth(750)
        self.table.setMinimumHeight(520)
        self.table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.btn_upload = QtWidgets.QPushButton(self)  # кнопка загрузить из файла
        self.btn_upload.move(50, 20)
        self.btn_upload.setText("Загрузить из файла")
        self.btn_upload.setFixedWidth(200)
        self.btn_upload.clicked.connect(self.click_btn_upload)
        self.btn_download = QtWidgets.QPushButton(self)  # кнопка сохранить в файл
        self.btn_download.move(250, 20)
        self.btn_download.setText("Сохранить в файл")
        self.btn_download.setFixedWidth(200)
        self.btn_download.clicked.connect(self.click_btn_download)
        self.btn_process = QtWidgets.QPushButton(self)  # кнопка сохранить в файл
        self.btn_process.move(450, 20)
        self.btn_process.setText("Начать проверку")
        self.btn_process.setFixedWidth(200)
        self.btn_process.clicked.connect(self.click_btn_process)
        self.txt_logs = QtWidgets.QTextEdit(self)
        self.txt_logs.resize(600, 100)
        self.txt_logs.move(50, 60)
        self.txt_logs.setReadOnly(True)
        self.mythread = MyTread()
        self.mythread.finished.connect(self.mythread_finish)
        self.mythread.mysignal.connect(self.mythread_change, QtCore.Qt.QueuedConnection)



    def mythread_change(self, s):
        # выводим информацию в лог
        self.txt_logs.append(s)

    def mythread_finish(self):
        match self.func:
            case 'click_btn_upload':
                self.btn_upload.setEnabled(True)
            case 'click_btn_download':
                self.btn_download.setEnabled(True)
            case 'run_process1':
                self.btn_process.setEnabled(True)
                #run process2
                ret = QtWidgets.QMessageBox.question(self, "Продолжение", "Запускать проверку чеков с перебором минут?",
                                                     QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
                if ret == QtWidgets.QMessageBox.Yes:
                    self.btn_process.setEnabled(False)
                    self.run_process2()
                else:
                    QtWidgets.QMessageBox.information(self, "ИНФО", "Проверка завершена")
            case 'run_process2':
                self.btn_process.setEnabled(True)
                QtWidgets.QMessageBox.information(self, "ИНФО", "Проверка завершена")


    def click_btn_upload(self):
        self.btn_upload.setEnabled(False)
        tmp = QtWidgets.QFileDialog.getOpenFileName(self, "Выберите файл", "", "Excel files (*.xlsx)")
        file_in = tmp[0]
        self.func = 'click_btn_upload'
        self.mythread.params = {'func': self.func, 'file_in': file_in}
        self.mythread.table = self.table
        self.mythread.start()

    def click_btn_download(self):
        self.btn_download.setEnabled(False)
        tmp = QtWidgets.QFileDialog.getSaveFileName(self, "Выберите файл", "", "Excel files (*.xlsx)")
        file_out = tmp[0]
        self.func = 'click_btn_download'
        self.mythread.params = {'func': self.func, 'file_out': file_out}
        self.mythread.start()


    def click_btn_process(self):
        self.btn_process.setEnabled(False)
        self.func = 'run_process1'
        self.mythread.params = {'func': self.func}
        self.mythread.start()


    def run_process2(self):
        self.func = 'run_process2'
        self.mythread.params = {'func': self.func}
        self.mythread.start()



def application():
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    application()
