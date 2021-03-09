import sys
import openpyxl
import pprint

from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from interface import Ui_MainWindow

TITLE_DICT = {"$1": "Название компании",
              "$2": "Тип упаковки",
              "$3": "Материал упаковки",
              "$4": "Партия",
              "$5": "Срок годности",
              "$6": "Дата производства",
              "$7": "Количество в палетте (литров)",
              "$8": "",
              "$9": "",
              "$10": "",
              "$11": "",
              "$12": "",
              "$13": "",
              "$14": "",
              "$15": "",
              "$16": "",
              "$17": "",
              "$18": "",
              "$19": "",
              "$20": "",
              "$21": "",
              "$22": ""}


class IndexNotFound(Exception):

    def __init__(self, indexes):
        super(IndexNotFound, self).__init__()
        self.indexes = indexes


class MyWidget(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.pick_btn.clicked.connect(self.pick_file)
        self.search_btn.clicked.connect(self.search)
        self.data_input.editingFinished.connect(self.search)
        self.matrix = []
        self.consignment_ind = None
        self.gtin = None

    def search(self):
        gtin, consignment = "", ""
        if self.data_input.text():
            line = self.data_input.text()
            gtin = line[line.index("020") + 3:line.index("020") + 16]
            if "10" in line and "11" in line:
                consignment = line[line.index("10") + 2:line.index("11") - 1]
            if self.matrix:
                for row in self.matrix:
                    if row[self.consignment_ind] == consignment or row[self.gtin] == gtin:
                        self.statusbar.showMessage("строка найдена!", msecs=5000)
                        break
                else:
                    self.statusbar.showMessage("строка не найдена!", msecs=5000)
            else:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Information)
                msg.setText(f"не выбран файл!")
                # msg.setInformativeText("добавочная")
                # msg.setWindowTitle("ошибка")
                # msg.setDetailedText(f"""""")
                msg.exec_()
        else:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText(f"Введите штрих-код!")
            msg.exec_()

    def pick_file(self):
        try:
            title = QFileDialog.getOpenFileName(self,
                                                "Выбрать таблицу Excel",
                                                "",
                                                "Tables (*.xlsx)")[0]
            wb = openpyxl.load_workbook(filename=title)
            sheet = wb.active
            matrix = [[cell.value for cell in row] for row in sheet.rows]
            check_row = [f"${i}" for i in range(1, 23)]
            nir = None
            for row in range(len(matrix)):
                if any([ind in matrix[row] for ind in check_row]):
                    nir = row  # needed_index_row
                    break
            if isinstance(nir, int):
                if all([ind in matrix[nir] for ind in check_row]):
                    self.matrix = matrix
                    self.consignment_ind = [ind for ind in range(len(matrix[nir])) if matrix[nir][ind] == "$4"][0]
                    self.gtin = [ind for ind in range(len(matrix[nir])) if matrix[nir][ind] == "$22"][0]
                else:
                    raise IndexNotFound([ind for ind in check_row if ind not in matrix[nir]])
            else:
                raise IndexNotFound(check_row)
        except IndexNotFound as error:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText(f"не найдены индексы {error.indexes}")
            # msg.setInformativeText("добавочная")
            # msg.setWindowTitle("ошибка")
            msg.setDetailedText(f"""""")
            msg.exec_()
        except openpyxl.utils.exceptions.InvalidFileException as error:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText(f"Либо данный файл не поддерживается, либо вы не выбрали файл. Расширение Excel файла - .xlsx")
            # msg.setInformativeText("добавочная")
            # msg.setWindowTitle("ошибка")
            # msg.setDetailedText(f"""""")
            msg.exec_()


def except_hook(cls, exception, traceback):
    sys.__excepthook__(cls, exception, traceback)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    form = MyWidget()
    form.show()
    sys.excepthook = except_hook
    sys.exit(app.exec())
