import os
import pathlib
import platform
import sys
import csv
import pandas as pd
import xlrd
from PyQt5.QtWidgets import QStyleFactory, QApplication, QVBoxLayout, QGroupBox, QPushButton, \
    QDialog, QGridLayout, QLineEdit, QComboBox, QFileDialog
from PyQt5.QtCore import Qt


class mainWindow(QDialog):

    def __init__(self):
        self.voltageDict = {}
        self.currentDict = {}
        self.download_path = './'.replace('/', '\\')
        # self.download_path = 'YT Songs'.replace('/','\\')

        if not os.path.exists(self.download_path):
            os.mkdir(self.download_path)

        os.chdir(self.download_path)

        super(mainWindow, self).__init__()

        self.setWindowTitle("Worst day finder")
        self.setWindowModality(Qt.ApplicationModal)
        self.setStyle(QStyleFactory.create('Cleanlooks'))
        self.resize(400, 100)

        self.Dialog_Layout = QVBoxLayout()

        self.GroupBox = QGroupBox(
            'Choose csv')
        self.GroupBox_Layout = QGridLayout()
        self.GroupBox_Layout.setAlignment(Qt.AlignCenter)

        self.Close_Button = QPushButton("Upload")
        self.Close_Button.clicked.connect(self.load_csv)

        self.GroupBox_Layout.addWidget(self.Close_Button, 1, 1)

        self.GroupBox.setLayout(self.GroupBox_Layout)
        self.Dialog_Layout.addWidget(self.GroupBox)

        self.setLayout(self.Dialog_Layout)

    def load_csv(self):

        fname = QFileDialog.getOpenFileName(self, 'Open CSV/XLSX file',
                                            str(pathlib.Path.home()), "CSV files (*.csv;*.xlsx)")

        if platform.system() == "Windows":
            fname = fname[0].replace('/', '\\')
        else:
            fname = fname[0]

        if fname.split('.')[-1] == 'csv':
            with open(fname, 'r', newline='') as file:
                dataCSV = {}
                csv_reader_object = csv.reader(file)
                name_col = next(csv_reader_object)

                for row in name_col:
                    dataCSV[row] = []

                for row in csv_reader_object:  ##Varendo todas as linhas
                    for ndata in range(0, len(name_col)):  ## Varendo todas as colunas
                        dataCSV[name_col[ndata]].append(row[ndata])
            print(dataCSV)

        elif fname.split('.')[-1] == 'xlsx':
            dataCSV = {}
            workbook = xlrd.open_workbook(fname, on_demand=True)
            worksheet = workbook.sheet_by_index(0)
            first_row = []  # The row where we stock the name of the column

            for col in range(worksheet.ncols):
                first_row.append(worksheet.cell_value(0, col))

            for col in first_row:
                dataCSV[col] = []

            for row in range(1, worksheet.nrows):
                for col in range(0, worksheet.ncols):  ## Varendo todas as colunas
                    dataCSV[first_row[col]].append(worksheet.cell_value(row, col))

            for key in dataCSV.keys():
                if str(key).lower() == 'data':
                    self.voltageDict[key] = dataCSV[key].copy()
                    self.currentDict[key] = dataCSV[key].copy()

                if str(key).lower().find('corrente') > 0:
                    self.currentDict[key] = dataCSV[key].copy()

                if str(key).lower().find('tensão') > 0:
                    self.voltageDict[key] = dataCSV[key].copy()

            # print(self.voltageDict.keys())
            # print(self.currentDict.keys())


            self.emptyVoltageList = []
            self.emptyCurrentList = []

            for key, values in self.voltageDict.items():
                for index, value in enumerate(values):
                    if value == '' and index not in self.emptyVoltageList:
                        self.emptyVoltageList.append(index)

            for key, values in self.currentDict.items():
                for index, value in enumerate(values):
                    if value == '' and index not in self.emptyCurrentList:
                        self.emptyCurrentList.append(index)

            for index in sorted(self.emptyVoltageList, reverse=True):
                for key in self.voltageDict.keys():
                    self.voltageDict[key].pop(index)

            for index in sorted(self.emptyCurrentList, reverse=True):
                for key in self.currentDict.keys():
                    self.currentDict[key].pop(index)

            # print(self.voltageDict)
            # print(self.currentDict)

            print(
                sum(self.currentDict['RDO_09Z4 - Corrente A - A']) / len(self.currentDict['RDO_09Z4 - Corrente A - A']))
            print(
                sum(self.currentDict['RDO_09Z4 - Corrente B - A']) / len(self.currentDict['RDO_09Z4 - Corrente B - A']))
            print(
                sum(self.currentDict['RDO_09Z4 - Corrente C - A']) / len(self.currentDict['RDO_09Z4 - Corrente C - A']))

    def raise_error(self):

        self.LineEdit.setStyleSheet("color: red;")


if __name__ == '__main__':
    APP = QApplication(sys.argv)

    GUI = mainWindow()
    GUI.show()

    sys.exit(APP.exec())