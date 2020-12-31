import os
import pathlib
import platform
import sys
import csv
import xlrd
from PyQt5.QtWidgets import QStyleFactory, QApplication, QVBoxLayout,QHBoxLayout, QGroupBox, QPushButton, \
    QDialog, QGridLayout, QLineEdit, QComboBox, QFileDialog, QTableWidget, QTableWidgetItem, QLabel, QWidget, QTabWidget
from PyQt5.QtCore import Qt

import pyodbc 

# COMPILADOR EXECUTÁVEL
# python -m PyInstaller testeAccess.py --onefile

class mainWindow(QDialog):

    def __init__(self):


        super(mainWindow, self).__init__()

        self.setWindowTitle("Viabilidade AUX")
        self.setWindowModality(Qt.ApplicationModal)
        self.setStyle(QStyleFactory.create('Cleanlooks'))
        self.resize(400, 100)

        self.Dialog_Layout = QVBoxLayout()

        self.TabWidget = QTabWidget()
        self.CorrectTab = CorrectionTab()
        self.FindRefTab = FindReferenceTab()

        self.TabWidget.addTab(self.FindRefTab,"Encontrar Ref")
        self.TabWidget.addTab(self.CorrectTab,"Corrigir EPs")
        
        self.Dialog_Layout.addWidget(self.TabWidget)

        self.setLayout(self.Dialog_Layout)

class CorrectionTab(QWidget):
        def __init__(self):
            super().__init__()

            self.InitUI()

        def InitUI(self):
            self.dataCSV = {}

            self.GroupBox_Layout = QGridLayout()
            self.GroupBox_Layout.setAlignment(Qt.AlignLeft)

            self.UploadMDB_Button = QPushButton("Escolha uma base MDB ")
            self.UploadMDB_Button.clicked.connect(self.loadMDB)
            self.UploadMDB_Button.setFixedWidth(150)
            

            self.pathMDB_LineEdit = QLineEdit("")
            self.pathMDB_LineEdit.setReadOnly(True)
            self.pathMDB_LineEdit.setFixedWidth(150)

            self.UploadCSV_Button = QPushButton("Escolha um CSV/XLSX ")
            self.UploadCSV_Button.clicked.connect(self.load_csv)
            self.UploadCSV_Button.setFixedWidth(150)

            self.pathCSV_LineEdit = QLineEdit("")
            self.pathCSV_LineEdit.setReadOnly(True)
            self.pathCSV_LineEdit.setFixedWidth(150)



            self.ColSelect_Groupbox = QGroupBox('Colunas')
            self.ColSelect_Layout = QGridLayout()
            self.ColSelect_Layout.setAlignment(Qt.AlignCenter)

            self.ColTrafo_Label = QLabel('Trafo')
            self.ColTrafo_ComboBox = QComboBox()
            self.ColKW_Label = QLabel('Pot. kW')
            self.ColKW_ComboBox = QComboBox()
            self.ColSelect_Layout.addWidget(self.ColTrafo_Label, 1,1)
            self.ColSelect_Layout.addWidget(self.ColTrafo_ComboBox, 1,2)
            self.ColSelect_Layout.addWidget(self.ColKW_Label, 2,1)
            self.ColSelect_Layout.addWidget(self.ColKW_ComboBox, 2,2)
            self.ColSelect_Groupbox.setLayout(self.ColSelect_Layout)




            self.Run_Button = QPushButton("Corrigir")
            self.Run_Button.clicked.connect(self.runCorrection)
            self.Run_Button.setFixedWidth(150)


            self.tableWidget = QTableWidget()
            self.tableWidget.left = 0
            self.tableWidget.top = 0
            self.tableWidget.width = 1000
            self.tableWidget.height = 1000
            

            self.GroupBox_Layout.addWidget(self.UploadMDB_Button, 1,1,1,1)
            self.GroupBox_Layout.addWidget(self.pathMDB_LineEdit, 2,1,1,1)
            self.GroupBox_Layout.addWidget(self.UploadCSV_Button, 5,1,1,1)
            self.GroupBox_Layout.addWidget(self.pathCSV_LineEdit, 6,1,1,1)

            self.GroupBox_Layout.addWidget(self.ColSelect_Groupbox, 9,1,1,1)


            self.GroupBox_Layout.addWidget(self.Run_Button, 10,1)
            self.GroupBox_Layout.addWidget(self.tableWidget, 1,2,10,4)

            
            self.setLayout(self.GroupBox_Layout)

        def loadMDB(self):
            pathMDB = QFileDialog.getOpenFileName(self, 'Open MDB file',
                                                str(pathlib.Path.home()), "Access files (*.mdb)")

            if platform.system() == "Windows":
                pathMDB = pathMDB[0].replace('/', '\\')
            else:
                pathMDB = pathMDB[0]

            self.pathMDB_LineEdit.setText(pathMDB)

        def load_csv(self):

            self.ColTrafo_ComboBox.clear()
            self.ColKW_ComboBox.clear()

            pathCSV = QFileDialog.getOpenFileName(self, 'Open CSV/XLSX file',
                                                str(pathlib.Path.home()), "CSV files (*.csv;*.xlsx)")

            if platform.system() == "Windows":
                pathCSV = pathCSV[0].replace('/', '\\')
            else:
                pathCSV = pathCSV[0]

            self.pathCSV_LineEdit.setText(pathCSV)

            if pathCSV.split('.')[-1] == 'csv':
                with open(pathCSV, 'r', newline='') as file:
                    self.dataCSV.clear()
                    csv_reader_object = csv.reader(file)
                    name_col = next(csv_reader_object)

                    for row in name_col:
                        self.dataCSV[row] = []

                    for row in csv_reader_object:  ##Varendo todas as linhas
                        for ndata in range(0, len(name_col)):  ## Varendo todas as colunas
                            self.dataCSV[name_col[ndata]].append(row[ndata])
            #  print(self.dataCSV)

            elif pathCSV.split('.')[-1] == 'xlsx':
                self.dataCSV.clear()
                workbook = xlrd.open_workbook(pathCSV, on_demand=True)
                worksheet = workbook.sheet_by_index(0)
                first_row = []  # The row where we stock the name of the column

                for col in range(worksheet.ncols):
                    first_row.append(worksheet.cell_value(0, col))

                for col in first_row:
                    self.dataCSV[col] = []

                for row in range(1, worksheet.nrows):
                    for col in range(0, worksheet.ncols):  ## Varendo todas as colunas
                        self.dataCSV[first_row[col]].append(worksheet.cell_value(row, col))
                
            #  print(self.dataCSV)

            rowsLen = len(list(self.dataCSV.values())[0]) 
            colLen = len(list(self.dataCSV.keys()))
            self.tableWidget.setRowCount(rowsLen)
            self.tableWidget.setColumnCount(colLen)


            index=0
            for key in self.dataCSV.keys():
                self.ColTrafo_ComboBox.addItem(str(key),index)
                self.ColKW_ComboBox.addItem(str(key),index)
                index+=1

            self.ColTrafo_ComboBox.setCurrentIndex(0)
            self.ColKW_ComboBox.setCurrentIndex(1)

            j=0
            for key,values in self.dataCSV.items():
                i=0
                self.tableWidget.setHorizontalHeaderItem(j,QTableWidgetItem(str(key)))
                for value in values:
                    self.tableWidget.setItem(i,j, QTableWidgetItem(str(value)))
                    i+=1
                j+=1

            self.adjustSize()

        def runCorrection(self):
            pathMDB=""
            pathCSV=""
            if self.pathMDB_LineEdit.text():
                pathMDB = self.pathMDB_LineEdit.text()
            
            if self.pathCSV_LineEdit.text():
                pathCSV = self.pathCSV_LineEdit.text()

            #print(pathMDB,pathCSV)

            driver = '{Microsoft Access Driver (*.mdb, *.accdb)}'
            myDataSources = pyodbc.dataSources()
            access_driver = myDataSources['MS Access Database']
            password = 'dai365mon'

            connection = pyodbc.connect(driver = access_driver,dbq=pathMDB,autocommit=True, PWD = password)
            cursor = connection.cursor()
        
            ColTrafo = self.ColTrafo_ComboBox.currentText()
            ColKW = self.ColKW_ComboBox.currentText()
            
            trafos = self.dataCSV[ColTrafo]
            #print(self.dataCSV[ColTrafo])

            demandas = self.dataCSV[ColKW]
            #print(self.dataCSV[ColKW])


            self.trafoDict = {}

            for trafo,demanda in zip(trafos,demandas):
                cursor.execute('select CARGA_ID from CARGA where CODIGO=?',trafo)
                trafo_id = cursor.fetchone()
                if trafo_id:
                    self.trafoDict[trafo] = [str(trafo_id).split('(')[1].split(',')[0],demanda]
                else:
                    self.trafoDict[trafo] = ['',demanda]

            print(self.trafoDict)

            for key,value in self.trafoDict.items():
                trafoID = value[0]
                demandaKW = float(value[1])
                kva = demandaKW/0.92
                demandaKVAr = kva*0.39191835

                cursor.execute('update MODELO_CARGA set PD=?,PE=?,PF=?,QD=?,QE=?,QF=? where CARGA_ID=?'
                ,demandaKW/3,demandaKW/3,demandaKW/3
                ,demandaKVAr/3,demandaKVAr/3,demandaKVAr/3,
                trafoID)

            if len(trafos) != int(cursor.rowcount):
                print(f'{int(cursor.rowcount) - len(trafos) } trafo(s) não foram encontrado(s) ! ')

class FindReferenceTab(QWidget):
    def __init__(self):
        super().__init__()

        self.InitUI()

    def InitUI(self):
        self.GroupBox_Layout = QGridLayout()
        self.GroupBox_Layout.setAlignment(Qt.AlignCenter)


        self.MDBGroupBox = QGroupBox()
        self.MDBGroupBox_Layout = QHBoxLayout()
        self.UploadMDB_Button = QPushButton("Escolha uma base MDB")
        self.UploadMDB_Button.clicked.connect(self.loadMDB)
        self.UploadMDB_Button.setFixedWidth(150)
        
        self.pathMDB_LineEdit = QLineEdit("")
        self.pathMDB_LineEdit.setReadOnly(True)
        self.pathMDB_LineEdit.setFixedWidth(150)
        self.MDBGroupBox_Layout.addWidget(self.UploadMDB_Button)
        self.MDBGroupBox_Layout.addWidget(self.pathMDB_LineEdit)
        self.MDBGroupBox.setLayout(self.MDBGroupBox_Layout)



        self.RefGroupBox = QGroupBox()
        self.RefGroupBox_Layout = QGridLayout()
        self.RefGroupBox_Layout.setAlignment(Qt.AlignCenter)
        self.Reference_Label = QLabel("Referência:")
        self.Reference_LineEdit = QLineEdit("")
        self.Reference_LineEdit.setFixedWidth(150)
        self.FoundCircuit_Label = QLabel("Alimentador:")
        self.FoundCircuit_LineEdit = QLineEdit("")
        self.FoundCircuit_LineEdit.setReadOnly(True)
        self.FoundCircuit_LineEdit.setFixedWidth(150)
        self.ReferenceEquipament_Label = QLabel("Equipamento:")
        self.ReferenceEquipament_LineEdit = QLineEdit("")
        self.ReferenceEquipament_LineEdit.setReadOnly(True)
        self.ReferenceEquipament_LineEdit.setFixedWidth(150)
        self.ReferenceStatus_Label = QLabel("Estado Normal/Atual:")
        self.ReferenceStatus_LineEdit = QLineEdit("")
        self.ReferenceStatus_LineEdit.setReadOnly(True)
        self.ReferenceStatus_LineEdit.setFixedWidth(150)

        self.RefGroupBox_Layout.addWidget(self.Reference_Label,0,0)
        self.RefGroupBox_Layout.addWidget(self.Reference_LineEdit,0,1)
        self.RefGroupBox_Layout.addWidget(self.FoundCircuit_Label,0,2)
        self.RefGroupBox_Layout.addWidget(self.FoundCircuit_LineEdit,0,3)
        self.RefGroupBox_Layout.addWidget(self.ReferenceEquipament_Label,1,2)
        self.RefGroupBox_Layout.addWidget(self.ReferenceEquipament_LineEdit,1,3)
        self.RefGroupBox_Layout.addWidget(self.ReferenceStatus_Label,2,2)
        self.RefGroupBox_Layout.addWidget(self.ReferenceStatus_LineEdit,2,3)
        self.RefGroupBox.setLayout(self.RefGroupBox_Layout)

        self.Find_Button = QPushButton("Procurar")
        self.Find_Button.clicked.connect(self.FindReference)
        self.Find_Button.setFixedWidth(150)

        self.GroupBox_Layout.addWidget(self.MDBGroupBox,2,0)
        self.GroupBox_Layout.addWidget(self.RefGroupBox,3,0)
        self.GroupBox_Layout.addWidget(self.Find_Button,4,0)

        self.GroupBox_Layout.setAlignment(Qt.AlignCenter)
        self.setLayout(self.GroupBox_Layout)

    
    ## TEMPORÁRIO, TORNAR ESSA FUNÇÃO GLOBAL
    def loadMDB(self): 
        pathMDB = QFileDialog.getOpenFileName(self, 'Open MDB file',
                                            str(pathlib.Path.home()), "Access files (*.mdb)")

        if platform.system() == "Windows":
            pathMDB = pathMDB[0].replace('/', '\\')
        else:
            pathMDB = pathMDB[0]

        self.pathMDB_LineEdit.setText(pathMDB)

    def FindReference(self):
        pathMDB=""
        if self.pathMDB_LineEdit.text():
            pathMDB = self.pathMDB_LineEdit.text()

        driver = '{Microsoft Access Driver (*.mdb, *.accdb)}'
        myDataSources = pyodbc.dataSources()
        access_driver = myDataSources['MS Access Database']
        password = 'dai365mon'

        connection = pyodbc.connect(driver = access_driver,dbq=pathMDB,autocommit=True, PWD = password)
        cursor = connection.cursor()

        reference = self.Reference_LineEdit.text().split(' ')[0].split('\n')[0].upper()
        
        self.FoundCircuit_LineEdit.clear()
        self.ReferenceEquipament_LineEdit.clear()
        self.ReferenceStatus_LineEdit.clear()

        # print(reference)
        refTypes = ['CARGA', 'BARRA','CHAVE','REGULADOR','REATOR','RAIOX','CAPACITOR']

        for ref in refTypes:
            cursor.execute('select CIRC_ID from ' + ref + ' where CODIGO=?',reference)
            circuit_id = cursor.fetchone() 
            if circuit_id :
                circuit_id = str(circuit_id).split('(')[1].split(',')[0]
                equipament = ref
                break

        cursor.execute('select CODIGO from CIRCUITO where CIRC_ID=?', circuit_id)

        circuit_name = cursor.fetchone()

        if circuit_id and int(circuit_id)>0:
            circuit_name = str(circuit_name).split('(')[1].split(',')[0].split('\'')[1]
            # print(circuit_name)
            self.FoundCircuit_LineEdit.setText(circuit_name)
            self.FoundCircuit_LineEdit.setStyleSheet("color : blue")
            self.ReferenceEquipament_LineEdit.setText(equipament)
            if equipament== "CHAVE":
                cursor.execute('select ESTADO from CHAVE where CODIGO=?',reference)
                if str(cursor.fetchone()).split('(')[1].split(',')[0] == '1':
                    currentState = 'Fechado'
                else:
                    currentState = 'Aberto'

                cursor.execute('select ESTADO_NORMAL from CHAVE where CODIGO=?',reference)
                if str(cursor.fetchone()).split('(')[1].split(',')[0] == '1':
                    normalState = 'Fechado'
                else:
                    normalState = 'Aberto'

                self.ReferenceStatus_LineEdit.setText(normalState + '/' + currentState)
                
        else:
            self.FoundCircuit_LineEdit.setText("Not Found")
            self.FoundCircuit_LineEdit.setStyleSheet("color : red")

if __name__ == '__main__':
    APP = QApplication(sys.argv)
    APP.setStyle('Fusion')

    GUI = mainWindow()
    GUI.show()

    sys.exit(APP.exec())