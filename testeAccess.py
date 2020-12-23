import pyodbc 

driver = '{Microsoft Access Driver (*.mdb, *.accdb)}'
filepath = r'C:\Users\mathe\Desktop\Coding\Python\RedeTeste\Ano_0\Rede.mdb'
myDataSources = pyodbc.dataSources()

access_driver = myDataSources['MS Access Database']

password = 'dai365mon'
connection = pyodbc.connect(driver = access_driver,dbq=filepath,autocommit=True, PWD = password)
cursor = connection.cursor()

tables_list = list(cursor.tables())

# for table in tables_list:
#     print(table)

trafos = ['T35952','T37830','T111111','T38330']
trafosID = []
for trafo in trafos:
    cursor.execute('select CARGA_ID from CARGA where CODIGO=?',trafo)

    trafo_id = cursor.fetchone()
    if trafo_id:
        trafosID.append( str(trafo_id).split('(')[1].split(',')[0])

print(trafosID)

for trafoID in trafosID:
    cursor.execute('update MODELO_CARGA set PD=?,PE=?,PF=? where CARGA_ID=?',"1","2","3",trafoID)
    cursor.execute('update MODELO_CARGA set QD=?,QE=?,QF=? where CARGA_ID=?',"10","50","60",trafoID)

if len(trafosID) != int(cursor.rowcount):
    print(f'{int(cursor.rowcount) - len(trafosID) } trafo(s) não foram encontrado(s) ! ')
