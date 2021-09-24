import pandas as pd
import pyodbc


# for driver in pyodbc.drivers():
#     print(driver)

# conn = pyodbc.connect('Driver={ODBC Driver 17 for SQL Server};'
#                       'Server=192.168.10.14;'
#                       'Database=MSM;'
#                       'username = SDEV;'
#                         'password = test123123;')


server = '192.168.10.14' 
database = 'MSM'  

cnxn = pyodbc.connect(
    'DRIVER={SQL Server Native Client 11.0}; \
    SERVER='+ server +'; \
    DATABASE='+ database +';\
    uid=SDEV;\
    pwd=test123123;'
)

cursor = cnxn.cursor()

data = pd.read_sql_query('SELECT TOP (200) AN, DOC, [DATE DOC], [INTITULE CLIENT], [N° Client], [Total HT], [Total TTC], Marge, Representant, CO_No, Montant_Reglement, Date_Regl FROM MS_M_CalculComms',cnxn)
print(data)
print(type(data))






database = 'MSMARINE'  

cnxn = pyodbc.connect(
    'DRIVER={SQL Server Native Client 11.0}; \
    SERVER='+ server +'; \
    DATABASE='+ database +';\
    uid=SDEV;\
    pwd=test123123;'
)

cursor = cnxn.cursor()

data = pd.read_sql_query('SELECT TOP (200) AN, DOC, [DATE DOC], [INTITULE CLIENT], [N° Client], [Total HT], [Total TTC], Marge, Representant, CO_No, Montant_Reglement, Date_Regl FROM MSMARINE_CalculComms',cnxn)
print(data)
print(type(data))


