import datetime
import numpy as np
import pandas as pd

"""def convert(date_time): 
	format = '%d/%m/%y' # The format 
	datetime_str = datetime.datetime.strptime(date_time, format) 
	datetime_str = datetime_str.date()
	return datetime_str"""

def convert(text):
	for fmt in ('%d/%m/%y', '%d/%m/%Y'):
		try:
			return datetime.datetime.strptime(text, fmt).date()
		except ValueError:
			pass
	raise ValueError('no valid date format found')

date_deb = input("Entrer une date début: ")
print(f'La date début est {convert(date_deb)}\n')

while True:             
	date_fin = input("Entrer une date fin: ") 
	if (convert(date_fin) >= convert(date_deb)):
		print(f'La date fin est {convert(date_fin)}\n')      
		break  

data = pd.read_excel('input1_old.xlsx', header = 0, encoding = "ISO-8859-1", error_bad_lines=False, warn_bad_lines=False)

del data['CO_No']

data2 = pd.read_excel('input1a.xlsx', header = 0, encoding = "ISO-8859-1", error_bad_lines=False, warn_bad_lines=False)

data = data.append(data2, ignore_index = True)
data.sort_values(by=['AN'])

pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)

del data['INTITULE CLIENT']
del data['N° CLIENT']
del data['Marge']

#print(data)
table_A = data.copy()

table_A['Somme'] = [0.0] * len(table_A.index)
table_A['RAP_p'] = [0.0] * len(table_A.index)
 
somme = 0
#for y in range(len(table_A.index)):
for index, row in table_A.iterrows():
	somme = row['DR_Montant'] + row['montant']
	#print(somme)
	table_A.at[index, 'Somme'] = somme
	table_A.at[index, 'RAP_p'] = row['Total TTC'] - somme
 

table_A['réglements'] = [0.0] * len(table_A.index)
table_A['RAP'] = [0.0] * len(table_A.index)


table_A = table_A[pd.notnull(table_A['Date Regl'])]
#table_A = table_A.fillna(0)

for index, row in table_A.iterrows():
	if ((row['Date Regl'] < convert(date_deb) and row['RAP_p'] <= 0) or (row['Date Regl'] > convert(date_fin) and row['RAP_p'] > 0)):
		table_A.at[index, 'réglements'] = -1
		table_A.at[index, 'RAP'] = -1
	#elif (row['Date Regl'] <= convert(date_fin)):
	else:
		table_A.at[index, 'réglements'] = row['Somme']
		table_A.at[index, 'RAP'] = row['Total TTC'] - row['réglements']

table_A = table_A.drop(columns =['DR_Montant', 'montant', 'Date Regl', 'Somme', 'RAP_p'])

table_A.to_excel("table_A.xlsx")

print(table_A)



