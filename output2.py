import datetime
import numpy as np
import pandas as pd
import re

isValid=False
while not isValid:
	userIn = input("Entrer une date début: ")
	try:
		date_deb = datetime.datetime.strptime(userIn, "%d/%m/%y").date()
		isValid=True
	except:
		print("Suivre ce format: dd/mm/yy !\n")

print(f'La date début est {date_deb}\n')

isValid = False
while not isValid:
	userIn = input("Entrer une date fin: ")
	try:
		date_fin = datetime.datetime.strptime(userIn, "%d/%m/%y").date()
		if (date_fin > date_deb):
			isValid = True
		else:
			print("Erreur: date fin < date debut !\n")
	except:
		print("Suivre ce format: dd/mm/yy !\n")
print(f'La date fin est {date_fin}\n')

#data = pd.read_excel('input1_old.xlsx', header = 0, encoding = "ISO-8859-1", error_bad_lines=False, warn_bad_lines=False)
data = pd.read_excel('CA_MS.xlsx', header = 0)

del data['CO_No']

data2 = pd.read_excel('CA_MSMARINE.xlsx', header = 0)
#data2 = pd.read_excel('input1a.xlsx', header = 0, encoding = "ISO-8859-1", error_bad_lines=False, warn_bad_lines=False)

data = data.append(data2, ignore_index = True)
data.sort_values(by=['AN'])

pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)

del data['INTITULE CLIENT']
del data['N° CLIENT']
del data['Marge']

data[['DR_sum', 'mt_sum']]=data.groupby('DOC')['DR_Montant', 'montant'].transform('sum')

data = data.sort_values('Date Regl').drop_duplicates('DOC', keep='last')
#data = data.groupby('DOC')['Date Regl'].max().reset_index()

#data = data.groupby(['DOC'])['montant'].sum().reset_index()

data['réglements'] = data['DR_sum'] + data['mt_sum']
data['RAP'] = data['Total TTC'] - data['réglements']

data = data.drop(['DR_Montant', 'montant', 'DR_sum', 'mt_sum'], axis = 1)

data.to_excel("TABLEAU A.xlsx")

print("TABLEAU A: \n")
print(data)
print("\n\n")

for index, row in data.iterrows():
	if (row['Date Regl'] < date_deb or row['Date Regl'] > date_fin):
		data.drop(index, inplace=True)

for index, row in data.iterrows():
	if (row['RAP'] > 0):
		data.drop(index, inplace=True)

print('TABLEAU B: \n')
print(data)
data.to_excel("TABLEAU B.xlsx")

table5 = pd.read_excel('table5.xlsx', header = 0)
#table5 = pd.read_excel('table5.xlsx', header = 0, encoding = "ISO-8859-1", error_bad_lines=False, warn_bad_lines=False)

ran = 2025 - 2016 + 1
ind = range(2016, 2016 + ran)
d = {	
    'CA HT periode':[0.0] * ran,
    'CA TTC période':[0.0] * ran,
    'total réglements période':[0.0] * ran,
    'total RAP':[0.0] * ran, 
    'comms période à payer':[0.0] * ran 
}
BS = pd.DataFrame(d, columns = ['CA HT periode' , 'CA TTC période', 'total réglements période' , 'total RAP' , 'comms période à payer'], index = list(ind))

for y in ind:
	HT = 0
	TTC = 0
	regl = 0
	RAP = 0
	for index, row in data.iterrows():
		if ((row['AN'] == y) and (row['Representant'] == 'Bejaoui Sahbi' or row['Representant'] == 'Sahbi Bejaoui REV')):
			HT = HT + row['Total HT']
			TTC = TTC + row['Total TTC']
			regl = regl + row['réglements']
			RAP = RAP + row['RAP']
	BS.at[y, 'CA HT periode'] = HT
	BS.at[y, 'CA TTC période'] = TTC
	BS.at[y, 'total réglements période'] = regl
	BS.at[y, 'total RAP'] = RAP

SA = pd.DataFrame(d, columns = ['CA HT periode' , 'CA TTC période', 'total réglements période' , 'total RAP' , 'comms période à payer'], index = list(ind))

for y in ind:
	HT = 0
	TTC = 0
	regl = 0
	RAP = 0
	for index, row in data.iterrows():
		if ((row['AN'] == y) and (row['Representant'] == 'Saidi Abdelkarim' or row['Representant'] == 'Abdelkarim Saidi REV')):
			HT = HT + row['Total HT']
			TTC = TTC + row['Total TTC']
			regl = regl + row['réglements']
			RAP = RAP + row['RAP']
	SA.at[y, 'CA HT periode'] = HT
	SA.at[y, 'CA TTC période'] = TTC
	SA.at[y, 'total réglements période'] = regl
	SA.at[y, 'total RAP'] = RAP

#print(table5)

BS['com %'] = [0.0] * ran
SA['com %'] = [0.0] * ran
x = 2016
for index, row in table5.iterrows():
	BS.at[x, 'com %'] = row['Bejaoui Sahbi']
	SA.at[x, 'com %'] = row['Saidi Abdelkarim']
	x = x + 1

for index, row in BS.iterrows():
	BS.at[index, 'comms période à payer'] = row['com %'] * row['CA HT periode']

for index, row in SA.iterrows():
	SA.at[index, 'comms période à payer'] = row['com %'] * row['CA HT periode']

BS = BS.append(BS.sum().rename('Total'))

SA = SA.append(SA.sum().rename('Total'))

del BS['com %']
del SA['com %']

print(f'commissions Sahbi Bejaoui à payer, période: {date_deb} - {date_fin}\n')
print(BS)
print('\n')

print(f'commissions abdeelkrim saidi à payer, période: {date_deb} - {date_fin}\n') 
print(SA)

SA.to_excel("TABLEAU D.xlsx")
BS.to_excel("TABLEAU C.xlsx")
