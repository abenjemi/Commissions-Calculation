import numpy as np
import pandas as pd
from tkinter import *

root = Tk()
root.title("Output 1")

# Create a label widget
myLabel = Label(root, text="Veuillez cliquer pour voir les fichiers excel suivants sur le meme dossier:").grid(row=0, column=0)
fichier1 = Label(root, text="1. CA_BejaouiS_SaidiA").grid(row=1, column=0)
fichier2 = Label(root, text="2. Commissions_BejaouiS_SaidiA_details").grid(row=2, column=0)
fichier3 = Label(root, text="3. Commissions_BejaouiS_SaidiA_total").grid(row=3, column=0)
fichier4 = Label(root, text="4. rapport_commissions_CA").grid(row=4, column=0)

root.resizable(width=False, height=False)

myButton = Button(root, text="Cliquez SVP", command=root.destroy, fg="blue", bg="white").grid(row=6)


root.mainloop()

data = pd.read_excel('MS_M_CalculComms.xlsx', header = 0)

del data['CO_No']

data2 = pd.read_excel('MSMARINE_CalculComms.xlsx', header = 0)

data = data.append(data2, ignore_index = True) #get all the data

data.drop_duplicates(subset = ["DOC"], keep="first", inplace=True) #drop duplicates
data.sort_values(by=['AN'])
pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)
data.to_excel("data.xlsx") #get the duplicates free table on excel

#2
ran = 2025 - 2016 + 1
ind = range(2016, 2016 + ran)
d = {	
    'Bejaoui Sahbi':[0.0] * ran,
    'Sahbi Bejaoui REV':[0.0] * ran,
    'Saidi Abdelkarim':[0.0] * ran,
    'Abdelkarim Saidi REV':[0.0] * ran  
}
df = pd.DataFrame(d, columns = ['Bejaoui Sahbi' , 'Sahbi Bejaoui REV', 'Saidi Abdelkarim' , 'Abdelkarim Saidi REV'], index = list(ind))

rpr = list(df.columns)
for y in ind:
	for rep in rpr:
		summ = 0
		for index, row in data.iterrows():
			if (row['AN'] == y and row['Representant'] == rep): 
				summ = summ + row['Total HT']
		df.at[y, rep] = summ

df.to_excel("CA_BejaouiS_SaidiA.xlsx") #CA

#3
input2_BS = pd.read_excel('Objectifs_BejaouiS.xlsx', header = 0)

input2_BS = input2_BS.fillna(0)
input2_BS['CA'] = [0.0] * ran
input2_BS['Bejaoui Sahbi'] = [0.0] * ran
input2_BS['Sahbi Bejaoui REV'] = [0.0] * ran

BS = df["Bejaoui Sahbi"]
SB_REV = df["Sahbi Bejaoui REV"]

input2_SA = pd.read_excel('Objectifs_SaidiA.xlsx', header = 0)

input2_SA = input2_SA.fillna(0)
input2_SA['CA'] = [0.0] * ran
input2_SA['Saidi Abdelkarim'] = [0.0] * ran
input2_SA['Abdelkarim Saidi REV'] = [0.0] * ran

SA = df["Saidi Abdelkarim"]
AS_REV = df["Abdelkarim Saidi REV"]

x = 0
for index, row in df.iterrows():
	input2_BS.at[x, 'CA'] = row[rpr[0]] + row[rpr[1]]
	input2_BS.at[x, 'Bejaoui Sahbi'] = row[rpr[0]]
	input2_BS.at[x, 'Sahbi Bejaoui REV'] = row[rpr[1]]
	input2_SA.at[x, 'CA'] = row[rpr[2]] + row[rpr[3]]
	input2_SA.at[x, 'Saidi Abdelkarim'] = row[rpr[2]]
	input2_SA.at[x, 'Abdelkarim Saidi REV'] = row[rpr[3]]
	x = x + 1	
#print('CA SAHBI\n\n')
#print(input2_BS)		

table3 = df.copy()
table3.insert(2, "Bejaoui Sahbi EX", [0.0] * ran, True)
table3.insert(5, "Saidi Abdelkarim EX", [0.0] * ran, True)

x = 2016
for index, row in input2_BS.iterrows():
	if (row['Bejaoui Sahbi'] > row['objectif min']):
		table3.at[x, 'Bejaoui Sahbi'] = (row['Bejaoui Sahbi'] - row['objectif min']) * row['% VD']
		table3.at[x, 'Sahbi Bejaoui REV'] = row['Sahbi Bejaoui REV'] * row['% VR']
	else:
		table3.at[x, 'Bejaoui Sahbi'] = 0
		table3.at[x, 'Sahbi Bejaoui REV'] = 0
	if (row['CA'] > row['objectif excellence']):
		table3.at[x, 'Bejaoui Sahbi EX'] = row['CA'] * row['% EX']	
	else:
		table3.at[x, 'Bejaoui Sahbi EX'] = 0
	x = x + 1

z = 2016
for index, row in input2_SA.iterrows():
	if (row['Saidi Abdelkarim'] > row['objectif min']):
		table3.at[z, 'Saidi Abdelkarim'] = (row['Saidi Abdelkarim'] - row['objectif min']) * row['% VD']
		table3.at[z, 'Abdelkarim Saidi REV'] = row['Abdelkarim Saidi REV'] * row['% VR']
	else:
		table3.at[z, 'Saidi Abdelkarim'] = 0
		table3.at[z, 'Abdelkarim Saidi REV'] = 0
	if (row['CA'] > row['objectif excellence']):
		table3.at[z, 'Saidi Abdelkarim EX'] = row['CA'] * row['% excellence']	
	else:
		table3.at[z, 'Saidi Abdelkarim EX'] = 0
	z = z + 1

table3.to_excel("Commissions_BejaouiS_SaidiA_details.xlsx") #tableau 3
#print ('Result 3:\n')
#print(table3)
#print('\n\n')

#4
f = {
    'Bejaoui Sahbi':[0.0] * ran,
    'Saidi Abdelkarim':[0.0] * ran  
}
table4 = pd.DataFrame(f, columns = ['Bejaoui Sahbi' , 'Saidi Abdelkarim'], index=list(ind))

for index, row in table3.iterrows():
	table4.at[index, 'Bejaoui Sahbi'] = row['Bejaoui Sahbi'] + row['Sahbi Bejaoui REV'] + row['Bejaoui Sahbi EX']
	table4.at[index, 'Saidi Abdelkarim'] = row['Saidi Abdelkarim'] + row['Abdelkarim Saidi REV'] + row['Saidi Abdelkarim EX']

table4.to_excel("Commissions_BejaouiS_SaidiA_total.xlsx")
#print('Result 4: \n')
#print(table4)

#5

table5 = table4.copy()
table4['CA BS'] = [0.0] * ran
table4['CA SA'] = [0.0] * ran

for index, row in df.iterrows():
	table4.at[index, 'CA BS'] = row[rpr[0]] + row[rpr[1]]
	table4.at[index, 'CA SA'] = row[rpr[2]] + row[rpr[3]]

for index, row in table4.iterrows():
	if(row['CA BS'] != 0):
		table5.at[index, 'Bejaoui Sahbi'] = row['Bejaoui Sahbi'] / row['CA BS']
	else:
		table5.at[index, 'Bejaoui Sahbi'] = 0
	if(row['CA SA'] != 0):
		table5.at[index, 'Saidi Abdelkarim'] = row['Saidi Abdelkarim'] / row['CA SA']
	else:
		table5.at[index, 'Saidi Abdelkarim'] = 0
	
#print('Result 5: \n')
#print(table5)

table5.to_excel("rapport_commissions_CA.xlsx")

