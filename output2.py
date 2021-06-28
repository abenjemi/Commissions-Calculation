import datetime
import numpy as np
import pandas as pd
import math
import re
from tkinter import *
from tkinter import messagebox

global e1
global e2
global button1
global button2

def deb_fun():
    global debut
    global date_deb
    debut = e1.get()
    #date_deb = e1.get()

    if (debut == ""):
        messagebox.showerror("", "Veuillez entrer une date debut")

    else:
        try:
            #date_deb = datetime.datetime.strptime(debut, "%d/%m/%y").strftime('%d/%m/%y')
            date_deb = datetime.datetime.strptime(debut, "%d/%m/%y")
        except:
            messagebox.showerror("", "Veuillez suivre ce format: jj/mm/aa")
            e1.delete(0,"end")
        #print(date_deb)
        label3 = Label(root, text="La date debut est %s" %(date_deb.strftime('%d/%m/%y')), font=("Arial", 13))
        label3.place(x=10, y=50)
        e1.delete(0,"end")
        e1.config(state='disabled')
        e2.config(state='normal')
        button1.config(state='disabled')
        button2.config(state='normal')
        #date_debut = date_deb
    
def fin_fun():
    global fin
    global date_fin
    fin = e2.get()

    if (fin == ""):
        messagebox.showerror("", "Veuillez entrer une date fin")
    else:
        try:
            #date_fin = datetime.datetime.strptime(fin, "%d/%m/%y").strftime('%d/%m/%y')
            date_fin = datetime.datetime.strptime(fin, "%d/%m/%y")
            if (date_fin < date_deb):
                messagebox.showerror("", "Erreur: date fin < date debut")
                #e2.delete(0,"end")
                return
        except:
            messagebox.showerror("", "Veuillez suivre ce format: dd/mm/yy")
            e2.delete(0,"end")
            return
        label4 = Label(root, text="La date fin est %s" %(date_fin.strftime('%d/%m/%y')), font=("Arial", 13))
        label4.place(x=10, y=130)
        e2.config(state='disabled')
        button2.config(state='disabled')
        messagebox.showinfo("Output 2", "Veuillez cliquer pour voir les resultats suivants sur le meme dossier:\n1. etat_reglements_factures\n2. etat_factures_reglees100%_période_saisie\n3. Commissions_BejaouiS_periode_saisie\n4. Commissions_SaidiA_periode_saisie")
        root.destroy()




root = Tk()
root.title("Entrer date debut et date fin")
root.geometry("450x200")

label1 = Label(root, text="Veuillez entrer une date début ", font=("Arial", 13))
label2 = Label(root, text="Veuillez entrer une date fin ", font=("Arial", 13))

label1.place(x=10, y=10)
label2.place(x=10, y=90)

e1 = Entry(root)
e1.place(x=240, y = 10)

button1 = Button(root, text="Cliquer", command=deb_fun)
button1.place(x=375, y=10)
#button1.bind("<Return>", deb)

e2 = Entry(root, state='disabled')
e2.place(x=220, y=90)

button2 = Button(root, text="Confirmer", state='disabled', command=fin_fun)
button2.place(x=375, y=90)



root.mainloop()


#data = pd.read_excel('input1_old.xlsx', header = 0, encoding = "ISO-8859-1", error_bad_lines=False, warn_bad_lines=False)
data = pd.read_excel('MS_M_CalculComms.xlsx', header = 0)

del data['CO_No']

data2 = pd.read_excel('MSMARINE_CalculComms.xlsx', header = 0)
#data2 = pd.read_excel('input1a.xlsx', header = 0, encoding = "ISO-8859-1", error_bad_lines=False, warn_bad_lines=False)

data = data.append(data2, ignore_index = True)
data.sort_values(by=['AN'])

pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)

#del data['INTITULE CLIENT']
#del data['N° CLIENT']
#del data['Marge']

#data[['DR_sum', 'mt_sum']]=data.groupby('DOC')['DR_Montant', 'montant'].transform('sum')

data = data.sort_values('Date_Regl')
#data.groupby(['DOC'])['Montant_Reglement'].sum().reset_index()#.drop_duplicates('DOC', keep='last')

data['Montant_Reglement'] = data.groupby(['DOC'])['Montant_Reglement'].transform('sum')

data = data.drop_duplicates('DOC', keep='last')

#data = data.groupby('DOC')['Date Regl'].max().reset_index()

#data = data.groupby(['DOC'])['Montant_Reglement'].sum()

#data['réglements'] = data['Montant_Reglement']
data['RAP'] = data['Total TTC'] - data['Montant_Reglement']
data['RAP'] = data['RAP'].apply(np.floor)
#math.floor(data['RAP'])
#if ((data['RAP'] < 1) and (data['RAP'] > 0)):
#    data['RAP'] =  0


#data = data.drop(['DR_Montant', 'montant', 'DR_sum', 'mt_sum'], axis = 1)
#data = data.drop(['Montant_Reglement'])

data.to_excel("etat_reglements_factures.xlsx")

#print("TABLEAU A: \n")
#print(data)
#print("\n\n")

for index, row in data.iterrows():
	if (row['Date_Regl'] < date_deb or row['Date_Regl'] > date_fin):
		data.drop(index, inplace=True)

# for index, row in data.iterrows():
# 	if (row['RAP'] > 0.5):
# 		data.drop(index, inplace=True)
data.drop(data[data['RAP'] > 0].index, inplace = True)
# index_names = data[ data['Age'] == 21 ].index
# df.drop(index_names, inplace = True)

#print('TABLEAU B: \n')
#print(data)
data.to_excel("etat_factures_reglees100%_période_saisie.xlsx")

table5 = pd.read_excel('rapport_commissions_CA.xlsx', header = 0)
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
			regl = regl + row['Montant_Reglement']
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
			regl = regl + row['Montant_Reglement']
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

#print(f'commissions Sahbi Bejaoui à payer, période: {date_deb} - {date_fin}\n')
#print(BS)
#print('\n')

#print(f'commissions abdeelkrim saidi à payer, période: {date_deb} - {date_fin}\n') 
#print(SA)

SA.to_excel("Commissions_SaidiA_periode_saisie.xlsx")
BS.to_excel("Commissions_BejaouiS_periode_saisie.xlsx")
