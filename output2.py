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
#pd.set_option('display.float_format', '{:,.0f}'.format)

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

orig_data = pd.read_excel('data.xlsx', header = 0)

table4 = pd.read_excel('Commissions_BejaouiS_SaidiA_total.xlsx', header = 0)

Obj_BS = pd.read_excel('Objectifs_BejaouiS.xlsx', header = 0)

Obj_BS.set_index('AN', inplace=True, drop=True)

del Obj_BS['charges']

Obj_SA = pd.read_excel('Objectifs_SaidiA.xlsx', header = 0)

Obj_SA.set_index('AN', inplace=True, drop=True)

del Obj_SA['charges']

ran = 2025 - 2016 + 1
ind = range(2016, 2016 + ran)
d = {	
    'factures HT année':[0.0] * ran,
    'factures TTC année':[0.0] * ran,
    'factures réglées période':[0.0] * ran,
    'pourcentage règlement':[0.0] * ran, 
    'commissions année':[0.0] * ran,
    'commission période':[0.0] * ran,
    'commission RAP':[0.0] * ran
}
BS = pd.DataFrame(d, columns = ['factures HT année', 'factures TTC année', 'factures réglées période', 'pourcentage règlement', 'commissions année', 'commission période', 'commission RAP'], index = list(ind))

for y in ind:
    HT = 0
    TTC = 0
    for index, row in orig_data.iterrows():
        if ((row['AN'] == y) and (row['Representant'] == 'Bejaoui Sahbi' or row['Representant'] == 'Sahbi Bejaoui REV')):
            HT = HT + row['Total HT']
            TTC = TTC + row['Total TTC']
    regl = 0
    PR = 0
    for index, row in data.iterrows():
        if ((row['AN'] == y) and (row['Representant'] == 'Bejaoui Sahbi' or row['Representant'] == 'Sahbi Bejaoui REV')):
            regl = regl + row['Total TTC']
    BS.at[y, 'factures réglées période'] = regl
    if (TTC == 0):
        PR = 0
    else:
        PR = (regl / TTC) * 100
    BS.at[y, 'pourcentage règlement'] = PR
    
    BS.at[y, 'factures HT année'] = HT
    BS.at[y, 'factures TTC année'] = TTC
     

SA = pd.DataFrame(d, columns = ['factures HT année', 'factures TTC année', 'factures réglées période', 'pourcentage règlement', 'commissions année', 'commission période', 'commission RAP'], index = list(ind))

for y in ind:
    HT = 0
    TTC = 0
    for index, row in orig_data.iterrows():
        if ((row['AN'] == y) and (row['Representant'] == 'Saidi Abdelkarim' or row['Representant'] == 'Abdelkarim Saidi REV')):
            HT = HT + row['Total HT']
            TTC = TTC + row['Total TTC']
    SA.at[y, 'factures HT année'] = HT
    SA.at[y, 'factures TTC année'] = TTC
    regl = 0
    PR = 0
    for index, row in data.iterrows():
        if ((row['AN'] == y) and (row['Representant'] == 'Saidi Abdelkarim' or row['Representant'] == 'Abdelkarim Saidi REV')):
            regl = regl + row['Total TTC']
    SA.at[y, 'factures réglées période'] = regl
    if (TTC == 0):
        PR = 0
    else:
        PR = (regl / TTC) * 100
    SA.at[y, 'pourcentage règlement'] = PR


#print(table4)

BS['com %'] = [0.0] * ran
SA['com %'] = [0.0] * ran
x = 2016
for index, row in table4.iterrows():
	BS.at[x, 'com %'] = row['Bejaoui Sahbi']
	SA.at[x, 'com %'] = row['Saidi Abdelkarim']
	x = x + 1

#print(BS)

CAn = 0
CP = 0
RAP = 0

y = 2016
for index, row in BS.iterrows():
    CAn = row['com %']
    BS.at[y, 'commissions année'] = CAn
    
    PR = row['pourcentage règlement']
    CP = (PR / 100) * CAn
    BS.at[y, 'commission période'] = CP
    
    RAP = CAn - CP
    BS.at[y, 'commission RAP'] = RAP

    y = y + 1

y = 2016
for index, row in SA.iterrows():
    CAn = row['com %']
    SA.at[y, 'commissions année'] = CAn
    
    PR = row['pourcentage règlement']
    CP = (PR / 100) * CAn
    SA.at[y, 'commission période'] = CP
    
    RAP = CAn - CP
    SA.at[y, 'commission RAP'] = RAP

    y = y + 1

BS = BS.append(BS.sum().rename('Total'))

SA = SA.append(SA.sum().rename('Total'))

del BS['com %']
del SA['com %']

#print(f'commissions Sahbi Bejaoui à payer, période: {date_deb} - {date_fin}\n')
#print(BS)
#print('\n')

#print(f'commissions abdeelkrim saidi à payer, période: {date_deb} - {date_fin}\n') 
#print(SA)

Obj_BS = pd.concat([Obj_BS, BS], axis=1)
#Obj_BS.round(1)

Obj_SA = pd.concat([Obj_SA, SA], axis=1)
#Obj_SA.round(1)

# pd.options.display.float_format = '{:, .2f}'.format

Obj_SA.to_excel("Commissions_SaidiA_periode_saisie.xlsx", float_format="%.1f")
Obj_BS.to_excel("Commissions_BejaouiS_periode_saisie.xlsx", float_format="%.1f")
