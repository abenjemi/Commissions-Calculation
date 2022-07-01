from tkinter import *
from tkinter import messagebox
import datetime

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
        messagebox.showinfo("", "Veuillez cliquer pour voir les resultats suivants sur le meme dossier:\n1. etat_reglements_factures\n2. etat_factures_reglees100%_période_saisie\n3. Commissions_BejaouiS_periode_saisie\n4. Commissions_SaidiA_periode_saisie")
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

print(date_deb)