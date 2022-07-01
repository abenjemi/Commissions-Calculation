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