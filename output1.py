import numpy as np
import pandas as pd

# reading the csv file
data = pd.read_excel('input1.xlsx', header = 0, encoding = "ISO-8859-1", error_bad_lines=False, warn_bad_lines=False)

doc = data['DOC'][:].tolist()
data.drop_duplicates(subset = ["DOC"], keep="first", inplace=True)
pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)

print('Result 1:\n')
print(data)
print('\n\n')

#2
ran = 2025 - 2016 + 1
d = {
    'Bejaoui Sahbi':[0.0] * ran,
    'Sahbi Bejaoui REV':[0.0] * ran,
    'Saidi Abdelkarim':[0.0] * ran,
    'Abdelkarim Saidi REV':[0.0] * ran  
}
ind = range(2016, 2016 + ran)
df = pd.DataFrame(d, columns = ['Bejaoui Sahbi' , 'Sahbi Bejaoui REV', 'Saidi Abdelkarim' , 'Abdelkarim Saidi REV'], index=list(ind))

rpr = list(df.columns)
for y in ind:
	for rep in rpr:
		summ = 0
		for index, row in data.iterrows():
			if (row['AN'] == y and row['Representant'] == rep): 
				summ = summ + row['Total HT']
		df.at[y, rep] = summ

print('Result 2:\n')
print (df)
print('\n\n')

#3
input2_BS = pd.read_excel('input2_BS.xlsx', header = 0, encoding = "ISO-8859-1", error_bad_lines=False, warn_bad_lines=False)
input2_BS = input2_BS.fillna(0)
input2_BS['CA'] = [0.0] * ran

input2_SA = pd.read_excel('input2_SA.xlsx', header = 0, encoding = "ISO-8859-1", error_bad_lines=False, warn_bad_lines=False)
input2_SA = input2_SA.fillna(0)
input2_SA['CA'] = [0.0] * ran

x = 0
for index, row in df.iterrows():
	input2_BS.at[x, 'CA'] = row[rpr[0]] + row[rpr[1]]
	input2_SA.at[x, 'CA'] = row[rpr[2]] + row[rpr[3]]
	x = x + 1	
		

table3 = df.copy()
table3.insert(2, "Bejaoui Sahbi EX", [0.0] * ran, True)
table3.insert(5, "Saidi Abdelkarim EX", [0.0] * ran, True)

x = 2016
for index, row in input2_BS.iterrows():
	if (row['CA'] > row['objectif min']):
		table3.at[x, 'Bejaoui Sahbi'] = (row['CA'] - row['objectif min']) * row['% VD']
		table3.at[x, 'Sahbi Bejaoui REV'] = (row['CA'] - row['objectif min']) * row['% VR']
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
	if (row['CA'] > row['objectif min']):
		table3.at[z, 'Saidi Abdelkarim'] = (row['CA'] - row['objectif min']) * row['% VD']
		table3.at[z, 'Abdelkarim Saidi REV'] = (row['CA'] - row['objectif min']) * row['% VR']
	else:
		table3.at[z, 'Saidi Abdelkarim'] = 0
		table3.at[z, 'Abdelkarim Saidi REV'] = 0
	if (row['CA'] > row['objectif excellence']):
		table3.at[z, 'Saidi Abdelkarim EX'] = row['CA'] * row['% excellence']	
	else:
		table3.at[z, 'Saidi Abdelkarim EX'] = 0
	z = z + 1

print ('Result 3:\n')
print(table3)
print('\n\n')

#4
f = {
    'Bejaoui Sahbi':[0.0] * ran,
    'Saidi Abdelkarim':[0.0] * ran  
}
table4 = pd.DataFrame(f, columns = ['Bejaoui Sahbi' , 'Saidi Abdelkarim'], index=list(ind))

for index, row in table3.iterrows():
	table4.at[index, 'Bejaoui Sahbi'] = row['Bejaoui Sahbi'] + row['Sahbi Bejaoui REV'] + row['Bejaoui Sahbi EX']
	table4.at[index, 'Saidi Abdelkarim'] = row['Saidi Abdelkarim'] + row['Abdelkarim Saidi REV'] + row['Saidi Abdelkarim EX']

print('Result 4: \n')
print(table4)
