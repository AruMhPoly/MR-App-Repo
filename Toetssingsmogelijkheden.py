import os
import pandas as pd 
import openpyxl

#In[]

#Input 

#Toetsingen MijnLab
df1 = pd.read_excel(r"P:\2023\23002 Jachthaven Klimbos Dordrecht\V1\07 Laboratorium\1.xlsx")
#PFAS toepassing
df2 = pd.read_excel(r"P:\2023\23002 Jachthaven Klimbos Dordrecht\V1\07 Laboratorium\3.xlsx")
#In[]
df2.replace("✔",0,inplace=True)
df2.replace("--",1,inplace=True)
#In[]

#Toetsingen
T = ["Monster","T3","T6","T7","T9","T11"]
Toep = ["Monster","Baggerspecie kwaliteit",
        "Verspreiden oppervlaktewaterlichaam in een zoet oppervlaktewaterlichaam",
        "Verspreiden oppervlaktewaterlichaam in een zout oppervlaktewaterlichaam",
        "(Grootschalige) toepassing landbodem","(Grootschalige) toepassing oppervlaktewaterlichaam"]

Dict = dict(zip(T, Toep))

# In[]:
keys= df1.columns.to_list()[:-1]
Columns = [Dict[key] for key in keys]
#In[]

df = pd.DataFrame(columns=Columns)

#In[]:

df['Monster'] = df1['Monster']
df['Baggerspecie kwaliteit'] = df1['T3']

#In[]

# Quite unconvenient, try to optimize it. 
my_list = []
for value in df1['T6']:
    if value == 'Verspreidbaar':
        my_list.append(0)
    else:
        my_list.append(1)

df[Toep[2]] = my_list

my_list = []
for value in df1['T7']:
    if value == 'Verspreidbaar':
        my_list.append(0)
    else:
        my_list.append(1)

df[Toep[3]] = my_list

my_list = []
for value in df1['T9']:
    if value == 'Toepasbaar in GBT':
        my_list.append(0)
    elif value == 'Overschrijding Emissietoetswaarde':
        my_list.append("Grootschalig: Eerst uitloog onderzoek")
    else:
        my_list.append(1)

df[Toep[4]] = my_list

my_list = []
for value in df1['T11']:
    if value == 'Toepasbaar in GBT':
        my_list.append(0)
    elif value == 'Overschrijding Emissietoetswaarde':
        my_list.append("Grootschalig: Eerst uitloog onderzoek")
    else:
        my_list.append(1)

df[Toep[5]] = my_list
my_list = []

#In[]
#Where PFAS is an issue 
my_list = []
for index, row in df2.iterrows():
    matches = [col_name for col_name, value in row.items() if value == 1]
    my_list.append(', '.join(matches))

#In[]:

df['Uitzondering toepassing bij PFAS'] = my_list
df.drop(columns=[Toep[3]], inplace=True)

#In[]:

for index, row in df.iterrows():
    count = len(row['Uitzondering toepassing bij PFAS'].split(','))
    if count > 3:
        try: 
            df.iloc[index, Toep[2]] = 1
            print(index)
        except: 
            pass
        try:
            df.iloc[index, Toep[3]] = 1
        except:
            pass


#In[]:

my_list = []
for index, row in df.iterrows():
    row_sum = 0
    for col in row:
        if isinstance(col, (int, float)):
            row_sum += col
    if not any(elem in row['Baggerspecie kwaliteit'] for elem in ["Klasse AT", "Klasse A"]) and row_sum > 0:
        my_list.append('Ja')
    else:
        my_list.append('Nee')

df['Afvoeren naar (Rijks)baggerdepot'] = my_list
#In[]:
