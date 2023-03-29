#In[]

import os
import pandas as pd 
import openpyxl

#In[]

#Input 

#Toetsingen MijnLab
df1 = pd.read_excel(r"P:\2022\22218 WNZ diverse vakken LN 2023\V1\07 Laboratorium\3 Toetsingen\RA04\EXCEL\Output_Botova.xlsx")
#PFAS toepassing
df2 = pd.read_excel(r"P:\2022\22218 WNZ diverse vakken LN 2023\V1\07 Laboratorium\3 Toetsingen\RA04\EXCEL\PFAS_Toepassing.xlsx")
#Path Save
Path_Toetsingen= r"P:\2022\22218 WNZ diverse vakken LN 2023\V1\07 Laboratorium\3 Toetsingen\RA04\EXCEL"
#In[]
df2.replace("✔",0,inplace=True)
df2.replace("--",1,inplace=True)
#Toetsingen
T = ["Monster","T1","T3","T6","T7","T9","T11"]
Toep = ["Monster","Baggerspecie kwaliteit - Landbodem",
        "Baggerspecie kwaliteit - Waterbodem",
        "Verspreiden oppervlaktewaterlichaam in een zoet oppervlaktewaterlichaam",
        "Verspreiden oppervlaktewaterlichaam in een zout oppervlaktewaterlichaam",
        "(Grootschalige) toepassing landbodem","(Grootschalige) toepassing oppervlaktewaterlichaam"]

Dict = dict(zip(T, Toep))
try:
    df1 = df1.drop('Parameters Overschreden bij T3', axis=1)
except:
    pass

df1 = df1.drop("Unnamed: 0", axis=1)
keys= df1.columns.to_list()[:-1]
Columns = [Dict[key] for key in keys]
df = pd.DataFrame(columns=Columns)
df['Monster'] = df1['Monster']
try: 
    df['Baggerspecie kwaliteit - Landbodem'] = df1['T1']
except:
    pass

df['Baggerspecie kwaliteit - Waterbodem'] = df1['T3']
# Quite unconvenient, try to optimize it. 
my_list = []
for value in df1['T6']:
    if value == 'Verspreidbaar':
        my_list.append(0)
    else:
        my_list.append(1)

df[Toep[3]] = my_list

#T7 won't be always tested, it's better to add error handeling

try: 

    my_list = []
    for value in df1['T7']:
        if value == 'Verspreidbaar':
            my_list.append(0)
        else:
            my_list.append(1)

    df[Toep[4]] = my_list

except:

    pass

my_list = []
for value in df1['T9']:
    if value == 'Toepasbaar in GBT':
        my_list.append(0)
    elif value == 'Overschrijding Emissietoetswaarde':
        my_list.append("Grootschalig: Eerst uitloog onderzoek")
    else:
        my_list.append(1)

df[Toep[5]] = my_list

my_list = []
for value in df1['T11']:
    if value == 'Toepasbaar in GBT':
        my_list.append(0)
    elif value == 'Overschrijding Emissietoetswaarde':
        my_list.append("Grootschalig: Eerst uitloog onderzoek")
    else:
        my_list.append(1)

df[Toep[6]] = my_list
my_list = []

#Where PFAS is an issue 
my_list = []
for index, row in df2.iterrows():
    matches = [col_name for col_name, value in row.items() if value == 1]
    my_list.append(', '.join(matches))

df['Uitzondering toepassing bij PFAS'] = my_list

for index, row in df.iterrows():
    count = len(row['Uitzondering toepassing bij PFAS'].split(','))
    if count > 3:
        try: 
            df.iloc[index, Toep[3]] = 1
            print(index)
        except: 
            pass
        try:
            df.iloc[index, Toep[4]] = 1
        except:
            pass

my_list = []
for index, row in df.iterrows():
    row_sum = 0
    for col in row:
        if isinstance(col, (int, float)):
            row_sum += col
    if not any(elem in row['Baggerspecie kwaliteit - Waterbodem'] for elem in ["Altijd toepasbaar", "Klasse A"]) and row_sum > 0:
        my_list.append('Ja')
    else:
        my_list.append('Nee')

df['Afvoeren naar (Rijks)baggerdepot'] = my_list
#Path Save
df.to_excel(os.path.join(Path_Toetsingen,'Mogelijkheden.xlsx'))

#In[]:

