# In[]:

import os
import pandas as pd 

# In[]: 

# Input Paths

PFAS_Kader = r"C:\Python\MR_APP\MR-App-Repo\Kader\PFAS.xlsx"
# This has to be change by a dataframe 
df_Input = pd.read_excel(r"P:\2022\22218 WNZ diverse vakken LN 2023\V1\07 Laboratorium\3 Toetsingen\RA01\EXCEL\Output_PFAS.xlsx")
#Path Save
Path_Toetsingen= r"P:\2022\22218 WNZ diverse vakken LN 2023\V1\07 Laboratorium\3 Toetsingen\RA01\EXCEL"

df_Kader =  pd.read_excel(PFAS_Kader)
df_Input.set_index(df_Input.columns.to_list()[0])
Mengmonsters = df_Input['Mengmonster'].tolist()
Columns = [df_Input.columns.tolist()[2],df_Input.columns.tolist()[3],
           df_Input.columns.tolist()[4],df_Input.columns.tolist()[5],
           df_Input.columns.tolist()[7]]
Names_Columns = df_Kader['Categorie'].to_list()
df = pd.DataFrame(columns=Names_Columns)
for y in range(df_Input.shape[0]):

    Cols = []
    for x in Columns:
        if isinstance(df_Input.loc[y,x], (int, float, complex)):
            Cols.append(x)
    Res = []
    # Iterate over each row of df_Kader and compare values with the selected columns from df_Input
    for index, row in df_Kader.iterrows():
        if (df_Input.loc[y, Cols] > row[Cols]).any():
            Res.append("--")
        else:
            Res.append("✔")

    new_row = pd.DataFrame([Res], columns=Names_Columns)
    df = pd.concat([df, new_row], ignore_index=True)
    Res = []
df = df.set_index(pd.Index(Mengmonsters,name="Mengmonster"))
df.to_excel(os.path.join(Path_Toetsingen,'PFAS_Toepassing.xlsx'))

#In[]:
