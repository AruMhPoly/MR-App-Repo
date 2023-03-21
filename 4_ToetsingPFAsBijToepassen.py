# In[]:

import os
import pandas as pd 

# In[]: 

# Input Paths

PFAS_Kader = r"C:\Python\MR_APP\MR-App-Repo\Kader\PFAS.xlsx"
# This has to be change by a dataframe 
df_Input = pd.read_excel(r"C:\Python\MR_APP\MR-App-Repo\Output\3.xlsx")

# In[]: 

df_Kader =  pd.read_excel(PFAS_Kader)
df_Input.set_index(df_Input.columns.to_list()[0])
# In[]: 

Mengmonsters = df_Input['Mengmonster'].tolist()

# In[]:

Columns = [df_Input.columns.tolist()[2],df_Input.columns.tolist()[3],
           df_Input.columns.tolist()[4],df_Input.columns.tolist()[5],
           df_Input.columns.tolist()[7]]

# In[] 


Names_Columns = df_Kader['Categorie'].to_list()
df = pd.DataFrame(columns=Names_Columns)


#In[]: 

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

#In[]:
df = df.set_index(pd.Index(Mengmonsters,name="Mengmonster"))


# In[]
df.to_excel(r"C:\Python\MR_APP\MR-App-Repo\Output\4.xlsx")


#In[]: