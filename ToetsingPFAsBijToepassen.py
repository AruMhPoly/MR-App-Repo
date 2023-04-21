# In[]:

import os
import pandas as pd 

# PFAS = r"P:\2022\22196 Egaliseren Theodorushaven Bergen op Zoom\V1\07 Laboratorium\3 Toetsingen\Theodorushaven\EXCEL\22218V1_Output_PFAS.xlsx"
# PS= r'P:\2022\22196 Egaliseren Theodorushaven Bergen op Zoom\V1\07 Laboratorium\3 Toetsingen\Theodorushaven\EXCEL'

class PFASToepassing():
    def __init__(self,PFASPath,PathSave,ProjectNummer):
        self.PFASPath = PFASPath
        self.PathSave = PathSave
        self.ProjectNummer = ProjectNummer
        self.PFAS_Kader = r"C:\Python\MR_APP\MR-App-Repo\Kader\PFAS.xlsx"
        
    def Toepassing(self):

        df_Kader =  pd.read_excel(self.PFAS_Kader)
        df_Input = pd.read_excel(self.PFASPath)
        df_Input.set_index(df_Input.columns.to_list()[0])
        Mengmonsters = df_Input['Mengmonster'].tolist()
        Columns = [df_Input.columns.tolist()[1],df_Input.columns.tolist()[2],
                df_Input.columns.tolist()[3],df_Input.columns.tolist()[4],
                df_Input.columns.tolist()[6]]
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
        Path_Save = os.path.join(self.PathSave, self.ProjectNummer +'_PFAS_Toepassing.xlsx')
        df.to_excel(Path_Save)

#In[]

# PFASToepassing(PFASPath=PFAS,PathSave=PS,ProjectNummer="22218V1").Toepassing()


#In[]: