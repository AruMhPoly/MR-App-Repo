#In[]: 
import os
import pandas as pd 


class PFASToepassing():
    def __init__(self,PFASPath,PathSave,ProjectNummer):
        self.PFASPath = PFASPath
        self.PathSave = PathSave
        self.ProjectNummer = ProjectNummer


    def Kade(self):

        Columns = ['Categorie', 'Som PFOS (µg/kg ds)', 'SOM PFOA (µg/kg ds)',
       'EtFOSAA (µg/kg ds)', 'MeFOSAA (µg/kg ds)', 'Concentratie (µg/kg ds)']
        A = ['4.1.1', '4.1.2', '4.1.3', '4.2', '4.3', '4.4', '4.7 (Rijkswater)',
            '4.7 (Regionale Water)', '4.8.1 (Rijkswater)',
            '4.8.1 (Regionale Water)', '4.8.2 (Rijkswater)',
            '4.8.2 (Anders)', '4.9.1', '4.9.2']
        B = [3. , 1.4, 1.4, 3. , 3. , 0.1, 8.2, 2.2, 8.2, 2.2, 3.7, 1.1, 3.7,
            1.1]
        C =[7. , 1.9, 1.9, 7. , 7. , 0.1, 0.8, 0.9, 0.8, 0.9, 0.8, 0.8, 0.8,
            0.8]
        D = [3. , 1.4, 1.4, 3. , 3. , 0.1, 5.5, 1.8, 5.5, 1.8, 0.8, 0.8, 0.8,
            0.8]
        E = [3. , 1.4, 1.4, 3. , 3. , 0.1, 1. , 0.8, 1. , 0.8, 0.8, 0.8, 0.8,
            0.8]
        F = [3. , 1.4, 1.4, 3. , 3. , 0.1, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8,
            0.8]

        # Create a dictionary from the lists
        data = {name: values for name, values in zip(Columns, [A, B, C, D, E, F])}

        # Create a pandas DataFrame
        df = pd.DataFrame(data)

        return df 
      
        
    def Toepassing(self):

        df_Kader =  self.Kade()
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
                    Res.append("1")
                else:
                    Res.append("0")

            new_row = pd.DataFrame([Res], columns=Names_Columns)
            df = pd.concat([df, new_row], ignore_index=True)
            Res = []
        df = df.set_index(pd.Index(Mengmonsters,name="Mengmonster"))
        Path_Save = os.path.join(self.PathSave, self.ProjectNummer +'_PFAS_Toepassing.xlsx')
        df.to_excel(Path_Save)
        return Path_Save

#In[]
# PFAS = r"C:\Python\MR_APP\SGS\23209V1\TOETSINGEN\23121V1_Output_PFAS.xlsx"
# PS= r"C:\Python\MR_APP\SGS\23209V1\TOETSINGEN"
# X =PFASToepassing(PFASPath=PFAS,PathSave=PS,ProjectNummer="22218V1").Toepassing()

#In[]: