#In[]: 
import os
import pandas as pd 
import openpyxl

class Toetssingsmogelijkheden:

        def __init__(self,MonstersBoToVa,MonsPFAS,PathSave,ProjectNummer):

                self.MonstersBoToVa = MonstersBoToVa
                self.MonsPFAS = MonsPFAS
                self.PathSave = PathSave
                self.ProjectNummer = ProjectNummer

        # I don't know if this is correct, so I just do it to be sure

        def DicValues(self):

                T = ["T3","T5","T6","T7","T9","T11"]
                Toep = ["Baggerspecie kwaliteit - Waterbodem",
                        'Verspreiden op een aangrenzend perceel (landbodem)',
                        "Verspreiden in een zoet oppervlaktewaterlichaam",
                        "Verspreiden in een zout oppervlaktewaterlichaam",
                        "(Grootschalige) toepassing landbodem",
                        "(Grootschalige) toepassing oppervlaktewaterlichaam"]
                # Create a dictionary to use it for the titles of the output 
                Dic = {k: v for k, v in zip(T, Toep)}
                return Dic
        
        def Mogelijkheden(self):
                #Results BoTova
                df1 = pd.read_excel(self.MonstersBoToVa,header=1)
                #Posibilities to use PFAS 
                df2 = pd.read_excel(self.MonsPFAS)
                #First, clean the data
                # Drop the first row
                df1 = df1.drop(0)
                # Identify columns with a dot (".") in the header
                columns_to_drop = [col for col in df1.columns if '.' in col]
                # Drop the columns with a point in their titles. These are the exceeded parameters. I don't need them. 
                df1 = df1.drop(columns=columns_to_drop)
                df1.reset_index(inplace=True)
                df1 = df1.iloc[:,1:]
                # First let's drop categorie 4.4 because that one is not considered for re-use. Source: Aline 
                df2 = df2.drop(columns=['4.4','Mengmonster']).reset_index()
                df2 =df2.drop(columns='index')
                #Divide it water and land
                Land = df2.iloc[:,:5]
                Water = df2.drop(columns=['4.1.1', '4.1.2', '4.1.3', '4.2', '4.3'])
                #Empty lists to create the output dataframe 
                Monsters = []
                Result = []
                Titles = ['Monster',]

                # Which one are the monsters? 
                Mons = df1['Toetsing'].tolist()
                # Iterate through the rows and print the index of each row
                for index, row in df1.iterrows():

                        df_filt = df1[df1["Toetsing"]==Mons[index]]

                        # Let's start to iterate through the columns 
                        for ColumnName, Value in df_filt.items():
                                # In which categories is PFAS a restriction?
                                #First Landbodem
                                T = Land.iloc[index,:]
                                PR1 = T[T>0].index.tolist()
                                PFAS = "- ".join(PR1)  
                                if len(PR1) == 0: 
                                        TL = "PFAS: Geen Beperking"
                                else: 
                                        TL = "PFAS beperkt door categoreën: " +  PFAS

                                #Now waterbodem
                                T = Water.iloc[index,:]
                                PR2 = T[T>0].index.tolist()
                                PFAS = "-".join(PR2)  
                                if len(PR2) == 0: 
                                        TW = "PFAS: Geen Beperking"
                                else: 
                                        TW = "PFAS beperkt door categoreën: " +  PFAS

                                # I have to do it test by test

                                #Call the dictionary

                                Dic = self.DicValues()
                        
                                if ColumnName in ["T5"]:

                                        #Get the title for last Pandas Dataframe 

                                        if Dic[ColumnName] not in Titles:

                                                Titles.append(Dic[ColumnName])

                                        if Value.values[0] not in ['Verspreidbaar']:
                                        
                                                TT = "Geen Toepassingsmogelijkhed" + "\n" + TL #Toetsing Text 
                                                Result.append(TT)
                                        else: 
                                                TT = Value.values[0] + "\n" + TL
                                                Result.append(TT)
                                
                                elif ColumnName in ["T6","T7"]:
                                        #Get the title for last Pandas Dataframe 

                                        if Dic[ColumnName] not in Titles:
                                                Titles.append(Dic[ColumnName])

                                        if Value.values[0] not in ['Verspreidbaar']:
                                        
                                                TT = "Geen Toepassingsmogelijkhed" + "\n" + TW #Toetsing Text 
                                                Result.append(TT)
                                        else: 
                                                TT = Value.values[0] + "\n" + TW
                                                Result.append(TT)
                                
                                elif ColumnName == "T9":

                                        #Get the title for  Pandas Dataframe 
                                        if Dic[ColumnName] not in Titles:

                                                Titles.append(Dic[ColumnName])

                                        if Value.values[0] == "Toepasbaar in GBT":

                                                TT = Value.values[0] + "\n" + TW #Toetsing Text 
                                                Result.append(TT)

                                        elif Value.values[0] == "Overschrijding Emissietoetswaarde":

                                                TT = "Eerst uitloog onderzoek" + "\n" + TL
                                                Result.append(TT)
                                        else:

                                                TT = "Geen Toepassingsmogelijkhed" + "\n" + TL
                                                Result.append(TT)

                                elif ColumnName == "T11":


                                        #Get the title for  Pandas Dataframe 

                                        if Dic[ColumnName] not in Titles:

                                                Titles.append(Dic[ColumnName])
                                        if Value.values[0] == "Toepasbaar in GBT":
                                                TT = Value.values[0] + "\n" + TW #Toetsing Text 
                                                Result.append(TT)
                                        elif Value.values[0] == "Overschrijding Emissietoetswaarde":

                                                TT = "Eerst uitloog onderzoek" + "\n" + TW
                                                Result.append(TT)
                                        else:
                                                TT = "Geen Toepassingsmogelijkhed" + "\n" + TW
                                                Result.append(TT)
                                                
                                elif ColumnName not in list(Dic.keys())[1:] and ColumnName not in ['Toetsing',"T1"]: # As a result it will only entry for T3
                                        # First the name of the column

                                        if "Afvoeren naar (Rijks)baggerdepot" not in Titles:

                                                Titles.append("Afvoeren naar (Rijks)baggerdepot")
                                        #Add the monster

                                        Monsters.append(df_filt['Toetsing'].values[0])
                                        
                                        if Value.values[0] in ["Klasse AT","Klasse A"]: 
                                                if len(PR1)!=len(Land.columns) and len(PR2)!=len(Water.columns):
                                                        TT = "Geen toepassingsmogelijkheid"
                                                        Result.append(TT)
                                                elif len(PR1)==len(Land.columns) and len(PR2)==len(Water.columns):
                                                        TT = "Wel toepassingsmogelijkheid"
                                                        Result.append(TT)

                                        elif Value.values[0] in ["Klasse B"]: 

                                                Uitlog = 0 #It means it is not needed an uitloog onderozek
                                                # I have to see if uitloog onderzoek is necessary
                                                if 'T9' in df_filt.columns.tolist():
                                                        T = df_filt['T9'].values[0]
                                                        if T == "Overschrijding Emissietoetswaarde":
                                                                Uitlog = 1

                                                if 'T11' in df_filt.columns.tolist():
                                                        T = df_filt['T11'].values[0]
                                                        if T == "Overschrijding Emissietoetswaarde":
                                                                Uitlog = 1

                                                if len(PR1)!=len(Land.columns) and len(PR2)!=len(Water.columns):
                                                        if Uitlog == 0:
                                                                TT = "Geen toepassingsmogelijkheid"
                                                                Result.append(TT)
                                                        elif Uitlog == 1:
                                                                TT = "Niet doorslaaggevend"
                                                                Result.append(TT)

                                                elif len(PR1)==len(Land.columns) and len(PR2)==len(Water.columns):
                                                        TT = "Wel toepassingsmogelijkheid"
                                                        Result.append(TT)

                                        else: 
                                                TT = "Indien geen toepassing voorhanden is, kan in overleg met de depotbeheerder besloten worden om het materiaal af te voeren naar een Rijksbaggerdepot"
                                                Result.append(TT)


                        # Reshape the list into a 2D array 
                        R = len(Titles)-1
                        Data = [Result[i:i+R] for i in range(0, len(Result), R)]
                        # Create the DataFrame
                        df = pd.DataFrame(Data)
                        df.columns = Titles[1:]
                        df[Titles[0]] = Monsters
                        df.set_index("Monster", inplace=True)
                        # Get the column names of the DataFrame
                        columns = df.columns.tolist()
                        # Move the desired column to the last position
                        columns.remove("Afvoeren naar (Rijks)baggerdepot")
                        columns.append("Afvoeren naar (Rijks)baggerdepot")
                        # Reorder the DataFrame columns
                        df = df[columns]
                        f = os.path.join(self.PathSave,self.ProjectNummer + "_Toepassingsmogelijkheden.xlsx")
                        df.to_excel(f)

# M = Toetssingsmogelijkheden(MonstersBoToVa=r"P:\2023\23116 Kade Zomerlust\V1\07 Laboratorium\3 Toetsingen\EXCEL\22216V1_Output_BoToVa.xlsx",
#                         MonsPFAS=r"P:\2023\23116 Kade Zomerlust\V1\07 Laboratorium\3 Toetsingen\EXCEL\22216V1_PFAS_Toepassing.xlsx",
#                         PathSave=r'P:\2023\23116 Kade Zomerlust\V1\07 Laboratorium\3 Toetsingen\EXCEL',
#                         ProjectNummer="22216V1")
# M.Mogelijkheden()

#In[]: