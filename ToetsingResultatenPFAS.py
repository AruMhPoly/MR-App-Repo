# In[]:

import os
import pandas as pd 
import openpyxl

# Path = r'C:\Python\MR_APP\TESTEN_API\EXCEL'
# PS = r'C:\Python\MR_APP\TESTEN_API\EXCEL'

#In[]

class PFAS: 

    def __init__(self,PathSave, Path_Certificaten,ProjectNummer):
        self.Path_Certificaten = Path_Certificaten
        self.PathSave = PathSave
        self.ProjectNummer = ProjectNummer

    def ResultatenPFAS(self):

        Monsters_PFAS = []
        Monsters = []
        Org_Stof =[]
        CommonPFAS = ["SOM PFOS","SOM PFOA","EtFOSAA","MeFOSAA"]
        IgnorePar = ["C08: PFOSv","C08: PFOS","C08: PFOAv","C08: PFOA",]
        PFOS = []
        PFOA = []
        EtFOSAA = []
        MeFOSAA = []
        HAP = []
        Name_Par = []

        for filename in os.listdir(self.Path_Certificaten):
            
            f = os.path.join(self.Path_Certificaten, filename)

            ROWS_PFAS = []
            ROWS_OS = []

            # Load the Excel file
            workbook = openpyxl.load_workbook(f)
            # Select the active worksheet
            worksheet = workbook.active

            # Iterate through the rows starting from row 7
            for row in worksheet.iter_rows(min_row=7):
                # Check if "µg/kg ds" is present in any cell of the row
                for cell in row:
                    if cell.value == "µg/kg ds":
                        if row[1].value not in Monsters_PFAS:
                            Monsters_PFAS.append(row[1].value)
                            ROWS_PFAS.append(cell.row)
                    elif cell.value == "mg/kg ds":
                        if row[1].value not in Monsters:
                            Monsters.append(row[1].value) 
                            ROWS_OS.append(cell.row)
            
            #Organisch stof
            OS_Col = None
            for column in worksheet.iter_cols():
                            for cell in column:
                                if cell.value == "Org.Stof.cor":
                                    OS_Col = cell.column
                                    break
                            if OS_Col is not None:
                                break

            #Extract the values of the Organisch Stof 

            for y in ROWS_OS:
                Org_Stof.append(worksheet.cell(row=y, column=OS_Col).value) 


            #Get the most common PFAS parameters

            for column in worksheet.iter_cols():
                        for cell in column:
                            if cell.value == CommonPFAS[0]:
                                for y in ROWS_PFAS:
                                    PFOS.append(worksheet.cell(row=y, column=cell.column).value)

                            elif cell.value == CommonPFAS[1]:
                                for y in ROWS_PFAS:
                                    PFOA.append(worksheet.cell(row=y, column=cell.column).value)

                            elif cell.value == CommonPFAS[2]:
                                for y in ROWS_PFAS:
                                    EtFOSAA.append(worksheet.cell(row=y, column=cell.column).value)

                            elif cell.value == CommonPFAS[3]:
                                for y in ROWS_PFAS:
                                    MeFOSAA.append(worksheet.cell(row=y, column=cell.column).value)

            #Get the highest other PFAS


            for y in ROWS_PFAS:
                
                    #List to add all PFAS values
                    UL = [] 
                    # List to add the name of the parameters
                    UL2 = []
                    # Columns with PFAS values. 
                    UL3 = []
                    for column in worksheet.iter_cols():
                        cell = column[y - 1]  # Adjusted for 0-based indexing
                        if cell.value == 'µg/kg ds':
                            UL.append(worksheet.cell(row=y, column=cell.column-1).value)
                            UL2.append(worksheet.cell(row=1, column=cell.column-1).value)

                    #Convert the list in float to extract the values

                    for item in UL:
                        try:
                            UL3.append(float(item))
                        except:
                            UL3.append((item))

                    #Get Maximum value 

                    max_value = float('-inf')
                    max_indices = []

                    for i, value in enumerate(UL3):
                        if isinstance(value, str):  # Ignore strings
                            continue
                        if value > max_value and UL2[i] not in CommonPFAS and UL2[i] not in IgnorePar:
                            max_value = value
                            max_indices = [i]
                        elif value == max_value:
                            max_indices.append(i)

                    # Concatenate the parameters that belong to the highest PFAS

                    n  = ""
                    #This is for a conditional so I can now where a new big value was found
                    z = 0
                    for x in max_indices:
                        if UL2[x] not in CommonPFAS:
                            n = n + "-" + str(UL2[x])
                            z = UL3[max_indices[0]]
                        else:
                            n = n + ""
                        

                    Name_Par.append(n)
                    HAP.append(z)

        df = pd.DataFrame(columns=['Mengmonster','Som PFOS (µg/kg ds)',
                                'SOM PFOA (µg/kg ds)','EtFOSAA (µg/kg ds)',
                                'MeFOSAA (µg/kg ds)','Hoogste andere PFAS'
                                ,"Concentratie (µg/kg ds)","Organische stof (%)","Gecorrigeerd?"])

        df['Mengmonster'] = Monsters_PFAS
        df['Som PFOS (µg/kg ds)'] = PFOS
        df['SOM PFOA (µg/kg ds)'] = PFOA
        df['EtFOSAA (µg/kg ds)'] = EtFOSAA
        df['MeFOSAA (µg/kg ds)'] = MeFOSAA
        df['Hoogste andere PFAS'] = Name_Par
        df ["Concentratie (µg/kg ds)"] = HAP

        df["Organische stof (%)"] = Org_Stof


        df[['Som PFOS (µg/kg ds)', 'SOM PFOA (µg/kg ds)',
            'EtFOSAA (µg/kg ds)',
            'MeFOSAA (µg/kg ds)',
            "Concentratie (µg/kg ds)",
            'Organische stof (%)']] = df[['Som PFOS (µg/kg ds)', 'SOM PFOA (µg/kg ds)',
            'EtFOSAA (µg/kg ds)',
            'MeFOSAA (µg/kg ds)',
            "Concentratie (µg/kg ds)",
            'Organische stof (%)']].apply(pd.to_numeric, errors='coerce').fillna(0)

        #The sum can not be 0.1, then it means it is not repported above the threshold value

        df.replace({'Som PFOS (µg/kg ds)': 0.1, 'SOM PFOA (µg/kg ds)': 0.1}, 0, inplace=True)

        # Find rows where column D is greater than 10
        condition = df['Organische stof (%)'] > 10
        # Update values in columns A and B



        try:
            df.loc[condition, 'Som PFOS (µg/kg ds)',] = (df['Som PFOS (µg/kg ds)'] / df['Organische stof (%)']) * 10
            df.loc[condition, 'SOM PFOA (µg/kg ds)',] = (df['SOM PFOA (µg/kg ds)'] / df['Organische stof (%)']) * 10
            df.loc[condition, 'EtFOSAA (µg/kg ds)',] = (df['EtFOSAA (µg/kg ds)'] / df['Organische stof (%)']) * 10
            df.loc[condition, 'MeFOSAA (µg/kg ds)',] = (df['MeFOSAA (µg/kg ds)'] / df['Organische stof (%)']) * 10
            df.loc[condition, 'Concentratie (µg/kg ds)',] = (df['Concentratie (µg/kg ds)'] / df['Organische stof (%)']) * 10
        except:
            pass
        
        df.replace({'Som PFOS (µg/kg ds)': 0.1, 'SOM PFOA (µg/kg ds)': 0.1}, 0, inplace=True)
        df["Gecorrigeerd?"] = df['Organische stof (%)'].apply(lambda x: 'Ja' if x > 10 else 'Nee')
        df.set_index('Mengmonster',inplace=True)
        df.replace(0,"--",inplace=True)
        Path_Save = os.path.join(self.PathSave, self.ProjectNummer + '_Output_PFAS.xlsx')
        df.to_excel(Path_Save)
        return Path_Save
#In[]:

# Test = PFAS(PathSave=PS,Path_Certificaten=Path,ProjectNummer="P_")
# Test.ResultatenPFAS()
#In[]