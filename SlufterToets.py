#In[]: 

import os
import pandas as pd 
import openpyxl
import datetime



class SlufterToets:
    
    def __init__(self,Path_Toetsingen,Projectnummer,Path_PFAS):

        self.Path_Toetsingen = Path_Toetsingen
        self.Path_PFAS = Path_PFAS
        self.ProjectNummer = Projectnummer
        self.Template = r"C:\Python\MR_APP\MR-App-Repo\SlufterToets\SluftertoetsTemplate.xlsx"

    def RunTest(self):

        Beoordelingen = {"Altijd toepasbaar":"AW",
                "Klasse A":"A",
                "Klasse B":"B",
                "Nooit toepasbaar":"NT"}


        # Empty lists

        As = [] 
        Cd = []
        Cr = []
        Cu = []
        Hg = []
        Pb = []
        Ni = []
        Zn = []
        naftaleen = []
        anthraceen = []
        fenanthreen = []
        fluorantheen = []
        benzaanthr = []
        chryseen = []
        benzoapyre = []
        benzoghipe = []
        benzokfluo = []
        indeno123p = []
        PCB28 = []
        PCB52 = []
        PCB101 = []
        PCB118 = []
        PCB138 = []
        PCB153 = []
        PCB180 = []
        OLIEFLG = []
        HEXACHLB =[]
        D24DDD =  []
        D44DDD =  []
        D24DDE =  []
        D44DDE =  []
        D24DDT =  []
        D44DDT =  []
        Aldrin = []
        Dieldrin = []
        Endrin = []
        Telodrin = []
        Isodrin = []
        AHCH = []
        BHCH = []
        YHCH = []
        heptachloor = []
        heptachloorepoxide = []
        hexachloorbutadieen = []


        Parameters = ["Monster","Toeetsordeel",
                    "As (S)","Cd (S)","Cr (S)","Cu (S)",
                    "Hg (S)","Pb (S)","Ni (S)","Zn (S)",
                    "naftaleen","anthraceen","fenanthreen",
                    "fluorantheen","benz(a)anthr",
                    "chryseen","benzo(a)pyre",
                    "benzo(ghi)pe","benzo(k)fluo",
                    "indeno(123)p",
                    "PCB28","PCB52","PCB101","PCB118","PCB138",
                    "PCB153","PCB180","OLIE+FL(G)",
                    "HEXACHLB",
                    "2,4-DDD (o,p-DDD)","4,4-DDD (p,p-DDD)",
                    "2,4-DDE (o,p-DDE)", "4,4-DDE (p,p-DDE)",
                    "2,4-DDT (o,p-DDT)", "4,4-DDT (p,p-DDT)"
                    ,"Aldrin","Dieldrin",
                        "Endrin","Telodrin","Isodrin",
                    "A-HCH","B-HCH","Y-HCH",
                    "Heptachloor","heptachloorepoxide","hexachloorbutadieen"]

        #WorkDataFrame
        df = pd.DataFrame({})
        Monsters =[]
        Classification = []

        # iterate over files in the directory
        for filename in os.listdir(self.Path_Toetsingen):
            if "T3" in filename:
                f = os.path.join(self.Path_Toetsingen, filename)
                # Load the Excel file
                workbook = openpyxl.load_workbook(f)
                # Select the active worksheet
                worksheet = workbook.active
                for row in worksheet.iter_rows():
                    for cell in row:
                        if cell.value == "Monster":
                            next_cell = worksheet.cell(row=cell.row + 1, column=cell.column)
                            if isinstance(next_cell.value, datetime.datetime):
                                UselessVariable = next_cell.value.strftime('%d.%m.%Y')
                                # Extract day, month and year components
                                date_obj = datetime.datetime.strptime(UselessVariable, "%d.%m.%Y")
                                day = date_obj.day
                                month = date_obj.month
                                year_last_two_digits = str(date_obj.year % 100)  # Take the last two digits of the year and convert to string
                                M = '{}.{}.{}'.format(day, month, year_last_two_digits)
                                Monsters.append(M)
                            else: 
                                Monsters.append(next_cell.value)
                        elif cell.value == "Toetsoordeel":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 1)
                            Classification.append(next_cell.value)
                        elif cell.value == "arseen (As)":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            As.append(next_cell.value)
                        elif cell.value == "cadmium (Cd)":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            Cd.append(next_cell.value)
                        elif cell.value == "chroom (Cr)":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            Cr.append(next_cell.value)
                        elif cell.value == "koper (Cu)":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            Cu.append(next_cell.value)
                        elif cell.value == "kwik (Hg) (niet vluchtig)":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            Hg.append(next_cell.value)
                        elif cell.value == "lood (Pb)":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            Pb.append(next_cell.value)
                        elif cell.value == "nikkel (Ni)":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            Ni.append(next_cell.value)
                        elif cell.value == "zink (Zn)":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            Zn.append(next_cell.value)
                        elif cell.value == "naftaleen":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            naftaleen.append(next_cell.value)
                        elif cell.value == "anthraceen":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            anthraceen.append(next_cell.value)
                        elif cell.value == "fenantreen":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            fenanthreen.append(next_cell.value)
                        elif cell.value == "fluoranteen":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            fluorantheen.append(next_cell.value)
                        elif cell.value == "benzo(a)antraceen":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            benzaanthr.append(next_cell.value)
                        elif cell.value == "chryseen":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            benzoapyre.append(next_cell.value)
                        elif cell.value == "benzo(a)pyreen":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            chryseen.append(next_cell.value)
                        elif cell.value == "benzo(ghi)peryleen":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            benzoghipe.append(next_cell.value)
                        elif cell.value == "benzo(k)fluoranteen":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            benzokfluo.append(next_cell.value)
                        elif cell.value == "indeno(1,2,3-cd)pyreen":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            indeno123p.append(next_cell.value)
                        elif cell.value == "PCB - 28":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            PCB28.append(next_cell.value)
                        elif cell.value == "PCB - 52":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            PCB52.append(next_cell.value)
                        elif cell.value == "PCB - 101":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            PCB101.append(next_cell.value)
                        elif cell.value == "PCB - 118":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            PCB118.append(next_cell.value)
                        elif cell.value == "PCB - 138":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            PCB138.append(next_cell.value)
                        elif cell.value == "PCB - 153":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            PCB153.append(next_cell.value)
                        elif cell.value == "PCB - 180":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            PCB180.append(next_cell.value)
                        elif cell.value == "minerale olie (florisil clean-up)":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            OLIEFLG.append(next_cell.value)
                        elif cell.value == "hexachloorbenzeen":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            HEXACHLB.append(next_cell.value)
                        elif cell.value == "2,4-DDD (o,p-DDD)":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            D24DDD.append(next_cell.value)
                        elif cell.value == "4,4-DDD (p,p-DDD)":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            D44DDD.append(next_cell.value)
                        elif cell.value == "2,4-DDE (o,p-DDE)":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            D24DDE.append(next_cell.value)
                        elif cell.value == "4,4-DDE (p,p-DDE)":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            D44DDE.append(next_cell.value)
                        elif cell.value == "2,4-DDT (o,p-DDT)":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            D24DDT.append(next_cell.value)
                        elif cell.value == "4,4-DDT (p,p-DDT)":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            D44DDT.append(next_cell.value)
                        elif cell.value == "aldrin":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            Aldrin.append(next_cell.value)
                        elif cell.value == "dieldrin":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            Dieldrin.append(next_cell.value)
                        elif cell.value == "endrin":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            Endrin.append(next_cell.value)
                        elif cell.value == "telodrin":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            Telodrin.append(next_cell.value)
                        elif cell.value == "isodrin":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            Isodrin.append(next_cell.value)
                        elif cell.value == "alfa - HCH":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            AHCH.append(next_cell.value)
                        elif cell.value == "beta - HCH":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            BHCH.append(next_cell.value)
                        elif cell.value == "gamma - HCH (lindaan)":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            YHCH.append(next_cell.value)
                        elif cell.value == "heptachloor":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            heptachloor.append(next_cell.value)
                        elif cell.value == "heptachloorepoxide (cis)":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            heptachloorepoxide.append(next_cell.value)
                        elif cell.value == "hexachloorbutadieen":
                            next_cell = worksheet.cell(row=cell.row, column=cell.column + 6)
                            hexachloorbutadieen.append(next_cell.value)

                workbook.close()

        #First the general parameters

        df[Parameters[0]] = Monsters
        df[Parameters[1]] = Classification
        df[Parameters[2]] = As
        df[Parameters[3]] = Cd
        df[Parameters[4]] = Cr
        df[Parameters[5]] = Cu
        df[Parameters[6]] = Pb
        df[Parameters[7]] = Hg
        df[Parameters[8]] = Ni
        df[Parameters[9]] = Zn
        df[Parameters[10]] = naftaleen
        df[Parameters[11]] = anthraceen
        df[Parameters[12]] = fenanthreen
        df[Parameters[13]] = fluorantheen
        df[Parameters[14]] = benzaanthr
        df[Parameters[15]] = chryseen
        df[Parameters[16]] = benzoapyre
        df[Parameters[17]] = benzoghipe
        df[Parameters[18]] = benzokfluo
        df[Parameters[19]] = indeno123p
        df[Parameters[20]] = PCB28
        df[Parameters[21]] = PCB52
        df[Parameters[22]] = PCB101
        df[Parameters[23]] = PCB118
        df[Parameters[24]] = PCB138
        df[Parameters[25]] = PCB153
        df[Parameters[26]] = PCB180
        df[Parameters[27]] = OLIEFLG
        df[Parameters[28]] = HEXACHLB
        df[Parameters[29]] = D24DDD
        df[Parameters[30]] = D44DDD
        df[Parameters[31]] = D24DDE
        df[Parameters[32]] = D44DDE
        df[Parameters[33]] = D24DDT
        df[Parameters[34]] = D44DDT
        df[Parameters[35]] = Aldrin
        df[Parameters[36]] = Dieldrin
        df[Parameters[37]] = Endrin
        df[Parameters[38]] = Telodrin
        df[Parameters[39]] = Isodrin
        df[Parameters[40]] = AHCH
        df[Parameters[41]] = BHCH
        df[Parameters[42]] = YHCH
        df[Parameters[43]] = heptachloor
        df[Parameters[44]] = heptachloorepoxide
        df[Parameters[45]] = hexachloorbutadieen
        # Replace non-numeric entries with zeros in specified columns
        df = df.apply(pd.to_numeric, errors='coerce').fillna(0)
        df[Parameters[0]] = Monsters
        df[Parameters[1]] = Classification


        # Replace non-numeric entries with zeros
        dfPFAS = pd.read_excel(self.Path_PFAS)
        dfPFAS = dfPFAS.sort_values(by='Mengmonster')
        dfPFAS = dfPFAS.apply(pd.to_numeric, errors='coerce').fillna(0)
        dfPFAS2 = dfPFAS[["EtFOSAA (µg/kg ds)","MeFOSAA (µg/kg ds)","Concentratie (µg/kg ds)"]]
        dfPFAS2 = dfPFAS2.apply(pd.to_numeric, errors='coerce').fillna(0)
        # Add a new column with the maximum value of each row
        dfPFAS2['MaxValue'] = dfPFAS.iloc[:,:].apply(lambda row: row.max(), axis=1)
        
        # Load the Excel workbook
        workbook = openpyxl.load_workbook(self.Template)
        sheet = workbook.active

        col = 2

        #Begin with replacing 

        for index, row in df.iterrows():
            col = col + 1
            sheet.cell(row=1, column=col, value=row["Monster"])
            sheet.cell(row=3, column=col, value=Beoordelingen[row["Toeetsordeel"]])
            
            if dfPFAS.loc[index,"SOM PFOA (µg/kg ds)"] < 0.8 and dfPFAS.loc[index,"Som PFOS (µg/kg ds)"] < 3.70:
                sheet.cell(row=5, column=col, value="JA")
            else: 
                sheet.cell(row=5, column=col, value="NEE")

            sheet.cell(row=16, column=col, value=row[df.columns[2]])
            sheet.cell(row=17, column=col, value=row[df.columns[3]])
            sheet.cell(row=18, column=col, value=row[df.columns[4]])
            sheet.cell(row=19, column=col, value=row[df.columns[5]])
            sheet.cell(row=20, column=col, value=row[df.columns[6]])
            sheet.cell(row=21, column=col, value=row[df.columns[7]])
            sheet.cell(row=22, column=col, value=row[df.columns[8]])
            sheet.cell(row=23, column=col, value=row[df.columns[9]])
            sheet.cell(row=25, column=col, value=row[df.columns[10]])
            sheet.cell(row=26, column=col, value=row[df.columns[11]])
            sheet.cell(row=27, column=col, value=row[df.columns[12]])
            sheet.cell(row=28, column=col, value=row[df.columns[13]])
            sheet.cell(row=29, column=col, value=row[df.columns[14]])
            sheet.cell(row=30, column=col, value=row[df.columns[15]])
            sheet.cell(row=31, column=col, value=row[df.columns[16]])
            sheet.cell(row=32, column=col, value=row[df.columns[17]])
            sheet.cell(row=33, column=col, value=row[df.columns[18]])
            sheet.cell(row=34, column=col, value=row[df.columns[19]])
            sheet.cell(row=38, column=col, value=row[df.columns[20]])
            sheet.cell(row=39, column=col, value=row[df.columns[21]])
            sheet.cell(row=40, column=col, value=row[df.columns[22]])
            sheet.cell(row=41, column=col, value=row[df.columns[23]])
            sheet.cell(row=42, column=col, value=row[df.columns[24]])
            sheet.cell(row=43, column=col, value=row[df.columns[25]])
            sheet.cell(row=44, column=col, value=row[df.columns[26]])
            sheet.cell(row=47, column=col, value=row[df.columns[27]])
            sheet.cell(row=49, column=col, value=row[df.columns[28]])
            sheet.cell(row=52, column=col, value=row[df.columns[33]]+ row[df.columns[34]])
            sheet.cell(row=53, column=col, value=row[df.columns[29]]+ row[df.columns[30]])
            sheet.cell(row=54, column=col, value=row[df.columns[31]]+ row[df.columns[32]])
            sheet.cell(row=57, column=col, value=row[df.columns[35]])
            sheet.cell(row=58, column=col, value=row[df.columns[36]])
            sheet.cell(row=59, column=col, value=row[df.columns[37]])
            sheet.cell(row=61, column=col, value=row[df.columns[38]])
            sheet.cell(row=62, column=col, value=row[df.columns[39]])
            sheet.cell(row=64, column=col, value=row[df.columns[40]])
            sheet.cell(row=65, column=col, value=row[df.columns[41]])
            sheet.cell(row=66, column=col, value=row[df.columns[42]])
            sheet.cell(row=67, column=col, value=row[df.columns[43]])
            sheet.cell(row=68, column=col, value=row[df.columns[44]])
            sheet.cell(row=69, column=col, value=row[df.columns[45]])
            sheet.cell(row=72, column=col, value=dfPFAS.loc[index,"Som PFOS (µg/kg ds)"])
            sheet.cell(row=73, column=col, value=dfPFAS.loc[index,"SOM PFOA (µg/kg ds)"])
            sheet.cell(row=74, column=col, value=dfPFAS2.loc[index,"MaxValue"])
            
        workbook.save(os.path.join(self.Path_Toetsingen , self.ProjectNummer + "_Sluftertoets.xlsx"))
        workbook.close()


# Path_Toetsingen = r"P:\2023\23121 WNZ monitoring 2023\V1\09 Laboratorium\03 Toetsingen\EXCEL"
# Projectnummer = "WNZ_"
# Path_PFAS = r"P:\2023\23121 WNZ monitoring 2023\V1\09 Laboratorium\03 Toetsingen\EXCEL\WNZ_Vakken__Output_PFAS.xlsx"
# x = SlufterToets(Path_Toetsingen=Path_Toetsingen,Projectnummer=Projectnummer,Path_PFAS=Path_PFAS)
# x.RunTest()

#In[]: 