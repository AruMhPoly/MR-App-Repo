# In[]:

import os
import pandas as pd 
import openpyxl
import datetime

class Botova:
    def __init__(self,Path_Toetsingen,ProjectNummer):

        self.Path_Toetsingen = Path_Toetsingen
        self.ProjectNummer = ProjectNummer

    def ResultatenBotova(self):
        #WorkDataFrame
        df = pd.DataFrame({})
        #Empty lists
        Columns_Names = []
        Monsters =[]
        Monsters_Temporal = []
        Classification = []
        Toetsingen = []

        # iterate over files in the directory
        for filename in os.listdir(self.Path_Toetsingen):
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
                            Monsters_Temporal.append(M)
                        else: 
                            Monsters_Temporal.append(next_cell.value)

                    if cell.value == "Toetsoordeel":
                        next_cell = worksheet.cell(row=cell.row, column=cell.column + 1)
                        Classification.append(next_cell.value)
                    if cell.value == "Toetsing":
                        next_cell = worksheet.cell(row=cell.row, column=cell.column + 1)
                        Toetsingen.append(next_cell.value)

            x =f.split("_")[-1].split(".")[0]
            for i in range(len(Monsters_Temporal)):
                Columns_Names.append(x)

            Monsters.extend(Monsters_Temporal)
            Monsters_Temporal = []

        df["Monster"] = Monsters
        df["Toetsing"] = Columns_Names
        df['Classification'] = Classification
        df2 = df.pivot_table(index='Monster', columns='Toetsing', values='Classification', aggfunc=lambda x: ', '.join(x))

        Temporal_Exceeded_Parameters = [] 
        Exceeded_Parameters = []
        for filename in os.listdir(self.Path_Toetsingen):
            f = os.path.join(self.Path_Toetsingen, filename)
            x =f.split("_")[-1].split(".")[0]
            if x == "T3":
                #### What parameters produce that the klasse is B or NoT. 
                # Load the Toetsing T3
                workbook = openpyxl.load_workbook(f)
                # Select the first worksheet
                worksheet = workbook.worksheets[0]

                # Create a list to store the row positions containing "Parameters" that are in high concentrations

                parameter_rows = []

                # Loop through all rows in the worksheet
                for row_index, row in enumerate(worksheet.iter_rows(), start=1):

                    # Loop through all cells in the row
                    for cell in row:

                        # Check if the cell contains the word "Parameters"
                        if cell.value == "Parameters":

                            # If it does, add the row position to the list
                            parameter_rows.append(row_index)

                            # Check if the cell contains the word "Parameters"

                        if cell.value == "Legenda":

                            # If it does, add the row position to the list
                            parameter_rows.append(row_index)
                
                end_rows=[]

                # #This will allow me to restrict my area to a monster at the time

                for x in range(0,len(parameter_rows)):
                    end_rows.append(parameter_rows[x]- 2)
                results = []
                Concentratie = []
                # Iterate through the first rows 
                for row in range(parameter_rows[0], parameter_rows[1]):
                    # Iterate through each column in the row
                    for column in range(1, worksheet.max_column + 1):
                        # Get the value of the cell
                        cell_value = worksheet.cell(row=row, column=column).value
                        # Check if the value is "T.Oordel"
                        if cell_value == "T.Oordeel":
                            FirstColumn = column 
                            # Exit the loop, since we've found what we're looking for
                            break

                for x in range(len(parameter_rows)-1):

                    # Iterate over the rows in the specified range
                    for row in worksheet.iter_rows(min_row=parameter_rows[x], max_row=end_rows[x+1]):
                        # Check if the row contains the value 'B' in the last column
                        # print(row[-1].value)
                        if row[FirstColumn-1].value == 'B':
                            # If it does, add the value from the first column to the results list
                            results.append(row[0].value)
                            Concentratie.append(str(row[-2].value) + " mg/kg ds  ")
                        if row[FirstColumn-1].value == 'NoT':
                            # If it does, add the value from the first column to the results list
                            results.append(row[0].value)
                            Concentratie.append(str(row[-2].value) + " mg/kg ds  ")

                    TemporalList = list(zip(results, Concentratie)) 
                    Temporal_Exceeded_Parameters.append(' - '.join([f"{x}:{y}" for x, y in TemporalList]))
                    TemporalList = []
                    results = []
                    Concentratie = []
                    Exceeded_Parameters.extend(Temporal_Exceeded_Parameters)
                    Temporal_Exceeded_Parameters = []

            else:
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
                                Monsters_Temporal.append(M)
                            else: 
                                Monsters_Temporal.append(next_cell.value)


                for i in range(len(Monsters_Temporal)):
                    Temporal_Exceeded_Parameters.append(' ')

                Exceeded_Parameters.extend(Temporal_Exceeded_Parameters)
                Temporal_Exceeded_Parameters = []
                Monsters_Temporal = []

        # Add a new column for the results
        df["Monster"] = Monsters
        df["Toetsing"] = Columns_Names
        df['Classification'] = Classification
        df["Parameters Overschreden bij T3"] = Exceeded_Parameters

        df2 = df.pivot_table(index='Monster', columns='Toetsing', 
                            values=['Classification', 'Parameters Overschreden bij T3'], aggfunc=lambda x: ', '.join(x))
        Path_Save = os.path.join(self.Path_Toetsingen,self.ProjectNummer + '_Output_BoToVa.xlsx')
        df2.to_excel(Path_Save)
        return Path_Save 

#In[]: 
Path= r'C:\Python\MR_APP\Testen_DiverseVakken\TOETSINGEN'
df = Botova(Path_Toetsingen=Path,ProjectNummer="22218V1").ResultatenBotova()

#In[]: