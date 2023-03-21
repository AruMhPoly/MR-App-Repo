# In[]:

import os
import pandas as pd 
import openpyxl

#In[]

#PATHS - ONLY INPUT REQUIRED

#General path for all Toetsingen
Path= r"C:\Python\MR_APP\MR-App-Repo\Toetsingen"
#Path for only T3: To identify the parameters 
# That cause a bad soil quality
Path_T3 = r"C:\Python\MR_APP\MR-App-Repo\Toetsingen\Botova_1449959 + 1449958 + 1449957 + 1449956_T3.xlsx"
#Certificaten
Path_C = r"C:\Python\MR_APP\MR-App-Repo\Certificaten"

# In[]:
#WorkDataFrame
df = pd.DataFrame({})
#Empty lists
Columns_Names = []
Monsters =[]
Classification = []
Toetsingen = []

# In[]

# iterate over files in
# that directory
for filename in os.listdir(Path):
    f = os.path.join(Path, filename)
    # Load the Excel file
    workbook = openpyxl.load_workbook(f)
    # Select the active worksheet
    worksheet = workbook.active
    for row in worksheet.iter_rows():
        for cell in row:
            if cell.value == "Monster":
                next_cell = worksheet.cell(row=cell.row, column=cell.column + 1)
                Monsters.append(next_cell.value)
            if cell.value == "Toetsoordeel":
                next_cell = worksheet.cell(row=cell.row, column=cell.column + 1)
                Classification.append(next_cell.value)
            if cell.value == "Toetsing":
                next_cell = worksheet.cell(row=cell.row, column=cell.column + 1)
                Toetsingen.append(next_cell.value)

    New_Column=Toetsingen[0].split("-")[0].replace(".","").replace(" ","")
    Columns_Names.append(New_Column)
    # Add a new column for the new "Toetsing"
    df["Monster"] = Monsters
    df[New_Column] = Classification
    #Empty lists
    Columns_Names = []
    Monsters =[]
    Classification = []
    Toetsingen = []
# In[]:

#### What parameters produce that the klasse is B or NoT. 

# Load the Toetsing T3
workbook = openpyxl.load_workbook(Path_T3)

# Select the first worksheet
worksheet = workbook.worksheets[0]

# Create a list to store the row positions containing 
# "Parameters" that are in high concentrations

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

#This will allow me to restrict my area to a monster at
#the time

for x in range(0,len(parameter_rows)):
    end_rows.append(parameter_rows[x]- 2)

results = []
Monsters= []
Exceeded_Parameters = []

for x in range(len(parameter_rows)-1):

    # Iterate over the rows in the specified range
    for row in worksheet.iter_rows(min_row=parameter_rows[x], max_row=end_rows[x+1]):
        # Check if the row contains the value 'B' in the first column
        if row[-1].value == 'B':
            # If it does, add the value from the first column to the results list
            results.append(row[0].value)
        if row[-1].value == 'NoT':
            # If it does, add the value from the first column to the results list
            results.append(row[0].value)
        if row[-3].value == 'Monster':
            # If it does, add the value from the first column to the results list
            Monsters.append(row[-2].value)

    Exceeded_Parameters.append(','.join(results))
    results = []

# Add a new column for the results
df["Parameters Overschreden bij T3"] = Exceeded_Parameters
#In[]

Monsters_Lab = []
Monster_MHPoly = []
for filename in os.listdir(Path_C):
    f = os.path.join(Path_C, filename)
    # Load the Excel file
    workbook = openpyxl.load_workbook(f)
    # Select the active worksheet
    worksheet = workbook.active
    for row in worksheet.iter_rows(max_row=10):
        for cell in row:
            if cell.value in Monsters:
                next_cell = worksheet.cell(row=cell.row + 1, column=cell.column)
                Monsters_Lab.append(cell.value)
                Monster_MHPoly.append(next_cell.value)


df['Monster'].replace(to_replace=Monsters_Lab, value=Monster_MHPoly, inplace=True)
df.sort_values('Monster',inplace=True)
#In[]

df.to_excel(r"C:\Python\MR_APP\MR-App-Repo\Output\2.xlsx")
# In[]
