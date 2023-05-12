#In[]: 
import pandas as pd 
import os
import win32com.client as win32
import fitz
import openpyxl

#In[]: 

#Paths 

WP = r"C:\Python\MR_APP\Testen_DiverseVakken\ZINTUIGLIJK"
TP= r"C:\Python\MR_APP\Testen_DiverseVakken\TOETSINGEN\22218V1_Output_BoToVa.xlsx"

#In[]:

for filename in os.listdir(WP):
    if filename.endswith(".docx"):
        # Open the Word application and the document
        word = win32.Dispatch("Word.Application")
        word.Visible = False  # Set to True if you want to see the Word application
        filepath = os.path.join(WP, filename)
        doc = word.Documents.Open(filepath)

        # Loop through the tables in the document
        for i in range(len(doc.Tables)):
            table = doc.Tables[i]
            if "Deelmonsters" in table.Range.Text:
                break

        # Create an empty list to hold the table data
        data = []

        # Loop through the rows and cells of the table and append the cell text to the data list
        for row in table.Rows:
            row_data = []
            for cell in row.Cells:
                row_data.append(cell.Range.Text.strip())
            data.append(row_data)

        # Convert the data list to a Pandas DataFrame
        df = pd.DataFrame(data[1:], columns=data[0])
        df = df.replace('\r', '', regex=True)
        new_column_names = ['Analyse Mengmonster', 'Traject (m-mv)', 'Deelmonsters', 'Analysepakekt']
        df = df.rename(columns=dict(zip(df.columns, new_column_names)))
        df['Deelmonsters'] = df['Deelmonsters'].str.replace(r"\(.*?\)", "", regex=True)
        df = df.replace('', ',', regex=True)

        # save the document as PDF
        f = os.path.join(WP, filename + '.pdf')
        doc.SaveAs(f, FileFormat=17)

        # close the document and quit Word application
        doc.Close()
        word.Quit()

#In[]:

lines = []

# iterate over all files in the folder
for filename in os.listdir(WP):
    if filename.endswith(".pdf"):
        filepath = os.path.join(WP, filename)
        with fitz.open(filepath) as doc:
            # iterate over all pages in the PDF
            for page_num in range(doc.page_count):
                page = doc.load_page(page_num)
                text = page.get_text("text")
                # split the text content into lines and add them to the list
                lines.extend(text.splitlines())


#In[]: 


Monster_MHPoly = pd.read_excel(TP)


#In[]: 

Index = [] 
for i, entry in enumerate(lines):
    if "Hoofd grondsoort" in entry:
        Index.append(i)

#In[]:

Output_Monsters = []
Output_Grondstof = [] 
Useless_List = []
Row = 13942

for x in Index: 
    Useless_List = [elem for elem in lines[x-70:x] if elem in Monster_MHPoly]
    Output_Monsters.extend(Useless_List)
    for item in lines[x+2:x+2+len(Useless_List)]:
        Output_Grondstof.append(item)
    Useless_List = []


#In[]: 

import pandas as pd 

df = pd.DataFrame({'Monster': Output_Monsters, 'Grondsoort': Output_Grondstof})

#In[]: