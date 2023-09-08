#In[]:

import os
import win32com.client as win32
import fitz
import openpyxl

#In[]:

Word_Path = r"C:\Python\MR_APP\MR-App-Repo\Word"
Path_Toetsingen= r"C:\Python\MR_APP\MR-App-Repo\Toetsingen\Botova_1449959 + 1449958 + 1449957 + 1449956_T3.xlsx"
Path_C = r"C:\Python\MR_APP\MR-App-Repo\Certificaten"

#In[]:

# # create a new instance of Word application
# word = win32.gencache.EnsureDispatch('Word.Application')

# # open the document
# doc = word.Documents.Open(Word_Path)

# # save the document as PDF
# pdf_path = os.path.splitext(doc.FullName)[0] + ".pdf"
# doc.ExportAsFixedFormat(pdf_path, win32.constants.wdExportFormatPDF)

# # close the document and quit Word application
# doc.Close()
# word.Quit()

#In[]:

lines = []

# iterate over all files in the folder
for filename in os.listdir(Word_Path):
    if filename.endswith(".pdf"):
        filepath = os.path.join(Word_Path, filename)
        with fitz.open(filepath) as doc:
            # iterate over all pages in the PDF
            for page_num in range(doc.page_count):
                page = doc.load_page(page_num)
                text = page.get_text("text")
                # split the text content into lines and add them to the list
                lines.extend(text.splitlines())


#In[]: 


Monster_MHPoly = []
for filename in os.listdir(Path_C):
    f = os.path.join(Path_C, filename)
    # Load the Excel file
    workbook = openpyxl.load_workbook(f)
    # Select the active worksheet
    ws  = workbook.active
    # Iterate through the first 10 rows
    for row in ws.iter_rows(min_row=1, max_row=10):
        # Check if column A contains the target text
        if row[0].value == "Projectomschrijving":
            # Extract values after column C
            Monster_MHPoly.extend(cell.value for cell in row[3:])
        

#In[]:  

Useless_List = []

for val in Monster_MHPoly:
    if val != None :
        Useless_List.append(val)

Monster_MHPoly = Useless_List
Useless_List = []
for x in Monster_MHPoly:
    Useless_List.append(x + " ")

Monster_MHPoly = Useless_List


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