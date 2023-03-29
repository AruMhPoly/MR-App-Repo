# In[]:
import fitz
import tabula
import pandas as pd 
import os

#In[]: 

#Functions 

# Define function to split and return unique values
def get_unique(x):
    values = set(x.str.split(", ").sum())
    values.discard('')
    return ", ".join(sorted(values))

#Input
#Path certificaten
#PDF
Path_Certif = r"P:\2022\22218 WNZ diverse vakken LN 2023\V1\07 Laboratorium\2 Certificaten\RA04\PDF"
#Monsters
Mons = r"P:\2022\22218 WNZ diverse vakken LN 2023\V1\07 Laboratorium\3 Toetsingen\RA04\EXCEL\Output_Botova.xlsx"
Mons_PFAS = r"P:\2022\22218 WNZ diverse vakken LN 2023\V1\07 Laboratorium\3 Toetsingen\RA04\EXCEL\Output_PFAS.xlsx"
#Save Path
Path_Toetsingen= r"P:\2022\22218 WNZ diverse vakken LN 2023\V1\07 Laboratorium\3 Toetsingen\RA04\EXCEL"


# Empty lists to store results

#Monster
M = []
# Stof
S = []
# Bericht
B = []
#In[]:
for filename in os.listdir(Path_Certif):

    f = os.path.join(Path_Certif, filename)

    Pages=[]
    # Open the PDF file in binary mode
    with open(f, 'rb') as pdf_file:
        # Create a PdfReader object to read the PDF file
        pdf_reader = fitz.open(stream=pdf_file.read(), filetype="pdf")
        # Iterate through each page in the PDF file
        for page_num in range(pdf_reader.page_count):
            # Get the text content of the page
            page = pdf_reader[page_num]
            page_text = page.get_text()
            # Check if both texts appear in the page
            if 'Opmerkingen m.b.t. analyses' in page_text:
                Pages.append(page_num)
            if 'OLIE-ONDERZOEK' in page_text:
                Pages.append(page_num)
                break
    #Name of monsters
    Monster_Names = pd.read_excel(Mons)["Monster"].to_list()
    Monster_PFAS = pd.read_excel(Mons_PFAS)["Mengmonster"].to_list()
    Monster_Names.extend(Monster_PFAS)
    Monster_Names = list(set(Monster_Names))
    Monster_Names = [str(num) for num in Monster_Names]

    # Open the PDF file
    pdf_file = fitz.open(f)

    for x in range(Pages[0],Pages[1]):

        # The number of the page
        page = pdf_file[x]

        # Extract the text from the page
        text = page.get_text()

        # Split the text into lines
        lines = text.split('\n')

        # Position where the monster are located
        Pos = []
        for i, line in enumerate(lines):
            for monster in Monster_Names:
                if monster in line:
                    Pos.append(i)

        for i, line in enumerate(lines):
            if "Tabel" in line and "van" in line:
                Pos.append(i)

        Pos.sort()

        for x in range(0,len(Pos)-1):
            for i, line in enumerate(lines[Pos[x]:Pos[x+1]]):
                if "verhoogde rapportagegrens" in line:
                    M.append(lines[Pos[x]])
                    S.append(lines[Pos[x] + i-2])
                    B.append(lines[Pos[x] + i])

    df_Out = pd.DataFrame({
        "Mengmonster": M,
        "Parameters": S,
        "Oorzak": B
    })

    # Group rows by "Mengmonster" and join values in other columns
    grouped = df_Out.groupby("Mengmonster").agg({"Parameters": ", ".join, "Oorzak": ", ".join})

    # Reset index 
    result = grouped.reset_index()

    # Apply function to "Oorzak" column
    result["Oorzak"] = result["Oorzak"].apply(lambda x: get_unique(pd.Series(x)))
    result["Parameters"] = result["Parameters"].apply(lambda x: get_unique(pd.Series(x)))

# In[]:

df_Out.to_excel(os.path.join(Path_Toetsingen,'VerhoogdeRapportageGrenzen.xlsx'))

#In[]: