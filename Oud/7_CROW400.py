# In[]:
import fitz
import tabula
import pandas as pd 
import os

#In[]: 

Path_Crow = r'C:\Python\MR_APP\MR-App-Repo\CROW400'

#In[]:

#Input

Mon = []
Monsters = []
Klasse = []
VeiligheidsKlassen = ["Geen veiligheidsklasse","oranje",
                      "rood",'zwart']
Useless_List = ["Kadastraalnummer"]


#In[]: 

for filename in os.listdir(Path_Crow):

    f = os.path.join(Path_Crow, filename)

    # Open the PDF file
    pdf_file = fitz.open(f)

    # The number of the page
    page = pdf_file[0]

    # Extract the text from the page
    text = page.get_text()

    # Split the text into lines
    lines = text.split('\n')

    for word in Useless_List:
        for text in lines:        
            if word in text:
                Mon.append(text)
                break

    for word in VeiligheidsKlassen:
        for text in lines:        
            if word in text:
                Klasse.append(text)
                break
                    

for entry in Mon:
    Monsters.append(entry.split(":")[1].strip())
pdf_file.close()


#In[]: 

df_Out = pd.DataFrame({
    "Mengmonster": Monsters,
    "Veiligheidsklasse":Klasse,
})

#In[]: 

df_Out.to_excel(r"C:\Python\MR_APP\MR-App-Repo\Output\9.xlsx")

#In[]: