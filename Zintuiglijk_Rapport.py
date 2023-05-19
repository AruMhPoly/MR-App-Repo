#In[]: 
import pandas as pd 
import os
import fitz
import docx2txt

#In[]: 

#Paths 

WP = r"C:\Python\MR_APP\Testen_DiverseVakken\ZINTUIGLIJK"
TP= r"C:\Python\MR_APP\Testen_DiverseVakken\TOETSINGEN\22218V1_Output_BoToVa.xlsx"

#In[]:

for filename in os.listdir(WP):
    if filename.endswith(".docx"):
        
        # Path to the Word file
        word_file = os.path.join(WP, filename)
        break 

# Path to the output text file
text_file = os.path.join(WP, filename + ".txt")

# Extract text from the Word file
text = docx2txt.process(word_file)

# Write the text to the output file
with open(text_file, "w") as f:
    f.write(text)    


#In[]: 


Monster_MHPoly = pd.read_excel(TP)

List_Monsters = Monster_MHPoly.iloc[2:,0].tolist()



#In[]: 

monster = []
traject = []
deelmonsters = []
analysepakket = []

with open(text_file, 'r') as file:
    lines = file.readlines()

# Find the start and end rows
start_row = None
end_row = None
for i, line in enumerate(lines):
    if "Tabel 4: Monsterselectie" in line:
        start_row = i
    elif "Tabel 5: Analyses grondwater" in line:
        end_row = i
        break


for x in List_Monsters: 

    for i in range(start_row, end_row):
        if str(x) in lines[i]:
            monster.append(lines[i].strip())
            traject_row = i + 2
            traject.append(lines[traject_row].strip())
            deelmonsters_start_row = traject_row + 2
            deelmonsters_end_row = None
            for j in range(deelmonsters_start_row, end_row):
                if lines[j].strip() == "":
                    deelmonsters_end_row = j
                    break
            deelmonsters.append(",".join(lines[deelmonsters_start_row:deelmonsters_end_row]).strip())
            analysepakket_row = deelmonsters_end_row + 1
            analysepakket.append(lines[analysepakket_row].strip())
            break

deelmonsters_clean = []
for deelmonster in deelmonsters:
    deelmonster_clean = []
    for part in deelmonster.split("\n"):
        part_clean = part.split(" ")[0]
        deelmonster_clean.append(part_clean)
    deelmonsters_clean.append("".join(deelmonster_clean))


#In[]: 

GrondSoort = []
Monster = []

with open(text_file, "r") as file:
    lines = file.readlines()

# Find the row with the "Hoofd grondsoort" text
hoofd_grondsoort_row = None
for i, line in enumerate(lines):
    if "Hoofd grondsoort" in line:
        hoofd_grondsoort_row = i

        # Extract the values from the rows below "Hoofd grondsoort"
        if hoofd_grondsoort_row is not None:
            # Calculate the row indices for the values we want to extract
            GS1 = hoofd_grondsoort_row + 4
            GS2 = GS1 + 2
            GS3 = GS2 + 2

            # Extract the values for the material
            Grondstof1 = lines[GS1].strip()
            Grondstof2 = lines[GS2].strip()
            Grondstof3 = lines[GS3].strip()

            #Extract the Monsters names

            M1 = hoofd_grondsoort_row -110 +4
            M2 = M1 + 2
            M3 = M2 + 2

            #Extract the values

            Monster1 = lines[M1].strip()
            Monster2 = lines[M2].strip()
            Monster3 = lines[M3].strip()
            

            # Add the values to the lists
            GrondSoort.extend([Grondstof1, Grondstof2, Grondstof3])
            Monster.extend([Monster1,Monster2,Monster3])
            # print(Monster1,Monster2,Monster3)
#In[]:

df1 = pd.DataFrame(
    {"Mengmonster": monster,
    "Traject (m-mv)": traject,
    "Deelmonsters": deelmonsters_clean,
    "Analysepakket": analysepakket})


df2 = pd.DataFrame({"Mengmonster":Monster,
                    "Hoofd grondsoort": GrondSoort,})

#In[]:

merged_df = pd.merge(df1, df2, on="Mengmonster", how="inner")
merged_df.set_index("Mengmonster")

merged_df.to_excel(r'C:\Python\MR_APP\Testen_DiverseVakken\TOETSINGEN\ZINTUIGLIJK.xlsx')

#In[]: