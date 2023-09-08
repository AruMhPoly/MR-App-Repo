#In[]:

import win32com.client as win32
import os
import glob


#In[]: 

Excel_Path = r'C:\Python\ToetsingKaderLokaleWaarden\TESTEN\22096V1'
#In[]:

for filename in os.listdir(Excel_Path):
    
    f = os.path.join(Excel_Path, filename)

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.DisplayAlerts = False
    print(f)
    workbook = excel.Workbooks.Open(f)
    workbook.SaveAs(os.path.join(Excel_Path,filename + "x"), FileFormat=51)  # 51 is the XLSX file format
    workbook.Close()
    excel.Quit()

#In[]: 

# Use glob to find all files with .xls extension in the folder
files_to_delete = glob.glob(os.path.join(Excel_Path, "*.xls"))

# Use a loop to delete each file
for file_path in files_to_delete:
    os.remove(file_path)

#In[]: 