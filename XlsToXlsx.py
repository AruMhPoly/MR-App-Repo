#In[]:

import win32com.client as win32
import os
import glob


#In[]: 

Path = r'P:\2023\23041 Perenlaantje Volgerlanden\B1\07 Laboratorium\3 Toetsingen\Excel'

#In[]:

class ExcelConverter ():

    def __init__(self,Pathcertificaten):
        self.Pathcertificaten = Pathcertificaten

    def Convert(self):
        
        for filename in os.listdir(self.Pathcertificaten):
            
            f = os.path.join(self.Pathcertificaten, filename)

            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.DisplayAlerts = False
            print(f)
            workbook = excel.Workbooks.Open(f)
            workbook.SaveAs(os.path.join(self.Pathcertificaten,filename + "x"), FileFormat=51)  # 51 is the XLSX file format
            workbook.Close()
            excel.Quit()

        # Use glob to find all files with .xls extension in the folder
        files_to_delete = glob.glob(os.path.join(self.Pathcertificaten, "*.xls"))

        # Use a loop to delete each file
        for file_path in files_to_delete:
            os.remove(file_path)

#In[]: 