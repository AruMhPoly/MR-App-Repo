# In[]:

import os
import pandas as pd 
import tabula
import camelot


#In[]:

# specify the path to your PDF file
pdf_path = r"C:\Python\MR_APP\Input\botova_T1 WB.pdf"

# read the table(s) from the PDF file and store them in a DataFrame
df = tabula.read_pdf(pdf_path, pages='all')