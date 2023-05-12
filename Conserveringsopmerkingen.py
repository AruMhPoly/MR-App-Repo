import fitz
import pandas as pd 
import os

class Conserveringsopmerkingen: 
    def __init__(self,PathCertifPdf):
        self.PathCertifPdf = PathCertifPdf

    def Overschrijding(self):

        UL = [] 

        for filename in os.listdir(self.PathCertifPdf):

            f = os.path.join(self.PathCertifPdf, filename)

            with open(f, 'rb') as pdf_file:
                # Create a PdfReader object to read the PDF file
                pdf_reader = fitz.open(stream=pdf_file.read(), filetype="pdf")
                # Iterate through each page in the PDF file
                for page_num in range(pdf_reader.page_count):
                    # Get the text content of the page
                    page = pdf_reader[page_num]
                    page_text = page.get_text()
                    # Check if both texts appear in the page
                    if 'conserveringsopmerkingen' in page_text:
                        UL.append(filename)

        return UL
    
#In[]:

# Test = conserveringsopmerkingen(r'P:\2022\22218 WNZ diverse vakken LN 2023\V1\07 Laboratorium\2 Certificaten\RA01\PDF').Overschrijding()

#In[]: 