#In[]: 

import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk


#In[]: 


# create the main window

root = tk.Tk()
root.title("Milieu & Ruimte App")

# add logo image to the window
logo_path = r"C:\Python\MR_APP\MR-App-Repo\UserInterface\Logo.jpeg"
logo_img = Image.open(logo_path)
logo_tk = ImageTk.PhotoImage(logo_img)
logo_label = tk.Label(root, image=logo_tk)
logo_label.pack()

# add field for selecting the "Toetsingen" folder
toetsingen_label = tk.Label(root, text="Selecteer de Toetsingen folder:")
toetsingen_label.pack()
toetsingen_path = tk.StringVar()
toetsingen_entry = tk.Entry(root, textvariable=toetsingen_path)
toetsingen_entry.pack()
def browse_toetsingen():
    selected_path = filedialog.askdirectory()
    toetsingen_path.set(selected_path)
browse_toetsingen_button = tk.Button(root, text="Bladeren", command=browse_toetsingen)
browse_toetsingen_button.pack()

# add field for selecting the "Toetsing 3" file
toetsing_label = tk.Label(root, text="Selecteer het 'Toetsing 3' bestand:")
toetsing_label.pack()
toetsing_path = tk.StringVar()
toetsing_entry = tk.Entry(root, textvariable=toetsing_path)
toetsing_entry.pack()
def browse_toetsing():
    selected_path = filedialog.askopenfilename()
    toetsing_path.set(selected_path)
browse_toetsing_button = tk.Button(root, text="Bladeren", command=browse_toetsing)
browse_toetsing_button.pack()

# add field for selecting the "Certificaten" folder
certificaten_label = tk.Label(root, text="Selecteer de Certificaten folder:")
certificaten_label.pack()
certificaten_path = tk.StringVar()
certificaten_entry = tk.Entry(root, textvariable=certificaten_path)
certificaten_entry.pack()
def browse_certificaten():
    selected_path = filedialog.askdirectory()
    certificaten_path.set(selected_path)
browse_certificaten_button = tk.Button(root, text="Bladeren", command=browse_certificaten)
browse_certificaten_button.pack()

# run the main event loop
root.mainloop()

#In[]: