#In[]:
import tkinter as tk
from tkinter import filedialog

def select_folder():
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory()
    folder_path_text.delete("1.0", tk.END)
    folder_path_text.insert(tk.END, folder_path)

root = tk.Tk()
root.geometry("300x200")

folder_path_text = tk.Text(root, height=1)
folder_path_text.pack()

button = tk.Button(root, text="Select Folder", command=select_folder)
button.pack()

root.mainloop()

#In[]: