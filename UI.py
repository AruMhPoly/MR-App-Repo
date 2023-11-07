#In[]:
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter.ttk import *
from tkinter import messagebox
from PIL import ImageTk, Image
from tkinter import filedialog
from BoToVaResultaten import Botova
from ToetsingResultatenPFAS import PFAS
from ToetsingPFAsBijToepassen import PFASToepassing
from VerhoogdeRapportageGrenzen import VerhoogdeRapportageGrenzen
from Conserveringsopmerkingen import Conserveringsopmerkingen
from ToepassingsMogelijkheden import Toetssingsmogelijkheden
from SlufterToets import SlufterToets
window = tk.Tk()

# window.geometry("550x300+300+150")
window.resizable(width=True, height=True)

class Vista:

    def __init__(self, window):
        
        self.Project_Number = tk.StringVar()
        self.create_widgets()

    def create_widgets(self):
        
        # Welcome Frame

        welcome_frame = tk.LabelFrame(window, text="Mileu & Ruimte App", relief=tk.RIDGE, padx=5, pady=5)
        welcome_frame.grid(row=1, column=0,pady=7,columnspan = 3)
        welcome_frame.grid(sticky=tk.E + tk.W + tk.N + tk.S)

        # Logo

        image1 = Image.open(r"C:\Python\Coordinaten_App\Logo\Logo.jpeg")
        image1 = image1.resize((150, 220), Image.LANCZOS)
        test = ImageTk.PhotoImage(image1)
        label1 = tk.Label(image=test)
        label1.grid(row=0, column=0,columnspan = 3)
        label1.image = test

        # Welcome Text

        window.title("Milieu & Ruimte App")
        lbl_welcome = Label(welcome_frame,
                            text="Hoi! Wat is jouw projectnummer?")
        lbl_welcome.pack()

        # Frame Project gegevens
        global ProjectGegevensFrame
        ProjectGegevensFrame = tk.LabelFrame(window, text="Project gegevens", relief=tk.RIDGE)
        ProjectGegevensFrame.grid(row=2, sticky="nsew",pady = 5)

        # Frame Path gegevens
        global PathFrame
        PathFrame = tk.LabelFrame(window, text="Paths", relief=tk.RIDGE)
        PathFrame.grid(row=3, sticky="nsew",pady = 5)

        # Frame Créditos

        credit_frame = tk.LabelFrame(window,
                                            text="Contact persoon", relief=tk.RIDGE)
        credit_frame.grid(row=6, sticky="nsew", pady=5)

        #Field number of project

        Label(ProjectGegevensFrame,
              text="Projectnummer:              ").grid(row=1, column=0, pady=5, sticky="w")

        ttk.Entry(ProjectGegevensFrame, width=20, textvariable= self.Project_Number
                  ).grid(row=1,column=1,sticky=tk.E + tk.W + tk.N + tk.S, pady=5,)
        
        #Field Paths

        ##PATH TOETSINGEN (EXCEL)
        Label(PathFrame,
              text="Path Toetsingen (Excel):").grid(row=1, column=1, pady=5, sticky="w")

        self.PathToetsingenEntry = ttk.Entry(PathFrame, width=20)
        self.PathToetsingenEntry.grid(row=1, column=2, sticky=tk.E + tk.W + tk.N + tk.S, pady=5)
        self.PathToetsingenEntry.bind("<Button-1>", lambda event: self.SelectFolderPathToetsingen())

        ##PATH CERTIFICATEN (EXCEL)
        Label(PathFrame,
              text="Path Certificaten (Excel):").grid(row=2, column=1, pady=5, sticky="w")

        self.PathCertificateEntry = ttk.Entry(PathFrame, width=20)
        self.PathCertificateEntry.grid(row=2, column=2, sticky=tk.E + tk.W + tk.N + tk.S, pady=5)
        self.PathCertificateEntry.bind("<Button-1>", lambda event: self.SelectFolderPathCertificaten())

        ##PATH CERTIFICATEN (PDF)
        Label(PathFrame,
              text="Path Certificaten (PDF):").grid(row=3, column=1, pady=5, sticky="w")

        self.PathCertificatePdfEntry = ttk.Entry(PathFrame, width=20)
        self.PathCertificatePdfEntry.grid(row=3, column=2, sticky=tk.E + tk.W + tk.N + tk.S, pady=5)
        self.PathCertificatePdfEntry.bind("<Button-1>", lambda event: self.SelectFolderPathCertificatenPdf())
        
        # Button 
        Tabels_Button = tk.Button(PathFrame,text = "Tabelen crëeren"
                                        ,command = self.Tabels)
        Tabels_Button.grid(row=4,column=2, padx=10,pady=10)

        ##PATH CERTIFICATEN (EXCEL)
        Label(PathFrame,
              text="Dat was het!").grid(row=4, column=1, pady=5, sticky="w")


        # Texto de Créditos

        Label(credit_frame,text="Gecrëerd door: MH Poly").grid(row=0, sticky="nsew", pady=2)
        Label(credit_frame, text="E-mail: aru@mhpoly.com").grid(row=1, sticky="nsew", pady=2)
        Label(credit_frame, text="Telefoon nummer: +31 (0)164 245 566").grid(row=2, sticky="nsew", pady=2)



    def SelectFolderPathToetsingen(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.PathToetsingenEntry.delete(0, tk.END)
            self.PathToetsingenEntry.insert(0, folder_path)

    def SelectFolderPathCertificaten(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.PathCertificateEntry.delete(0, tk.END)
            self.PathCertificateEntry.insert(0, folder_path)

    def SelectFolderPathCertificatenPdf(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.PathCertificatePdfEntry.delete(0, tk.END)
            self.PathCertificatePdfEntry.insert(0, folder_path)

    def Tabels(self):

        Path_BoToVa = Botova(Path_Toetsingen=self.PathToetsingenEntry.get(),
               ProjectNummer=self.Project_Number.get()
               ).ResultatenBotova()
        Path_PFAS = PFAS(Path_Certificaten=self.PathCertificateEntry.get(),PathSave=self.PathToetsingenEntry.get(),
             ProjectNummer=self.Project_Number.get()).ResultatenPFAS()
        Path_PFAS_Toepassing = PFASToepassing(PFASPath=Path_PFAS,PathSave=self.PathToetsingenEntry.get(),ProjectNummer=self.Project_Number.get()).Toepassing()
        VerhoogdeRapportageGrenzen(PathCertifPdf=self.PathCertificatePdfEntry.get(),MonstersBoToVa=Path_BoToVa,MonsPFAS=Path_PFAS,
                                   PathSave=self.PathToetsingenEntry.get(),ProjectNummer=self.Project_Number.get()).Grenzen()
        
        #Sluftertoets
        SlufterToets(Path_Toetsingen=self.PathToetsingenEntry.get(),
                     Projectnummer=self.Project_Number.get(),
                     Path_PFAS=Path_PFAS).RunTest()

        
        Toetssingsmogelijkheden(MonstersBoToVa=Path_BoToVa,
                                MonsPFAS=Path_PFAS_Toepassing,
                                PathSave=self.PathToetsingenEntry.get(),
                                ProjectNummer=self.Project_Number.get()).Mogelijkheden()
        self.MessageBoxOverschriding()
        window.destroy()
        
    def MessageBoxOverschriding(self):
        List =Conserveringsopmerkingen(self.PathCertificatePdfEntry.get()).Overschrijding()
        if len(List)>0:
            message = "Let op! Er zijn monsters waar de conververingstermijnen overschreden zijn in de certificaten: {}".format(", ".join(List))
            # Display the messagebox
            messagebox.showwarning("Warning", message)

        else: 
            #Display a message if the list is empty
            messagebox.showinfo("Overschrijding conserveringstermijnen","In geen monster zijn de conververingstermijnen overschreden!")
            
        
# Create the entire GUI program
Vista(window)

# Start the GUI event loop
tk.mainloop()    

#In[]: 
