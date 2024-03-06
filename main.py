from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

from os import path
from PIL import ImageTk, Image

from reportBook import *

# COLOR
COLOR_WHITE = "#ffffff"

COLOR_SECOND = "#F4D03F"
COLOR_SECOND_CLICK = "#f4c23f"

COLOR_LINE = "#e8e8e5"

COLOR_BLACK = "#222222"
COLOR_LIGHT_GREY = "#F5F5F2"
COLOR_GREY = "#a1a19f"


class Application(Tk):

    def __init__(self):
        super().__init__()

        # PARAMETRE APPLICATION
        self.title("Livret")
        self.geometry(f"{round(self.winfo_screenwidth() * 0.8)}x{round(self.winfo_screenheight()* 0.8)}")
        self.state('zoomed')
        self.minsize(self.ProportionScreen(40), self.ProportionScreen(80, width=False))
        self.iconphoto(True, PhotoImage(file='pictures/icon.png'))

        # CONSTANTES
        self.ROOT = self

        self.PICTURES = {
            "home" : self.LaunchPicture("pictures/home.png", proportion=3),           
            "document" : self.LaunchPicture("pictures/document.png", proportion=3),
            "folder" : self.LaunchPicture("pictures/folder.png", proportion=3),
            "exit" : self.LaunchPicture("pictures/exit.png", proportion=2.5),
            "reportBook" : self.LaunchPicture("pictures/reportBook.png", proportion=12),
            "add" : self.LaunchPicture("pictures/add.png", proportion=12),
            "multipleAdd" : self.LaunchPicture("pictures/multipleAdd.png", proportion=12),
            "addBtn" : self.LaunchPicture("pictures/addBtn.png", proportion=2),
            "cross" : self.LaunchPicture("pictures/cross.png", proportion=2),
            "person" : self.LaunchPicture("pictures/person.png", proportion=6),
            "people" : self.LaunchPicture("pictures/people.png", proportion=6),
            "desktop": self.LaunchPicture("pictures/desktop.png", proportion=2.5),
            "laptop": self.LaunchPicture("pictures/laptop.png", proportion=2.5),
            "openFolder" : self.LaunchPicture("pictures/openFolder.png", proportion=2.5),
            "documentB" : self.LaunchPicture("pictures/documentB.png", proportion=1.5),
            "time" : self.LaunchPicture("pictures/time.png", proportion=1.5),
            "menu" : self.LaunchPicture("pictures/menu.png", proportion=3),
            "save" : self.LaunchPicture("pictures/save.png", proportion=1.5),
            "left" : self.LaunchPicture("pictures/left.png", proportion=6),
            "right" : self.LaunchPicture("pictures/right.png", proportion=6),
            "arrow" : self.LaunchPicture("pictures/arrow.png", proportion=4),

        }

        # style bouton
        self.BTN_STYLE_SECOND_COLOR =   {"<Enter>": {"bg" : COLOR_SECOND_CLICK},
            "<Leave>": {"bg" : COLOR_SECOND},
            "<Leave>": {"bg" : COLOR_SECOND},
            "<FocusIn>": {"bg" : COLOR_SECOND_CLICK},
            "<Activate>": {"bg" : COLOR_SECOND_CLICK},
            "<FocusOut>": {"bg" : COLOR_SECOND},}

        self.BTN_STYLE_WHITE_COLOR =   {"<Enter>": {"bg" : COLOR_LIGHT_GREY},
            "<Leave>": {"bg" : COLOR_LIGHT_GREY},
            "<Leave>": {"bg" : COLOR_WHITE},
            "<FocusIn>": {"bg" : COLOR_LIGHT_GREY},
            "<Activate>": {"bg" : COLOR_LIGHT_GREY},
            "<FocusOut>": {"bg" : COLOR_WHITE},}

        # VARIABLES
        self.mode_display = 1
        self.menu_chose = "home"

        self.display = None
        self.file_name = None
        self.active = False

        self.number_subjects = 1
        self.position_list = 0

        self.file_name = "back.xlsx"
        self.file_used = "back.xlsx"

        self.UpdateWindow()

# ============= FONCTIONS SET ===============

    def SetFrame(self, parent, width, height, cell, span = (1,1), sk = None, background=COLOR_WHITE, ipadx=0, ipady=0, padx=0, pady=0, sizeText= False, cursor=None):
        # evite les erreurs si le parent n'as pas de taille
        try: w = self.CheckSettingSize(width, parent["width"])
        except: w = 0
        try: h = self.CheckSettingSize(height, parent["height"], width=False)
        except:  h = 0

        frameX = Frame(parent, bg=background, width= w,  height= h, cursor=cursor)
  
        frameX.grid(column=cell[0], row=cell[1], columnspan=span[0], rowspan=span[1], sticky=sk, padx=padx, pady=pady, ipadx=ipadx, ipady=ipady)
        frameX.grid_propagate(sizeText)
        return frameX

    def SetLabel(self, parent, text, fontSize, cell, color=COLOR_BLACK, background=COLOR_WHITE, span = (1,1), sk = None, ipadx=10, ipady=10, fontStyle="", image=None, placementText=TOP,
    padx=0, pady=0):
        labelX = Label(parent, text=text, fg=color,font=f"Arial {self.SizeFont(fontSize)} {fontStyle}", bg=background, image=image, compound=placementText)
        labelX.grid(column=cell[0], row=cell[1], columnspan=span[0], rowspan=span[1], sticky=sk, ipadx=ipadx, ipady=ipady, padx=padx, pady=pady)
        return labelX
        
    def SetText(self, parent, fontSize, cell, color=COLOR_BLACK, background=COLOR_WHITE, span = (1,1), sk = None, ipadx=10, ipady=10, fontStyle="",
    padx=0, pady=0,height= 1, width=5):
        textX = Text(parent, fg=color,font=f"Arial {self.SizeFont(fontSize)} {fontStyle}", bg=background,height=height, width=width, relief=FLAT)
        textX.grid(column=cell[0], row=cell[1], columnspan=span[0], rowspan=span[1], sticky=sk, ipadx=ipadx, ipady=ipady, padx=padx, pady=pady)
        return textX
    
    def SetButton(self, parent, cell, color=COLOR_BLACK, background=COLOR_WHITE, span = (1,1), sk = None ,ipadx=0, ipady=0,imageResize = False, image=None, text=None, fontSize=11, 
    sizeText=False, placementText=None, activeBg=COLOR_WHITE, activeFg=COLOR_BLACK, var=None, command=None, width=None, height=None, padx=0, pady=0):
        try: w = self.CheckSettingSize(width, parent["width"])
        except: w = 0
        try: h = self.CheckSettingSize(height, parent["height"], width=False)
        except:  h = 0
        if imageResize == True:
            image.resize((100,100))

        buttonX = Button(parent, fg=color,font=f"Arial {self.SizeFont(fontSize)}", bg=background, text=text, image=image, compound=placementText, relief=FLAT, activebackground=activeBg,
         activeforeground=activeFg, cursor="hand2", command=command, width=w, height=h)
        buttonX.grid(column=cell[0], row=cell[1], columnspan=span[0], rowspan=span[1],sticky=sk, ipadx=ipadx, ipady=ipady, padx=padx, pady=pady)

        buttonX.grid_propagate(sizeText)

        return buttonX  

    def SetDisplay(self):
        if self.display == None:
            self.display = Frame(
                self, bg=COLOR_WHITE, background=COLOR_WHITE)
            self.display.grid(row=0, column=0, sticky=(N,S,W,E))
            self.grid_columnconfigure(0, weight=5)
            self.rowconfigure(0, weight=5)

    def SetLine(self, parent, cell ,lenght, thickness, background="#fafafa", sk=None, horizontal=True):
        if horizontal == True:
            try: l = self.CheckSettingSize(lenght, parent["width"])
            except: l = 0
            
            canvasX = Canvas(parent, width=l, height=thickness, bg=background, relief=FLAT, borderwidth=0, highlightthickness=0)
            canvasX.grid(column=cell[0], row=cell[1], padx=3, sticky=sk)
        else:
            try: l = self.CheckSettingSize(lenght, parent["height"])
            except: l = 0
            
            canvasX = Canvas(parent, width=thickness, height=l, bg=background, relief=FLAT, borderwidth=0, highlightthickness=0)
            canvasX.grid(column=cell[0], row=cell[1], pady=5, sticky=sk)
            
        return canvasX

    def SetEntry(self, parent, cell, width, sk=None, fontSize=11, fontStyle="", padx=0, pady=1, ipadx=5, ipady=5, background=None, hiLig=1,
     hiLigBg="#8D8D8D"):
        try: w = self.CheckSettingSize(width, parent["width"])
        except: w = 0     

        XEntry = Entry(parent, relief=FLAT, width=w, font=f"Arial {self.SizeFont(fontSize)} {fontStyle}", bg=background, highlightthickness = hiLig,
         highlightcolor = hiLigBg)
        XEntry.grid(column=cell[0], row=cell[1], sticky=sk, padx=padx, pady=pady, ipadx=ipadx,ipady = ipady )

        return XEntry
            
    def SetSpinBox(self, parent, cell, sk=None, fontSize=11, fontStyle="", padx=0, pady=1, ipadx=5, ipady=5, background=None, hiLig=1,
        hiLigBg="#8D8D8D", size_number=(0,2)):

        spin_box_x = Spinbox(parent, from_=size_number[0], to=size_number[1], font=f"Arial {fontSize} {fontStyle}", bg=background, relief=FLAT, width=len(str(size_number[1]))+1)
        spin_box_x.grid(column=cell[0], row=cell[1], sticky=sk, padx=padx, pady=pady, ipadx=ipadx,ipady = ipady )
        return spin_box_x

    def SetFrameScroll(self, parentCanvas, cellCanvas, parentScrollBar, cellScrollBar, skCanvas=(N,S,W,E), skScrollBar=(N,S), orient= VERTICAL, widthFrame=True,
     heightFrame=False, x=False, y=False, background="white", scrollBarUsed=True):
       
        # Canvas
        canvas = Canvas(parentCanvas, bg=background, relief=FLAT, bd=0, highlightthickness=0)
        canvas.grid(column=cellCanvas[0], row=cellCanvas[1], sticky=skCanvas, padx=1, pady=1)

        # Frame
        frame_contener = Frame(canvas, background=background)
        
        # ScroolBar
        if scrollBarUsed == True:
            scroll_bar = Scrollbar(parentScrollBar, orient=orient, command= canvas.yview)                 
            scroll_bar.grid(column=cellScrollBar[0], row=cellScrollBar[1], sticky=skScrollBar)
            canvas.configure(yscrollcommand=scroll_bar.set)

        # Setting For Scroll
        canvas_frame = canvas.create_window(0,0, window=frame_contener, anchor=NW)
        canvas.bind('<Configure>', lambda e :  self.FrameWidth(e, frame_contener, canvas, canvas_frame, width=widthFrame, height=heightFrame))
        
        frame_contener.bind("<Configure>", lambda e : self.OnFrameConfigure(canvas))

        # Scroll molette
        frame_contener.bind('<Enter>', lambda e: self.BoundToMousewheel(canvas, y =y, x=x))
        frame_contener.bind('<Leave>', lambda e: self.UnboundToMousewheel(canvas))

        return frame_contener


    def SetPopup(self, event, menu): 
        try: 
            menu.tk_popup(event.x_root, event.y_root) 
        finally: 
            menu.grab_release() 

    def SetPopupChoice(self, parent, menu, background, active_color, font="Arial 10"):
        # [["manger", chanter], "SEP", "chanter", "rire"]
        menu_popup = Menu(self, tearoff = 0, border=None, bg=background, relief=FLAT)

        for m in menu:
            if m == "SEP":
                menu_popup.add_separator() 
            elif len(m) == 2 and isinstance(m, tuple) or isinstance(m, list):
                menu_popup.add_command(label = m[0], font=font, activebackground=active_color, command=m[1], hidemargin=False) 
            elif  isinstance(m, str):
                menu_popup.add_command(label = m, font=font, activebackground=active_color, hidemargin=False) 
     
        parent.bind("<Button-3>", lambda e: self.SetPopup(e, menu_popup)) 

        return menu_popup
    

    def SetWidgetSubject(self, e, entryOut, frameX, valueEntry=None):
        if valueEntry == None:
            valueEntry = entryOut.get()

        if valueEntry != None and valueEntry != "":
 
            entryX = self.SetEntry(frameX, (0,self.number_subjects),  width= (), background=COLOR_WHITE, fontSize=8, hiLig=1, hiLigBg=COLOR_SECOND, sk=(W, E)) 
            btnX = self.SetButton(frameX, (1,self.number_subjects), image=self.PICTURES["cross"], sk=E, ipady=0, background=COLOR_WHITE, command= lambda : self.WindowClosed(entryX, btnX.destroy()))  

            entryX.insert(-1, valueEntry)

            self.number_subjects += 1
            entryOut.delete(0, END)

            self.update()
            return entryX
       
    def SetWidgetPart(self, i , subject, value):
        frame_part = self.SetFrame(self.frame_workspace, (),(5,"w"),(0,i), sk=(W,E), pady=10, background=COLOR_WHITE, sizeText=True)
        
        text = self.SetText(frame_part, 8, (1,0), padx=5, background=COLOR_LIGHT_GREY, sk=(W,E), width=40, height=8)
        if value != None and value != "":
            text.insert(END, value.strip())
        
        self.SetLabel(frame_part, subject, 10, (0,0), padx=30, background=COLOR_WHITE, sk=(W,E))

        frame_part.grid_columnconfigure(0, weight=5)


    def SetStyle(self, item, style):
        item.style = style
        for event, setting_style in style.items():            
            item.bind(event, self.ChangeStyle)
    
# =========== FONCTIONS GET =============
    def GetSubjects(self):
        self.list_subjects = []

        for child in self.frameSubjects.winfo_children():
            if isinstance(child , Entry):
                self.list_subjects.append(child.get())
        return self.list_subjects

    def GetNames(self):
        self.list_names = []

        for child in self.frame_names.winfo_children():
            if isinstance(child , Entry):
                self.list_names.append(child.get())
        if len(self.list_names) == 0:
            self.list_names.append("x")
        return self.list_names

    def GetCommentaire(self):
        list_commentaire = {}
        
        for frame in self.frame_workspace.winfo_children():
            for child in frame.winfo_children():           
                if isinstance(child , Text):
                    text= child.get("1.0",END)
                if isinstance(child , Label):
                    list_commentaire[child["text"]] = text

        return list_commentaire

# =========== FONCTIONS REMOVE =============    
    def RemoveDisplay(self):
        try:
            self.display.destroy()
            self.display = None

        except:
            pass

# ============ FONCTIONS CHANGE ============
    def ChangeMenuRight(self, menu="home"):

        self.menu_chose = menu
        try :
            self.right.destroy
        except:pass
        self.right = self.SetFrame(self.display, (10,"w"), (), (1,0), sk = (N,S,W,E), background=COLOR_WHITE, sizeText=True)
        
        self.right.grid_columnconfigure(0, weight=5)
       
        match self.menu_chose:
                
                case "new":
                    # TOP
                    self.SetLabel(self.right, "Nouveau", 13, (0,0), background=COLOR_WHITE, ipadx= 30, ipady=10, sk=W)

                    # LINES
                    self.SetLine(self.right, (0,1), (100, "w"), 1, background=COLOR_LINE)
                    self.SetLine(self.right, (0,4), (100, "w"), 1, background=COLOR_LINE)

                    # AFFICHANGE FORMAT LIVRET DEJA CREE
                   
                    self.format_report_book =  self.SetFrame(self.right, (), (), (0, 6), background=COLOR_WHITE, sk=(N,S, W,E), padx=20)
                    self.DisplayNewFile()

                    # AFFICHAGE AJOUTER UN FORMAT DE LIVRET      
                    self.SetLabel(self.right, "Créer un nouveau formet de livret", 12, (0,5), background=COLOR_WHITE, ipadx= 30, ipady=10, sk=W)


                    self.SetButton(self.format_report_book, (0,0),  image=self.PICTURES['add'], text="Nouveau Livret", placementText='top', background=COLOR_WHITE, activeBg=COLOR_WHITE,
                     ipady=5, fontSize=9, command = lambda : self.WindowFormatReportBook(False, "Livret Vierge"),  padx=15, ipadx=10)
                    self.SetButton(self.format_report_book, (1,0), image=self.PICTURES['multipleAdd'], text="Nouveau Livret", placementText='top', background=COLOR_WHITE, activeBg=COLOR_WHITE,
                     ipady=5, fontSize=9 , command= lambda : self.WindowFormatReportBook(True, "Groupe de Livret Vierge"),  padx=15, ipadx=10, )

                     #BOTTOM
                    self.SetFrame(self.right, (), (1, "w"), (0, 8), background=COLOR_WHITE, sk=(N,S,W,E))

                    # WEIGHT                    
                    self.right.grid_rowconfigure(0, weight=1)
                    self.right.grid_rowconfigure(3, weight=10)
                    self.right.grid_rowconfigure(6, weight=1)
                    self.right.grid_rowconfigure(6, weight=10)
                
                case "open":
                    self.SetLabel(self.right, "Ouvrir un fichier", 13, (0,0), background=COLOR_WHITE, ipadx= 30, ipady=10, sk=W)
                    self.SetLine(self.right, (0,1), (100, "w"), 1, background=COLOR_LINE)

                    self.frame_open_bottom= self.SetFrame(self.right, (), (), (0, 2), sk= (S,W,E,N))

                    self.frame_open_menu= self.SetFrame(self.frame_open_bottom, (15, "w"), (), (0, 0), sk= (N,S,W,E), background=COLOR_WHITE)
                    self.SetLine(self.frame_open_bottom, (1,0), (100, "w"), 1, background=COLOR_LINE, horizontal=False)
                    self.frame_last_file= self.SetFrame(self.frame_open_bottom, (), (), (2, 0), sk= (E,W,N,S), background=COLOR_WHITE, padx=30, ipady=30)

                    # MENU
                    self.btn_open_explorer= self.SetButton(self.frame_open_menu, (0,0), text=" Parcourir", sk=(E,W), placementText=LEFT, image=self.PICTURES["openFolder"], fontSize=10, ipady=20,
                     background=COLOR_WHITE, activeBg=COLOR_WHITE, command= lambda : self.BrowseFiles(self,  (("all files", "*.*"),("Excel","*.xlsx*")) , entry = None, initialdir="xxx", openFile=True))

                    self.btn_open_desktop = self.SetButton(self.frame_open_menu, (0,1), text=" Bureau   ", sk=(E,W), placementText=LEFT, image=self.PICTURES["desktop"], fontSize=10, ipady=20,
                     background=COLOR_WHITE, activeBg=COLOR_WHITE, command= lambda : self.BrowseFiles(self,  (("all files", "*.*"),("Excel","*.xlsx*")) , entry = None, initialdir=path.normpath(path.expanduser("~/Desktop")), openFile=True))

                    # self.btn_open_computer = self.SetButton(self.frame_open_menu, (0,2),text=" Ce PC    ", sk=(E,W), placementText=LEFT, image=self.PICTURES["laptop"], fontSize=10, ipady=20,
                    #  background=COLOR_WHITE, activeBg=COLOR_WHITE, command= lambda : self.BrowseFiles(self,  (("all files", "*.*"),("Excel","*.xlsx*")) , entry = None, 
                    #  initialdir="Music", openFile=True))

                #   MyComputerFolder
                    self.SetStyle(self.btn_open_explorer, self.BTN_STYLE_WHITE_COLOR)
                    self.SetStyle(self.btn_open_desktop, self.BTN_STYLE_WHITE_COLOR)
                    # self.SetStyle(self.btn_open_computer, self.BTN_STYLE_WHITE_COLOR)

                    # LAST FILE
                    self.DisplayLastFile()
       
                    # WEIGHT                    
                    self.right.grid_rowconfigure(0, weight=1)
                    self.right.grid_rowconfigure(2, weight=20)

                    self.frame_open_bottom.grid_rowconfigure(0, weight=10)
                    self.frame_open_bottom.grid_columnconfigure(2, weight=10)

                    self.frame_open_menu.grid_columnconfigure(0, weight=10)
                    
                case "home" | _:
                    # ----- Top -----
                    self.SetLabel(self.right, "Bonjour", 12, (0,0), background=COLOR_WHITE, ipadx= 30, ipady=20, sk=W)

                    self.SetLine(self.right, (0,1), (100, "w"), 1, background=COLOR_LINE)
                    self.SetLine(self.right, (0,4), (100, "w"), 1, background=COLOR_LINE)

                    # ----- New -----
                    self.SetLabel(self.right, "Nouveau", 10, (0,2), background=COLOR_WHITE, ipadx= 30, ipady=20, sk=W)
                    self.format_report_book = self.SetFrame(self.right, (), (10, "w"), (0,3), background=COLOR_WHITE,sk=(W, E,N,S), sizeText=True)

                    self.SetLabel(self.right, "Nouveau", 10, (0,2), background=COLOR_WHITE, ipadx= 30, ipady=20, sk=W)
                    self.frame_last_file = self.SetFrame(self.right, (), (20, "w"), (0,5), sk = (W,E, N,S), background=COLOR_WHITE, padx=30, ipady=30 )
                    self.right.grid_rowconfigure(3, weight=10)
                    self.right.grid_rowconfigure(5, weight=5)
                    
                    self.format_report_book.grid_columnconfigure(0, weight=5)
                    # self.frame_last_file= self.SetFrame(self.right, (), (), (é0, 0), sk= (E,W,N,S), background=COLOR_WHITE, padx=30, ipady=30)

                    #  ----------- Last File -----------------

                    self.DisplayLastFile()
                    self.DisplayNewFile()
    
    def ChangeStyle(self, event):   

            for key, value in event.widget.style[f"<{self.RecupEvent(event.type)}>"].items(): # liste le dico grâce à une boucle for        
                event.widget[key] = value
   
# ============ FONCTIONS DISPLAY ============
    def DisplayLastFile(self):

        self.SetLabel(self.frame_last_file, "Dernier Fichier", 12, (0,0), background=COLOR_WHITE, ipadx= 0, ipady=20, sk=W)

        self.frame_LF_indication = self.SetFrame(self.frame_last_file, (), (), (0,1), sk=(W,E), background=COLOR_WHITE, sizeText=True, padx=0)
        self.SetLabel(self.frame_LF_indication, " Nom", 8, (0,0), background=COLOR_WHITE, ipadx= 10, ipady=5, sk=W, color=COLOR_GREY, placementText=LEFT, image=self.PICTURES["documentB"])
        self.SetLabel(self.frame_LF_indication, " Date de modification", 8, (0,0), background=COLOR_WHITE, ipadx= 20, ipady=5, sk=E, color=COLOR_GREY, placementText=LEFT, image=self.PICTURES["time"])
        self.SetLine(self.frame_LF_indication, (0,2), (100, "w"), 1, background=COLOR_LINE)
        
        self.frame_LF_contener = self.SetFrameScroll(self.frame_last_file, (0,2), self.frame_last_file, (3,2), y=True, background=COLOR_WHITE)

        i = 1
        # Button last file
        for name, directory, date in reversed(GetListLastFile()):
            self.CreateBtnLastFile(name, directory, date, i)
            i += 2

        # WEIGHT                  
     
        self.frame_last_file.grid_columnconfigure(0, weight=10)
        self.frame_last_file.grid_rowconfigure(2, weight=10)

        self.frame_LF_indication.grid_columnconfigure(0, weight=10)

        self.frame_LF_contener.grid_columnconfigure(0, weight=10)
        self.frame_LF_contener.grid_rowconfigure(0, weight=10)

    def DisplayNewFile(self):
        # AFFICHANGE FORMAT LIVRET DEJA CREE
        self.SetLabel(self.right, "Format livret déjà créé", 11, (0,2), background=COLOR_WHITE, ipadx= 30, ipady=10, sk=W)
        self.frame_existing_FRB = self.SetFrame(self.right, (), (), (0, 3), background=COLOR_WHITE, sk=(N,S,W,E))
        
        self.frame_existing_FRB_canvas = self.SetFrameScroll(self.frame_existing_FRB, (0,0), self.frame_last_file, (10,10), x=True, background=COLOR_WHITE, heightFrame=True, widthFrame=True, scrollBarUsed=False)

        i = 0
        self.pictures_report_book =  {}                   

        for key, format in reversed(getFormats().items()):
            self.CreateBtnExisitingFRB(i, format) # fait dans une fonction pour que les émlément puissent se reconaitre                         
            i += 1

        self.update()

        # WEIGHT
        self.frame_existing_FRB.grid_rowconfigure(0, weight=5)
        self.frame_existing_FRB.grid_columnconfigure(0, weight=5)
        self.frame_existing_FRB_canvas.grid_rowconfigure(0, weight=10)
        self.format_report_book.rowconfigure(0, weight=10)
    
# ============ FONCTIONS CREATE ============
    def CreateBtnLastFile(self,name, direcotry, date, i):
        
        frame_LF_button = self.SetFrame(self.frame_LF_contener, (), (5, "w"), (0,i -1), sk=(W,E), background=COLOR_WHITE, pady=10,padx=10, sizeText=True, cursor="hand2")      
        
        frame_LF_btn_left = self.SetFrame(frame_LF_button, (), (5, "w"), (0,0), sk=(W,E), background=COLOR_WHITE, pady=5, sizeText=True)
            
        frame_LF_button.grid_columnconfigure(0, weight=10)
        frame_LF_button.grid_columnconfigure(1, weight=10)
        frame_name = self.SetLabel(frame_LF_btn_left, name, 10, (0,0), background=COLOR_WHITE, ipadx= 10, ipady=0, sk=W)
        frame_directory = self.SetLabel(frame_LF_btn_left, direcotry, 8, (0,1), background=COLOR_WHITE, ipadx= 10, ipady=0, sk=W, color=COLOR_GREY)

        frame_date = self.SetLabel(frame_LF_button, date, 8, (1,0), background=COLOR_WHITE, padx= 20, ipady=0, sk=E, color=COLOR_GREY)
        self.SetLine(self.frame_LF_contener, (0,i), (), 1, background=COLOR_LINE, sk=(W,E))
        
        frame_directory.bind("<Button-1>", lambda e : self.WindowApplication(direcotry + "/" + name))
        frame_date.bind("<Button-1>",  lambda e: self.WindowApplication(direcotry + "/" + name))
        frame_name.bind("<Button-1>", lambda e: self.WindowApplication(direcotry + "/" + name))
        frame_LF_button.bind("<Button-1>",  lambda e: self.WindowApplication(direcotry + "/" + name))
        frame_LF_btn_left.bind("<Button-1>",  lambda e: self.WindowApplication(direcotry + "/" + name))

    def CreateBtnExisitingFRB(self, i, format):
                                           
        try: 
            self.pictures_report_book[i]  = self.LaunchPicture(format["picture"], proportion=12)  
        except:    
            self.pictures_report_book[i]  = self.PICTURES["reportBook"]

        btn = self.SetButton(self.frame_existing_FRB_canvas, (i,0) ,image=self.pictures_report_book[i], text=format["title"], placementText='top', background=COLOR_WHITE, activeBg=COLOR_WHITE,
            fontSize=9, padx=15, ipadx=10, ipady=5, command=lambda : self.WindowCreteNewFile(format))


        self.SetPopupChoice(btn, [("Créer", lambda : self.WindowCreteNewFile(format)), "SEP", ("Modifier", lambda : self.CommandModifyFormat(format)),
                        ("Supprimer", lambda : self.CommandRemoveFormat(format["title"]))], COLOR_WHITE, COLOR_SECOND)

# ========== FONCTIONS COMMANDES ==============       
    def CommandRemoveFormat(self, title):
        RemoveFormat(title)
        self.ChangeMenuRight(menu="new")

    def CommandModifyFormat(self, format):
        self.WindowFormatReportBook(title= format["title"], button_result="Modifier", remove_format=format["title"], groupReportBook=format["group"], text_name=format["title"], text_picture=format["picture"], text_subjects=format["subjects"])

    def CommandRemoveCreate(self, old_name, name):   
        if old_name != name and old_name != "" and old_name != None:
            RemoveFormat(old_name)
        self.ChangeMenuRight(menu="new")

# ======== FONCTIONS COMPLEMENTAIRES ========
    def FrameWidth(self, event, frame_contener, canvas, canvas_frame, width=True, height=False):
        if event.width != frame_contener.winfo_width():
            if width == True and event.width > frame_contener.winfo_width():
                canvas.itemconfig(canvas_frame, width=event.width + 1)
            elif width == True and event.width < frame_contener.winfo_width():
                canvas.itemconfig(canvas_frame, width=event.width - 1)


            if height == True and event.height > frame_contener.winfo_height():
                    canvas.itemconfig(canvas_frame, height=event.height+4)
            elif width == True and event.width < frame_contener.winfo_width():
                canvas.itemconfig(canvas_frame, width=event.width - 4)
      
    def OnFrameConfigure(self, canvas):
        canvas.configure(scrollregion = canvas.bbox("all"))
   
    def BrowseFiles(self, parentFocus, fileGood, entry=None, initialdir = "/", openFile=False, open=True): 

        if open == True:
           
            self.file_name = filedialog.askopenfilename(initialdir = initialdir, 
                                                title = "Ouvrir un fichier",
                                                filetypes = fileGood) 
        else:
            self.file_name = filedialog.asksaveasfilename(initialdir = initialdir, 
                                                title = "Enregistrer un fichier", 
                                                filetypes = fileGood)
            if self.file_name.find("xlsx") == -1 and self.file_name != "" and self.file_name != None:
                self.file_name += ".xlsx"

        parentFocus.focus()
        if self.file_name != "" and self.file_name != None:

            if entry != None:
                entry.delete(0, END)
                entry.insert(-1, self.file_name)

            if openFile == True:
                self.WindowApplication(self.file_name)     

        return self.file_name

    def BoundToMousewheel(self, canvas, x=False, y=False):
        canvas.bind_all("<MouseWheel>", lambda e: self.OnMousewheel(e, canvas, x , y))   

    def UnboundToMousewheel(self, canvas):
        canvas.unbind_all("<MouseWheel>") 

    def OnMousewheel(self, event, canvas, x , y):
        if x:
            canvas.xview_scroll(int(-1*(event.delta/120)), "units")
        if y:
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def ProportionScreen(self, percentage, width=True):
        if width == True:
            return round ((percentage / 100 )* self.winfo_screenwidth())
        else:
            return round ((percentage / 100 )* self.winfo_screenheight())

    def RecupEvent(self, number):
        dicoEvent = {'2': 'KeyPress', 'KeyPress': 'Key', '3': 'KeyRelease', '4': 'ButtonPress', 'ButtonPress': 'Button', '5': 'ButtonRelease', '6': 'Motion', '7': 'Enter', 
        '8': 'Leave', '9': 'FocusIn', '10': 'FocusOut', '11': 'Keymap', '12': 'Expose', '13': 'GraphicsExpose', '14': 'NoExpose', '15': 'Visibility', '16': 'Create', '17': 'Destroy', 
        '18': 'Unmap', '19': 'Map', '20': 'MapRequest', '21': 'Reparent', '22': 'Configure', '23': 'ConfigureRequest', '24': 'Gravity', '25': 'ResizeRequest', 
        '26': 'Circulate', '27': 'CirculateRequest', '28': 'Property', '29': 'SelectionClear', '30': 'SelectionRequest', '31': 'Selection', '32': 'Colormap', 
        '33': 'ClientMessage', '34': 'Mapping', '35': 'VirtualEvent', '36': 'Activate', '37': 'Deactivate', '38': 'MouseWheel'}


        if dicoEvent.get(str(number)) != None:
            return dicoEvent.get(str(number))
    
    def ProportionParent(self, pourcentage, parent):
        if pourcentage == None:
            return None
        else:
            return round(parent * (pourcentage/100))

    def SizeFont(self, size):
        return round(self.winfo_screenwidth() / 1000 * size)
    
    def CheckSettingSize(self, number, parentSize, width=True):
        n = 0

        if False == (type(number) is list or type(number) is tuple):
            number = [number , "o"]

        if len(number) != 2:
            number = [0, "o"]

        match number[1]:
            case "p": # parent
                n = self.ProportionParent(number[0], parentSize)
            case "w": # window
                if width == True:
                    n = self.ProportionScreen(number[0])
                else:
                    n = self.ProportionScreen(number[0], width=False)

            case "o": # window
                n = number[0]
            case _:
                n = 0
        return n

    def LaunchPicture(self, picture, resize = True, proportion = 10):
        if resize == True:
            image = Image.open(picture)
            w, h = image.size
            if w > h:
                newWidth = ((w / w)) * proportion * (self.winfo_screenwidth() / 100)
                newHeight = ((h / w)) *  proportion * (self.winfo_screenwidth() / 100)
            else:
                newWidth = ((w / h)) * proportion * (self.winfo_screenwidth() / 100)
                newHeight = ((h / h)) * proportion * (self.winfo_screenwidth() / 100)

            img = Image.open(picture).resize((round(newWidth), round(newHeight)))

        else:
            img = Image.open(picture)
        return ImageTk.PhotoImage(img)
    
    def MoveList(self, n):
        self.position_list += n

    def CopySheet(self):
        try :             
            list_values= list(self.GetCommentaire().values())

            # Evite une erreur si on a écrit __SPLIT__ dans le fichier
            list_values = "ZX1O0ER".join(list_values)
            list_values = list_values.replace("_SPLIT_", "",100)
            list_values = list_values.replace("ZX1O0ER", "_SPLIT_", 100)

            self.clipboard_clear()
            self.clipboard_append(list_values)
        except : pass

# ========== FONCTION WINDOW + APPLICATION ===========
    def WindowClosed(self, window, *args):
        self.active = False
        window.destroy()
        self.update()

    def WindowCreteNewFile(self, format):
        if self.active == False:

            # PARAMETRE FENETRE
            self.window_create_file = Toplevel()
            self.window_create_file.configure(background=COLOR_WHITE)
            self.window_create_file.title("Paramètres de création")

            # S'active si la fenetre est fermé
            self.window_create_file.protocol("WM_DELETE_WINDOW", lambda :  self.WindowClosed(self.window_create_file))

            self.frame_create_file_top = self.SetFrame(self.window_create_file, (), (), (0, 0), background=COLOR_WHITE, sk=(W,E), sizeText=True, ipady=10, ipadx=5)
            
            self.frame_create_file = self.SetFrame(self.frame_create_file_top, (), (), (0, 0), background=COLOR_WHITE, sk=(W,E,N,S), sizeText=False, ipady=10, ipadx=5)

            # top
            self.SetLabel(self.frame_create_file, format["title"], 13, (0,0), sk=W, background=COLOR_WHITE)

            # Spinbox Trimestre 
            self.frame_spinbox_trimester = self.SetFrame(self.frame_create_file, (), (), (0, 1), background=COLOR_WHITE, sk=(W,E,N), sizeText=True, ipady=15, ipadx=5)
            self.SetLabel(self.frame_spinbox_trimester, "Semestre ", 11, (0,0), sk=E, background=COLOR_WHITE, ipady=0)    
            self.spinbox_trimester = self.SetSpinBox(self.frame_spinbox_trimester, (1,0), size_number=(0,10), background=COLOR_LIGHT_GREY)  

            # Entry classe     
            self.frame_entry_class = self.SetFrame(self.frame_create_file, (), (), (0, 2), background=COLOR_WHITE, sk=(W,E,N), sizeText=True, ipady=15, ipadx=5)
            self.SetLabel(self.frame_entry_class, "Niveau Classe", 11, (0,0), sk=E, background=COLOR_WHITE, ipady=0)    
            self.entry_level = self.SetEntry(self.frame_entry_class, (1,0), (8 , "o"), sk=(N), background=COLOR_LIGHT_GREY, padx=5, fontSize=8)

            # Entry année 
            self.frame_entry_year = self.SetFrame(self.frame_create_file, (), (), (0, 3), background=COLOR_WHITE, sk=(W,E,N), sizeText=True,  ipady=15, ipadx=5)
            self.SetLabel(self.frame_entry_year, "Année", 11, (0,0), sk=E, background=COLOR_WHITE, ipady=0)    
            self.entry_year1 = self.SetEntry(self.frame_entry_year, (1,0), (10, "o"), background=COLOR_LIGHT_GREY, padx=5, fontSize=8, sk=(W,E))
            self.SetLabel(self.frame_entry_year, "-", 11, (2,0), sk=E, background=COLOR_WHITE, ipady=0)    
            self.entry_year2 = self.SetEntry(self.frame_entry_year, (3,0), (10,"o"), background=COLOR_LIGHT_GREY, padx=5, fontSize=8, sk=(W,E))

            # Entry prenom élève
            self.frame_entry_name = self.SetFrame(self.window_create_file, (), (), (0, 4), background=COLOR_WHITE, sk=(W,E,N,S),  ipady=15, ipadx=5)

            if format["group"]:
                self.window_create_file.resizable(False, True)

                self.window_create_file.geometry(f"{round(self.winfo_screenwidth() * 0.4)}x{round(self.winfo_screenheight()* 0.6)}")

                self.SetLabel(self.frame_create_file_top, "", 13, (1,0), image=self.PICTURES["people"], background=COLOR_WHITE, sk=(N,W, E), ipady=30) # image

                self.entry_name = self.SetEntry(self.frame_entry_name, (2,0), (10,"o"), sk=(W,E), background=COLOR_LIGHT_GREY, padx=5, fontSize=8)
                self.SetLabel(self.frame_entry_name, "Prénoms", 11, (0,0), sk=N, background=COLOR_WHITE, ipady=0) 
                self.btn_add_name = self.SetButton(self.frame_entry_name, (1,0), image=self.PICTURES["addBtn"], padx=5,  ipady=0, background=COLOR_SECOND, activeBg=COLOR_SECOND_CLICK,
                color=COLOR_WHITE,activeFg=COLOR_WHITE, command= lambda : self.SetWidgetSubject("x", self.entry_name,  self.frame_names))   
                self.entry_name.bind("<Return>", lambda e: self.SetWidgetSubject("x", self.entry_name,  self.frame_names))
                
                self.frame_names_bottom = self.SetFrame(self.frame_entry_name, (), (), (2, 1), background=COLOR_LIGHT_GREY,sk=(N,S,W,E))
                self.frame_entry_name.grid_columnconfigure(2, weight= 5)
            
                self.frame_names = self.SetFrameScroll(self.frame_names_bottom, (0,0), self.frame_names_bottom, (1,0), y=True, background=COLOR_WHITE, widthFrame=True, skScrollBar=(N,S))
                self.SetStyle(self.btn_add_name, self.BTN_STYLE_SECOND_COLOR)


                # WEIGHT
                self.frame_entry_name.grid_rowconfigure(1, weight= 5)
                self.frame_entry_name.grid_columnconfigure(2, weight= 5)
                
                self.frame_names_bottom.grid_rowconfigure(0, weight= 5)
                self.frame_names_bottom.grid_columnconfigure(0, weight= 5)
           
                self.frame_names.grid_columnconfigure(0, weight=10)

            else:
                self.window_create_file.resizable(False, False)

                self.window_create_file.geometry(f"{round(self.winfo_screenwidth() * 0.4)}x{round(self.winfo_screenheight()* 0.4)}")

                self.SetLabel(self.frame_entry_name, "Prénom", 11, (0,0), sk=N, background=COLOR_WHITE, ipady=0)    
                self.frame_names = self.SetFrame(self.frame_entry_name, (), (), (1,0), background=COLOR_WHITE, sk = (W,E, N,S))
                self.entry_name = self.SetEntry(self.frame_names, (0,0), (), sk=(W,E), background=COLOR_LIGHT_GREY, padx=5, fontSize=8)
                
                self.SetLabel(self.frame_create_file_top, "", 13, (1,0), image=self.PICTURES["person"], background=COLOR_WHITE, sk=(N,W, E), ipady=30) # image
                self.frame_entry_name.grid_columnconfigure(1, weight= 5)
                self.frame_entry_name.grid_columnconfigure(1, weight= 5)
                self.frame_names.grid_columnconfigure(0, weight= 5)

            self.frameBottom = self.SetFrame(self.window_create_file, (), (), (0, 10), background=COLOR_WHITE,sk=(E,S), sizeText=True)

            self.btn_cancel= self.SetButton(self.frameBottom, (0,0), text = "Annuler", fontSize=11, padx=10 , pady=5, command = lambda :  self.WindowClosed(self.window_create_file),
                    background=COLOR_SECOND, activeBg=COLOR_SECOND_CLICK, color=COLOR_WHITE,activeFg=COLOR_WHITE)
            
            self.btn_create_format = self.SetButton(self.frameBottom, (1,0), text = "Créer", fontSize=11,  padx=5, pady=5, background=COLOR_SECOND, activeBg=COLOR_SECOND_CLICK,
                color=COLOR_WHITE,activeFg=COLOR_WHITE, command= lambda : self.WindowClosed(self.window_create_file, CreateFile(self.spinbox_trimester.get(), self.entry_level.get(), f"{self.entry_year1.get()} {self.entry_year2.get()}", format["subjects"], self.GetNames()), self.WindowApplication("")))
            
            self.SetStyle(self.btn_create_format, self.BTN_STYLE_SECOND_COLOR)
            self.SetStyle(self.btn_cancel, self.BTN_STYLE_SECOND_COLOR)

            

            # WEIGHT                
            self.window_create_file.grid_columnconfigure(0, weight= 10)
            self.window_create_file.grid_rowconfigure(4, weight= 10)

            self.frame_create_file_top.grid_columnconfigure(0, weight= 10)
            self.frame_create_file_top.grid_columnconfigure(1, weight= 5)
            self.frame_create_file_top.grid_rowconfigure(0, weight= 10)
            
            self.frame_entry_year.grid_columnconfigure(1, weight= 5)
            self.frame_entry_year.grid_columnconfigure(3, weight= 5)
            
            self.frame_create_file.grid_columnconfigure(0, weight= 5)
            
            self.frame_create_file.grid_rowconfigure(0, weight= 1)
            self.frame_create_file.grid_rowconfigure(1, weight= 1)
            self.frame_create_file.grid_rowconfigure(2, weight= 1)
            self.frame_create_file.grid_rowconfigure(3, weight= 1)
            
            self.active = True

    def WindowFormatReportBook(self, groupReportBook, title, button_result="Créer", remove_format=None, text_name="", text_picture="", text_subjects=None):
        if self.active == False:

            # PARAMETRE FENETRE
            self.window_add_format = Toplevel()
            self.window_add_format.geometry(f"{round(self.winfo_screenwidth() * 0.4)}x{round(self.winfo_screenheight()* 0.5)}")
            self.window_add_format.configure(background=COLOR_WHITE)
            self.window_add_format.title("Paramètres du nouveau format")

            # S'active si la fenetre est fermé
            self.window_add_format.protocol("WM_DELETE_WINDOW", lambda :  self.WindowClosed(self.window_add_format))
            
            # Top
            self.SetLabel(self.window_add_format, title, 13, (0,0), sk=W, background=COLOR_WHITE)

            # Entry Name
            self.frameNameFormat = self.SetFrame(self.window_add_format, (), (), (0, 1), background=COLOR_WHITE, sk=(W,E,S,N), sizeText=True, ipady=10, ipadx=5)
            self.SetLabel(self.frameNameFormat, "Nom format", 11, (0,0), sk=(E,N,S), background=COLOR_WHITE, ipady=0)
            self.entryNameFormat = self.SetEntry(self.frameNameFormat, (1,0), (), sk=(W,E), background=COLOR_LIGHT_GREY, padx=5, fontSize=8)

            # Entry decoration picture
            self.frameImageFormat = self.SetFrame(self.window_add_format, (), (), (0, 2), background=COLOR_WHITE,sk=(N,S,W,E), sizeText=True)

            self.SetLabel(self.frameImageFormat, "Image décoration", 11, (0,0), sk=E, background=COLOR_WHITE, ipady=0)                
            self.btnSetPicture = self.SetButton(self.frameImageFormat, (2,0), text = "Rechercher", command = lambda : self.BrowseFiles(self.window_add_format, (("all files", "*.*"),("PNG","*.png*"), ("JPEG","*.jpeg*,.jpg*.jpe*")), entry=self.entryPicture), fontSize=9, padx=5,
                ipady=0, background=COLOR_SECOND, activeBg=COLOR_SECOND_CLICK,  color=COLOR_WHITE,activeFg=COLOR_WHITE)
            self.entryPicture = self.SetEntry(self.frameImageFormat, (1,0), (), sk=(W,E), background=COLOR_LIGHT_GREY, padx=5, pady=5, fontSize=8)
            self.SetStyle(self.btnSetPicture, self.BTN_STYLE_SECOND_COLOR)

            # set subject
            self.frameSetSubjects = self.SetFrame(self.window_add_format, (), (), (0, 3), background=COLOR_WHITE,sk=(N,S,W,E))

            self.SetLabel(self.frameSetSubjects, "Matières", 11, (0,0), sk=(E,N), background=COLOR_WHITE, ipady=0)                
            self.btnSetSubject = self.SetButton(self.frameSetSubjects, (1,0), image=self.PICTURES["addBtn"], padx=5,  ipady=0, background=COLOR_SECOND, activeBg=COLOR_SECOND_CLICK,
                color=COLOR_WHITE,activeFg=COLOR_WHITE, command= lambda : self.SetWidgetSubject("x", self.entrySubjects, self.frameSubjects))
            self.SetStyle(self.btnSetSubject, self.BTN_STYLE_SECOND_COLOR)

            self.entrySubjects = self.SetEntry(self.frameSetSubjects, (2,0), (), sk=(W,E), background=COLOR_LIGHT_GREY, padx=5, pady=5, fontSize=8)   
            self.entrySubjects.bind("<Return>", lambda e: self.SetWidgetSubject(e, self.entrySubjects, self.frameSubjects))

            self.frameSubjectsBottom = self.SetFrame(self.frameSetSubjects, (), (), (2, 1), background=COLOR_LIGHT_GREY,sk=(N,S,W,E))
            
            self.frameSubjects = self.SetFrameScroll(self.frameSubjectsBottom, (0,0), self.frameSubjectsBottom, (1,0), y=True, background=COLOR_WHITE, widthFrame=True, skScrollBar=(N,S))

            # Bottom
            self.frameBottom = self.SetFrame(self.window_add_format, (), (), (0, 10), background=COLOR_WHITE,sk=(E,S), sizeText=True)
            self.btn_cancel= self.SetButton(self.frameBottom, (0,0), text = "Annuler", fontSize=11, padx=10 , pady=5, command = lambda :  self.WindowClosed(self.window_add_format),
                background=COLOR_SECOND, activeBg=COLOR_SECOND_CLICK, color=COLOR_WHITE,activeFg=COLOR_WHITE)
            self.btn_create_format = self.SetButton(self.frameBottom, (1,0), text = button_result, fontSize=11,  padx=5, pady=5, background=COLOR_SECOND, activeBg=COLOR_SECOND_CLICK,
                color=COLOR_WHITE,activeFg=COLOR_WHITE, command= lambda : self.WindowClosed(self.window_add_format, setFormat(self.entryNameFormat.get(), self.GetSubjects(), groupReportBook, self.entryPicture.get()), self.CommandRemoveCreate(text_name, self.entryNameFormat.get())))

            self.SetStyle(self.btn_cancel, self.BTN_STYLE_SECOND_COLOR)
            self.SetStyle(self.btn_create_format, self.BTN_STYLE_SECOND_COLOR)

            # INSERT
            if text_subjects != None:
                for subject in text_subjects:
                    self.SetWidgetSubject("x", self.entrySubjects,  self.frameSubjects, valueEntry=subject)
         
            self.entryNameFormat.insert(-1,text_name)
            self.entryPicture.insert(-1, text_picture)

            # WEIGHT                
            self.window_add_format.grid_columnconfigure(0, weight= 5)
            
            self.window_add_format.grid_rowconfigure(1, weight= 1)
            self.window_add_format.grid_rowconfigure(2,weight= 1)
            self.window_add_format.grid_rowconfigure(3,weight= 20)
            self.window_add_format.grid_rowconfigure(10,weight= 1)

            self.frameNameFormat .grid_columnconfigure(1, weight=5)
            self.frameImageFormat .grid_columnconfigure(1, weight=5)

            self.frameSetSubjects .grid_columnconfigure(2, weight=5)
            self.frameSetSubjects.grid_rowconfigure(1 , weight=5)

            self.frameSubjectsBottom.grid_columnconfigure(0 , weight=5)                
            self.frameSubjectsBottom.grid_rowconfigure(0 , weight=5)

            self.frameSubjects.grid_columnconfigure(0, weight=5)
            
            self.active = True
    
    def WindowSave(self):
        if self.active == False:
            self.active = True
            self.BrowseFiles(self, (("all files", "*.*"),("EXCEL","*.xlsx*")), open=False, initialdir="xxx")

            self.active = False

        if self.file_name != None and self.file_name != "" and "back.xlsx" != self.file_name:
            SaveFile (file_name= self.file_name, file_base=self.file_used)
            self.file_used = self.file_name
            
            f_name, f_way = SplitFile(self.file_name)
            SetLastFile(f_name, f_way)
  
    def UpdateWindow(self, n=None, *args):
        self.CopySheet()
        self.RemoveDisplay()
        self.SetDisplay()

        if n != None:
            self.mode_display = n 

        match self.mode_display:
            case 1:  # Menue de choix
                self.WindowMenu()
            case 2:  # Menue de choix
                self.WindowApplication(self.file_name)
            case 3:  # Menue de choix
                self.WindowMenu(True)
            case _:
                print("rien")

    def WindowMenu(self, veille=False):
        
    # ---------- LEFT -----------
        # Frame
        self.left = self.SetFrame(self.display, (10, "w"), (), (0,0), sk = (N,S), background=COLOR_SECOND)
        self.action= self.SetFrame(self.left, (10, "w"), (50, "w"), (0,1), sk = (N, W, E), background=COLOR_SECOND,  sizeText=True)

        # Label
        self.SetLabel(self.left, "Livret", 12, (0,0),color=COLOR_WHITE, sk=N, background= COLOR_SECOND, ipady=20)

        if veille == True:
            self.btn_menu= self.SetButton(self.action, (0,0), sk=(W, E), width=(9.5 ,"w") , image=self.PICTURES['arrow'],  placementText='top', background=COLOR_SECOND, activeBg=COLOR_SECOND_CLICK,
            color=COLOR_WHITE, ipady=30,  command= lambda : self.UpdateWindow(2), activeFg=COLOR_WHITE)
            
            self.SetStyle(self.btn_menu, self.BTN_STYLE_SECOND_COLOR)

        # Button
        self.btn_home= self.SetButton(self.action, (0,1), sk=(W, E), width=(9.5 ,"w") , image=self.PICTURES['home'], text="Accueil", placementText='top', background=COLOR_SECOND, activeBg=COLOR_SECOND_CLICK,
         color=COLOR_WHITE, ipady=30,  command= lambda : self.ChangeMenuRight("home"), activeFg=COLOR_WHITE)

        self.btn_new = self.SetButton(self.action, (0,2), sk=(W, E), image=self.PICTURES['document'], text="Nouveau", placementText='top', background=COLOR_SECOND, activeBg=COLOR_SECOND_CLICK,
         color=COLOR_WHITE, ipady=30, command= lambda : self.ChangeMenuRight("new"), activeFg = COLOR_WHITE)
        self.btn_open_file = self.SetButton(self.action, (0,3), sk=(W, E), image=self.PICTURES['folder'], text="Ouvrir",placementText='top', background=COLOR_SECOND, activeBg=COLOR_SECOND_CLICK, 
        color=COLOR_WHITE, ipady=30,  command= lambda : self.ChangeMenuRight("open") ,activeFg=COLOR_WHITE)
        
        self.exit = self.SetButton(self.left, (0,6), sk=(W, E), text="Quitter",command=self.destroy, image=self.PICTURES["exit"], sizeText=11, placementText='left', background=COLOR_SECOND, activeBg=COLOR_SECOND,
         color=COLOR_WHITE, ipady=30, activeFg=COLOR_WHITE)
        
        # Style Button        
        self.SetStyle(self.btn_home, self.BTN_STYLE_SECOND_COLOR)

        self.SetStyle(self.btn_new, self.BTN_STYLE_SECOND_COLOR)

        self.SetStyle(self.btn_open_file, self.BTN_STYLE_SECOND_COLOR)

    # ---------- RIGHT -------------
    
        self.ChangeMenuRight(menu="home")

        # WEIGHT
        self.display.grid_columnconfigure(1, weight=5)
        self.display.grid_rowconfigure(0, weight=5)
        
        self.left.grid_rowconfigure(1, weight=5)

        self.action.grid_columnconfigure(0, weight=5)
        self.action.grid_rowconfigure(0, weight=2)
        self.action.grid_rowconfigure(1, weight=2)
        self.action.grid_rowconfigure(2, weight=2)


    def UpdateApplication(self, add=None):

        dico_commentaire = self.GetCommentaire()

        for subject_v, value_v in dico_commentaire.items():
            self.all_sheet[self.position_list][subject_v] = value_v.strip()

        SetSubjects(self.file, self.all_sheet[self.position_list]["name"], dico_commentaire, self.file_used)

        if add != None:
            if self.position_list + add < 0 :
                self.position_list = len(self.all_sheet) - 1
            elif self.position_list + add > len(self.all_sheet) - 1:
                self.position_list = 0
            elif add < 0 and self.position_list != 0:
                self.position_list += add
            elif add > 0 and self.position_list + add  < len(self.all_sheet) :
                self.position_list += add

        if self.file_used == "back.xlsx":
            self.WindowSave()

        self.number_students["text"] = f"{self.position_list + 1}/{len(self.all_sheet)}"

        if add != 0:
            # suprrimer tout de base
            frame_list = []
            for frame in self.frame_workspace.children.items():
                frame_list.append(frame[1])
            for fr in frame_list:
                fr.destroy()

            # Ecriture
            i=0
            for subject, value in self.all_sheet[self.position_list].items():
                if subject != "name":
                    self.SetWidgetPart(i, subject, value)
                else:
                    self.name_student.configure(text=value)
                i += 1 

    def WindowApplication(self, file_name):

        if file_name == "":
            self.file = OpenFileExcelFormat("back.xlsx")

            self.file_used = "back.xlsx"
            self.file_name = "back.xlsx"
       
        elif file_name != self.file_used:
            self.position_list = 0 # permet de revenir a la page 1 lorsque c'est un nouveau fichier
       
            self.file = OpenFileExcelFormat(file_name)
            OpenFileExcelFormat(file_name)

            self.file_name = file_name
            self.file_used = file_name


        if self.file != None and self.file_name != None and self.file_name != "":
            f_name, f_way = SplitFile(self.file_name)
            SetLastFile(f_name, f_way)

            self.RemoveDisplay()
            self.SetDisplay()
            

            self.all_sheet = GetAllTableauSheet(self.file)

            self.frame_menu_top = self.SetFrame(self.display, (),(8, "w"),(0,0), sk=(N,W,E), background=COLOR_WHITE)
            self.SetLine(self.display, (0,1), (), 1, background=COLOR_LINE, sk=(W,E))
            self.frame_bottom = self.SetFrame(self.display, (),(),(0,2), sk=(N,W,E,S), background=COLOR_LIGHT_GREY)

            # MENU TOP
            self.SetButton(self.frame_menu_top, (0,0), image=self.PICTURES["menu"], ipadx=10, ipady=10, sk=(S,N), command= lambda : self.UpdateWindow(3, self.UpdateApplication(add=0)))

            self.frame_menu_btn = self.SetFrame(self.frame_menu_top, (),(50, "w"),(1,0), sk=(N,S,W,E), background=COLOR_WHITE)

            
            self.SetButton(self.frame_menu_btn, (1,0), text=" Enregistrer", image=self.PICTURES["save"], placementText=LEFT, ipadx=20, ipady=10, sk=(S,N), command= lambda : self.UpdateApplication(add=0))
            self.SetButton(self.frame_menu_btn, (2,0), text=" Enregistrer Sous", image=self.PICTURES["save"], placementText=LEFT, ipadx=20, ipady=10, sk=(S,N),command = self.WindowSave)
            self.SetButton(self.frame_menu_btn, (3,0), text=" Copier", ipadx=20, ipady=10, sk=(W,S,N), command= self.CopySheet)

            self.number_students = self.SetLabel(self.frame_menu_top, "1/20", 12, (2,0), padx=30)
            self.name_student = self.SetLabel(self.frame_menu_top, "Nom", 12, (3,0), padx=30)
            
            # WORKSPACE
            self.frame_center = self.SetFrame(self.frame_bottom,(50,"w"),(), (2,0),sk = (N,S), background=COLOR_WHITE)
            self.SetButton(self.frame_bottom, (1,0), image=self.PICTURES["left"], ipadx=10, ipady=10, sk=(S,N), background=COLOR_LIGHT_GREY, activeBg= COLOR_LIGHT_GREY, command= lambda : self.UpdateApplication(-1))
            self.SetButton(self.frame_bottom, (3,0), image=self.PICTURES["right"], ipadx=10, ipady=10, sk=(S,N), background=COLOR_LIGHT_GREY, activeBg=COLOR_LIGHT_GREY, command= lambda : self.UpdateApplication(1))

            self.frame_workspace = self.SetFrameScroll(self.frame_center, (0,0), self.frame_center, (1,0), y=True, background=COLOR_WHITE, widthFrame=True, skScrollBar=(N,S))

            self.UpdateApplication()

            # EVENT
            self.bind('<Control-KeyPress-s>', lambda e : self.UpdateApplication(add=0))
            self.bind('<Alt-KeyPress-c>', lambda e : self.CopySheet())
            

            # WEIGHT
            self.display.grid_columnconfigure(0, weight=10)
            self.display.grid_rowconfigure(2, weight=10)

            self.frame_menu_top.grid_rowconfigure(0, weight=10)
            self.frame_menu_top.grid_columnconfigure(1, weight=5)
    
        
            self.frame_bottom.grid_rowconfigure(0, weight=10)
            self.frame_bottom.grid_columnconfigure(2, weight=10)

            self.frame_center.grid_rowconfigure(0, weight=10)
            self.frame_center.grid_columnconfigure(0, weight=10)

            self.frame_workspace.grid_columnconfigure(0,weight=10)
        
        else:
            messagebox.showerror("Le fichier n'est pas valide", f"Le fichier {file_name} n'est pas bon ou est corrompus")

if __name__ == "__main__":
    
    app = Application()
    app.mainloop()