from tkinter import *
from PIL import ImageTk
from tkinter import ttk, messagebox
import pyodbc
from datetime import datetime
from datetime import timedelta
import tkinter as tk
from subprocess import call
from time import strftime
import time


class commande:
    def disable_close_button():
        pass  # Cette fonction ne fait rien lorsqu'elle est appelée


    def __init__(self, root):
        self.root = root
        self.root.title("_List Commande_")
        self.root.geometry("1199x900+50+50")
        root.minsize(1199, 880)  
        self.root.resizable(False, False)

        screen_width=self.root.winfo_screenwidth()
        screen_height=self.root.winfo_screenheight()
        window_width=1199
        window_height=880
        position_x=(screen_width//2)-(window_width//2)
        position_y=(screen_height//2)-(window_height//2)
        self.root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")
        self.root.rowconfigure(0,weight=1)
        self.root.columnconfigure(0,weight=1)
        root.protocol("WM_DELETE_WINDOW", self.disable_close_button)

        
        # table
        style=ttk.Style()
        style.theme_use('default')
        style.configure("Treeview",
                        background="white",
                        foreground="black",
                        rowheight=25,
                        fieldbackground="white"
                        )
        style.map('Treeview',background=[('selected',"#347083")])

        # //navbar
        self.logo = PhotoImage(file="images/logo.png")
        navbar_frame = tk.Frame(root, bg="lightgrey", padx=10, pady=10)
        navbar_frame.pack(side="top", fill="x")
        section1_button = Label(navbar_frame, image=self.logo, width=2105, height=65)
        section1_button.place(relx=0.5, rely=0.5, anchor="center")
        section1_button.pack(side="left", padx=10)

        # //affichage table
        details_frame = Frame(self.root, bd=5, relief=GROOVE, bg="lightcyan", borderwidth=1)
        details_frame.place(x=37, y=120, width=1125, height=700)

        message_label = Label(details_frame, text="Welcome to the app by M.yassin Elasri  !", font=('Courier', 12, 'underline'), bg='lightcyan', fg='black')
        message_label.place(x=220,y=630,height=35, anchor="center")

        titre_commande = Label(details_frame, text="Order list:", font=("times new roman", 20, "bold"), bg="white")
        titre_commande.place(x=50, y=30)
                
        # Affichage de l'heure
        self.titre_commande_time = Label(details_frame, text="", font=("times new roman", 20, "bold"), bg="lightcyan")
        self.titre_commande_time.place(x=960, y=10)
        self.update_time()

        # # Affichage de crono

        # self.temps_debut = time.time()
        # self.temps_ecoule = 0
        
        # self.label_temps = tk.Label(root, text="00:00:00", font=("Helvetica", 24))
        # self.label_temps.pack(pady=10)
        
        # self.actualiser_chronometre()


        idbobina = Label(details_frame, text="APN :", font=("times new roman", 20, "italic"), bg="white")
        idbobina.place(x=200, y=100)

        self.id_text_bobina = Entry(details_frame, font=("times new roman", 20, "italic"), bg="white")
        self.id_text_bobina.place(x=350, y=100, width=200)
        self.id_text_bobina.bind("<Return>",self.ecouter_clavier)
        

        btn_Livre = Button(details_frame, command=self.Modifier, text="Livre", fg="black", bg="springgreen", activebackground="#00B0F0", activeforeground="white", cursor="hand2", font=("times new roman", 20, "italic")).place(x=600, y=100, width=200, height=30)
        btn_Livre = Button(details_frame, command=self.Missingcable, text="Missing cable", fg="black", bg="springgreen", activebackground="#00B0F0", activeforeground="white", cursor="hand2", font=("times new roman", 20, "italic")).place(x=850, y=100, width=200, height=30)

        recherche = Label(details_frame, text="Search by :", font=("times new roman", 20, "italic"), bg="white")
        recherche.place(x=200, y=180)

        self.recherche_selec = ttk.Combobox(details_frame, font=("times new roman", 20), state="readonly")
        self.recherche_selec["values"] = ("APN", "status","matricule","Mc")
        self.recherche_selec.current(0)
        self.recherche_selec.place(x=350, y=180, width=200)

        self.recherche_text = Entry(details_frame, font=("times new roman", 20, "italic"), bg="white")
        self.recherche_text.place(x=350, y=220, width=200)
        btn_Search = Button(details_frame, command=self.rechercher, text="Search", fg="black", bg="bisque", activebackground="#00B0F0", activeforeground="white", cursor="hand2", font=("times new roman", 20, "italic")).place(x=600, y=180, width=200, height=30)
        btn_Search = Button(details_frame,command=self.affiche_commande,  text="Show all", fg="black", bg="bisque", activebackground="#00B0F0", activeforeground="white", cursor="hand2", font=("times new roman", 20, "italic")).place(x=600, y=220, width=200, height=30)

        # btn_export = Button(details_frame, command=self.Modifier, text="Export", fg="white", bg="black", activebackground="#00B0F0", activeforeground="white", cursor="hand2", font=("times new roman", 20, "italic")).place(x=420, y=650, width=200, height=30)
        btn_Refresh = Button(details_frame, command=self.Refresh, text="Refresh", fg="white", bg="aqua", activebackground="#00B0F0", activeforeground="white", cursor="hand2", font=("times new roman", 20, "italic")).place(x=660, y=650, width=200, height=30)
        btn_Logout = Button(details_frame, command=self.Logout, text="Logout", fg="white", bg="tomato", activebackground="#00B0F0", activeforeground="white", cursor="hand2", font=("times new roman", 20, "italic")).place(x=900, y=650, width=200, height=30)

        result_fram = Frame(details_frame, bd=5, relief=GROOVE, bg="white", borderwidth=0)
        result_fram.place(x=20, y=290, width=1080, height=320)

        scroll_x = Scrollbar(result_fram, orient=HORIZONTAL)
        scroll_y = Scrollbar(result_fram, orient=VERTICAL)

        self.table = ttk.Treeview(result_fram, columns=("id", "Apn", "Status", "Date_demande", "Date_livre", "Matricule", "Heurs","Mc"), xscrollcommand=scroll_x.set, yscrollcommand=scroll_y.set,show="headings",selectmode="browse")
        scroll_x.pack(side=BOTTOM, fill=X)
        scroll_y.pack(side=RIGHT, fill=Y)

        self.table.heading("id", text="Id",anchor="center")
        self.table.heading("Apn", text="Apn",anchor="center")
        self.table.heading("Status", text="Status",anchor="center")
        self.table.heading("Date_demande", text="Date_demande",anchor="center")
        self.table.heading("Date_livre", text="Date_livre",anchor="center")
        self.table.heading("Matricule", text="Matricule",anchor="center")
        self.table.heading("Heurs", text="Heurs",anchor="center")
        self.table.heading("Mc", text="Mc",anchor="center")

        self.table["show"] = "headings"


        self.table.column("id", width=80,anchor="center")
        self.table.column("Apn", width=80,anchor="center")
        self.table.column("Status", width=80,anchor="center")
        self.table.column("Date_demande", width=100,anchor="center")
        self.table.column("Date_livre", width=100,anchor="center")
        self.table.column("Matricule", width=80,anchor="center")
        self.table.column("Heurs", width=80,anchor="center")
        self.table.column("Mc", width=20,anchor="center")

                
        self.table.pack(fill="both",expand=1)
        self.table.tag_configure("oddrow",background="white")
        self.table.tag_configure("evenrow",background="lightblue")
        self.table.tag_configure("redrow",background="red")
        self.table.tag_configure("aquarow",background="bisque")

        self.table.pack()
        self.table.bind("<ButtonRelease-1>")
        self.affiche_commande()
    def ecouter_clavier(self,event):
        self.Modifier()

    def update_time(self):
        time_string = strftime("%I:%M:%S %p")
        self.titre_commande_time.config(text=time_string)
        self.root.after(1000, self.update_time)





    def affiche_commande(self):
                maintenantm = datetime.now()
                Date_main =maintenantm.strftime("%d/%m/%Y %H:%M:%S")
                connect = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};'
                        'SERVER=130.171.181.100;'
                        'DATABASE=Précontrole;'
                        'UID=Cutting;'
                        'PWD=Cutting')
                cur=connect.cursor()
                #  where password=%s and matricule=%s", (self.text_password.get(), self.text_matricule.get())
                cur.execute("select * from LP_Spool_demand_commande where status IN (?,?)",("en attende","Missing cable"))
                rows=cur.fetchall()
                
                if len(rows) !=0:
                        self.table.delete(*self.table.get_children())
                        for i, (id,Apn,Status,Date_demande,Date_livre,Matricule,Mc) in enumerate(rows, start=1):
                                
                                # Utilisez le tag "oddrow" pour les lignes impaires et "evenrow" pour les lignes paires
                                date1=datetime.strptime(Date_demande,"%d/%m/%Y %H:%M:%S")
                                date2=datetime.strptime(Date_main,"%d/%m/%Y %H:%M:%S")
                                deff=date2 - date1
                                # tags = ("oddrow",) if i % 2 == 1 else ("evenrow",)
                                if deff >timedelta(minutes=30):
                                        tags=("redrow",)
                                elif Status=="Missing cable":
                                        tags = ("aquarow",)
                                elif i%2==1:
                                        tags = ("oddrow",)
                                else:
                                        tags=("evenrow",)
                                self.table.insert("", END, values=(id,Apn,Status,Date_demande,Date_livre,Matricule,deff,Mc), tags=tags)
                                # self.table.insert('','end',values=('','','mmmmmmmmmmmm','','','','','mmmmmmmmmmmm'))




    def Modifier(self):
        if self.id_text_bobina.get() == "":
            messagebox.showerror("Erreur", "You have not completed the required fields", parent=self.root)
        elif self.id_text_bobina.get() != "":
            connect = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};'
                    'SERVER=130.171.181.100;'
                    'DATABASE=Précontrole;'
                    'UID=Cutting;'
                    'PWD=Cutting')
            cur = connect.cursor()
            cur.execute("select * from LP_Spool_demand_commande where apn=?", (self.id_text_bobina.get()))
            row = cur.fetchone()
            if row == None:
                messagebox.showerror("Erreur", "APN Invalide", parent=self.root)
            else:
                status_m = "Livre"
                maintenant = datetime.now()
                date_livre = maintenant.strftime("%d/%m/%Y %H:%M:%S")
                cur.execute("update LP_Spool_demand_archife set status=?, date_livre=? where apn=? and (status=? OR status=?)", (status_m, date_livre, self.id_text_bobina.get(),"en attende","Missing cable"))
                cur.execute("delete from LP_Spool_demand_commande where apn =?",self.id_text_bobina.get())
                connect.commit()
                self.affiche_commande()
                self.vide()
            connect.close()
    def Missingcable(self):
        
        if self.id_text_bobina.get() == "":
            messagebox.showerror("Erreur", "You have not completed the required fields", parent=self.root)
        else:
            connect = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};'
                    'SERVER=130.171.181.100;'
                    'DATABASE=Précontrole;'
                    'UID=Cutting;'
                    'PWD=Cutting')
            cur = connect.cursor()
            cur.execute("select * from LP_Spool_demand_commande where apn=?", (self.id_text_bobina.get()))
            row = cur.fetchone()
            if row == None:
                messagebox.showerror("Erreur", "APN Invalide", parent=self.root)
            else:
                status_m = "Missing cable"
                maintenant = datetime.now()
                date_livre = maintenant.strftime("%d/%m/%Y %H:%M:%S")
                cur.execute("update LP_Spool_demand_archife set status=?, date_livre=? where apn=? and status=?", (status_m, date_livre, self.id_text_bobina.get(),"en attende"))
                cur.execute("update LP_Spool_demand_commande set status=?, date_livre=? where apn=? and status=?", (status_m, date_livre, self.id_text_bobina.get(),"en attende"))
                connect.commit()
                self.affiche_commande()
                self.vide()
            connect.close()

    def Refresh(self):
        # root.destroy()
        # call(["python", "listcommande.py"])
        for widget in root.winfo_children():
                widget.destroy()
        self.__init__(root)

    def Logout(self):
        # root.withdraw()
        root.destroy()
        call(["python", "form_mysql.py"])
    def vide(self):
        if self.id_text_bobina.get():
                self.id_text_bobina.delete(0,"end")

    def rechercher(self):
        maintenantm = datetime.now()
        Date_main =maintenantm.strftime("%d/%m/%Y %H:%M:%S")
        column = self.recherche_selec.get()
        search_value = self.recherche_text.get()
        connect = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};'
                'SERVER=130.171.181.100;'
                'DATABASE=Précontrole;'
                'UID=Cutting;'
                'PWD=Cutting')
        cur = connect.cursor()
        query = f"SELECT * FROM LP_Spool_demand_commande WHERE {column} LIKE ?"
        cur.execute(query, ('%' + search_value + '%',))
        rows = cur.fetchall()
        if len(rows) != 0:
            self.table.delete(*self.table.get_children())
            for i, (id,Apn,Status,Date_demande,Date_livre,Matricule,Mc) in enumerate(rows, start=1):
                                # Utilisez le tag "oddrow" pour les lignes impaires et "evenrow" pour les lignes paires
                date1=datetime.strptime(Date_demande,"%d/%m/%Y %H:%M:%S")
                date2=datetime.strptime(Date_main,"%d/%m/%Y %H:%M:%S")
                deff=date2 - date1
                # tags = ("oddrow",) if i % 2 == 1 else ("evenrow",)
                if deff >timedelta(minutes=30):
                        tags=("redrow",)
                elif Status=="Missing cable":
                        tags = ("aquarow",)
                elif i%2==1:
                        tags = ("oddrow",)
                else:
                        tags=("evenrow",)
                self.table.insert("", END, values=(id,Apn,Status,Date_demande,Date_livre,Matricule,deff,Mc), tags=tags)

        connect.commit()
        connect.close()

root = Tk()
obj = commande(root)
root.mainloop()
