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
import csv


class commande:
    def disable_close_button():
        pass  # Cette fonction ne fait rien lorsqu'elle est appelée


    def __init__(self, root):
        self.root = root
        self.root.title("_Archive order Bobina_")
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

         # //navbartop
        navbar_frame2 = tk.Frame(root, bg="azure", padx=27, pady=27)
        navbar_frame2.pack(side="top", fill="x")
        Archife_button = Button(navbar_frame2, text="Archive", command=self.btn_archife, font=("times new roman",15,"italic"),bg="azure",fg="black",activebackground="azure", activeforeground="black", cursor="hand2")
        Archife_button.place(relx=0.5, rely=0.5, anchor="center")
        Archife_button.pack(side="left", padx=40)

        Add_commande = Button(navbar_frame2, text="Crud_Cable", command=self.btn_crudbd, font=("times new roman",15,"italic"),bg="azure",fg="black",activebackground="azure", activeforeground="black", cursor="hand2")
        Add_commande.place(relx=0.5, rely=0.5, anchor="center")
        Add_commande.pack(side="left", padx=40)

        add_user = Button(navbar_frame2, text="Crud_user",command=self.Add_user,font=("times new roman",15,"italic"),bg="azure",fg="black",activebackground="azure", activeforeground="black", cursor="hand2")
        add_user.place(relx=0.5, rely=0.5, anchor="center")
        add_user.pack(side="left", padx=40)

            # //navbarbottom
        self.logo = PhotoImage(file="images/logo.png")
        navbar_frame = tk.Frame(root, bg="lightgrey", padx=10, pady=10)
        navbar_frame.pack(side="bottom", fill="x")
        section1_button = Label(navbar_frame, image=self.logo, width=2105, height=65)
        section1_button.place(relx=0.5, rely=0.5, anchor="center")
        section1_button.pack(side="left", padx=10)

        # //affichage table
        details_frame = Frame(self.root, bd=5, relief=GROOVE, bg="azure", borderwidth=1)
        details_frame.place(x=37, y=120, width=1125, height=650)

        titre_commande = Label(details_frame, text="displays order:", font=("times new roman", 25, "bold"), bg="white")
        titre_commande.place(x=50, y=30)
                
        # Affichage de l'heure
        self.titre_commande_time = Label(details_frame, text="", font=("times new roman", 20, "bold"), bg="lightcyan")
        self.titre_commande_time.place(x=950, y=30)
        self.update_time()


        recherche = Label(details_frame, text="Search by :", font=("times new roman", 20, "italic"), bg="white")
        recherche.place(x=150, y=100)

        self.recherche_selec = ttk.Combobox(details_frame, font=("times new roman", 20), state="readonly")
        self.recherche_selec["values"] = ("APN", "status","matricule","Mc")
        self.recherche_selec.current(0)
        self.recherche_selec.place(x=330, y=100, width=200)

        self.recherche_text = Entry(details_frame, font=("times new roman", 20, "italic"), bg="white")
        self.recherche_text.place(x=550, y=100, width=200)
        btn_Search = Button(details_frame, command=self.rechercher, text="Search", fg="white", bg="lime", activebackground="#00B0F0", activeforeground="white", cursor="hand2", font=("times new roman", 20, "italic")).place(x=770, y=100, width=150, height=30)
        btn__afficher_tout = Button(details_frame,command=self.affiche_commande,  text="Show all", fg="black", bg="bisque", activebackground="#00B0F0", activeforeground="white", cursor="hand2", font=("times new roman", 20, "italic")).place(x=940, y=100, width=150, height=30)


        # btn_export = Button(details_frame, text="Export", fg="white", bg="black", activebackground="#00B0F0", activeforeground="white", cursor="hand2", font=("times new roman", 20, "italic")).place(x=400, y=600, width=200, height=30)
        btn_Refresh = Button(details_frame, command=self.Refresh, text="Refresh", fg="white", bg="aqua", activebackground="#00B0F0", activeforeground="white", cursor="hand2", font=("times new roman", 20, "italic")).place(x=640, y=600, width=200, height=30)
        btn_Logout = Button(details_frame, command=self.Logout, text="Logout", fg="white", bg="tomato", activebackground="#00B0F0", activeforeground="white", cursor="hand2", font=("times new roman", 20, "italic")).place(x=880, y=600, width=200, height=30)

        result_fram = Frame(details_frame, bd=5, relief=GROOVE, bg="white", borderwidth=0)
        result_fram.place(x=10, y=160, width=1100, height=400)

        scroll_x = Scrollbar(result_fram, orient=HORIZONTAL)
        scroll_y = Scrollbar(result_fram, orient=VERTICAL)

        self.table = ttk.Treeview(result_fram,height=23, columns=("id", "Apn", "Status", "Date_demande", "Date_livre", "Matricule", "Heurs","Mc"), xscrollcommand=scroll_x.set, yscrollcommand=scroll_y.set,show="headings",selectmode="browse")
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

        self.table.column("id", width=100,anchor="center")
        self.table.column("Apn", width=130,anchor="center")
        self.table.column("Status", width=100,anchor="center")
        self.table.column("Date_demande", width=150,anchor="center")
        self.table.column("Date_livre", width=150,anchor="center")
        self.table.column("Matricule", width=150,anchor="center")
        self.table.column("Heurs", width=180,anchor="center")
        self.table.column("Mc", width=50,anchor="center")
        
        self.table.pack(fill="both",expand=1)
        self.table.tag_configure("oddrow",background="white")
        self.table.tag_configure("evenrow",background="lightblue")
        self.table.tag_configure("redrow",background="red")
        self.table.tag_configure("aquarow",background="bisque")

        self.table.bind("<ButtonRelease-1>")
        self.affiche_commande()


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
        cur = connect.cursor()
        cur.execute("select * from LP_Spool_demand_archife")
        rows = cur.fetchall()
        if len(rows) != 0:
            self.table.delete(*self.table.get_children())
            for i, (id,Apn,Status,Date_demande,Date_livre,Matricule,Mc) in enumerate(rows, start=1):
                            # Utilisez le tag "oddrow" pour les lignes impaires et "evenrow" pour les lignes paires
                            date1=datetime.strptime(Date_demande,"%d/%m/%Y %H:%M:%S")
                            date2=datetime.strptime(Date_main,"%d/%m/%Y %H:%M:%S")
                            deff=date2 - date1
                            # tags = ("oddrow",) if i % 2 == 1 else ("evenrow",)
                            if deff >timedelta(minutes=30) and Status=="en attende" :
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




    def Refresh(self):
                for widget in root.winfo_children():
                        widget.destroy()
                self.__init__(root)
    def btn_archife(self):
        root.destroy()
        call(["python", "archife.py"])

    def btn_crudbd(self):
        root.destroy()
        call(["python", "Crud_bd.py"])

    def Add_user(self):
        root.destroy()
        call(["python", "add_user.py"])
        


    def Logout(self):
          root.withdraw()
          
        # root.destroy()
        # call(["python", "form_mysql.py"])
    def export(self):
    # Connexion à la base de données
        try:
                connection = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};'
                        'SERVER=130.171.181.100;'
                        'DATABASE=Précontrole;'
                        'UID=Cutting;'
                        'PWD=Cutting')
                cursor = connection.cursor()

                # Exécuter la requête SQL pour sélectionner les données de la table
                cursor.execute("SELECT * FROM LP_Spool_demand_commande")
                rows = cursor.fetchall()

                # Vérifier si des données ont été sélectionnées
                if rows:
                        # Nom du fichier CSV
                        file_path = "donne.csv"

                        # Écrire les données dans le fichier CSV
                        with open(file_path, "w", newline="") as file:
                                writer = csv.writer(file)
                                writer.writerow([i[0] for i in cursor.description])  # Écrire les noms de colonnes
                                writer.writerows(rows)  # Écrire les données

                        messagebox.showinfo("Export terminé", "Les données ont été exportées avec succès vers exported_data.csv")
                else:
                        messagebox.showwarning("Avertissement", "La table est vide, aucune donnée à exporter.")
                
        except pyodbc.Error as e:
                print("Erreur lors de l'exportation des données:", e)
                messagebox.showerror("Erreur", "Une erreur est survenue lors de l'exportation des données.")
        
        finally:
                # Fermer la connexion à la base de données
                if connection:
                        connection.close()

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
        query = f"SELECT * FROM LP_Spool_demand_archife WHERE {column} LIKE ?"
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
                            if deff >timedelta(minutes=30) and Status=="en attende":
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
