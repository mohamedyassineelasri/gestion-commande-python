from tkinter import *
from PIL import ImageTk
from tkinter import ttk,messagebox
import pyodbc
from datetime import datetime
from datetime import timedelta
import tkinter as tk
from subprocess import call
from time import strftime
import time
import pandas as pd
import csv
import os

from form_mysql import FormMysql


class commande :
        matri = FormMysql.matri
        def disable_close_button():
                pass  # Cette fonction ne fait rien lorsqu'elle est appelée


        def __init__(self,root):
                
                self.root=root
                self.temps_debut_commande = None
                self.root.title("_Order Bobina_")
                self.root.geometry("1199x900+50+50")
                root.minsize(1199, 880)  
                self.root.resizable(False, False)
                # root.overrideredirect(not root.overrideredirect())
                root.protocol("WM_DELETE_WINDOW", self.disable_close_button)

                screen_width=self.root.winfo_screenwidth()
                screen_height=self.root.winfo_screenheight()
                window_width=1199
                window_height=880
                position_x=(screen_width//2)-(window_width//2)
                position_y=(screen_height//2)-(window_height//2)
                self.root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")
                self.root.rowconfigure(0,weight=1)
                self.root.columnconfigure(0,weight=1)


                style=ttk.Style()
                style.theme_use('default')
                style.configure("Treeview",
                                background="#D3D3D3",
                                foreground="black",
                                rowheight=25,
                                fieldbackground="white"
                                )
                style.map('Treeview',background=[('selected',"#347083")])


                # //navbar
               
                self.logo=PhotoImage(file="images/logo.png")
                navbar_frame = tk.Frame(root, bg="lightgrey", padx=10, pady=10)
                navbar_frame.pack(side="top", fill="x")
                # Création des boutons de la barre de navigation
                section1_button = Label(navbar_frame, image=self.logo, width=2105, height=65)
                section1_button.place(relx=0.5, rely=0.5, anchor="center")
                section1_button.pack(side="left", padx=10)
                

                # les variable
                self.id_text_bobina=StringVar()
                
                # //background


                # self.bg=ImageTk.PhotoImage(file="images/aptiv.jpg")
                # self.bg_images=Label(self.root,image=self.bg).place(x=0,y=0,relwidth=1,relheight=1,width=15,height=60)
                gestion_fram=Frame(self.root,bd=5,relief=GROOVE,bg="lightcyan" ,borderwidth=1)
                gestion_fram.place(x=20,y=150,width=350,height=650)
                gestion_title=Label(gestion_fram,text="Add Cables :",font=("times new roman",25,"bold"),bg="floralwhite")
                gestion_title.place(x=50,y=50)
                Mc=Label(gestion_fram,text="Mc :",font=("times new roman",20,"italic"),bg="floralwhite")
                Mc.place(x=30,y=150)
                
                self.text_Mc = ttk.Combobox(gestion_fram,font=("times new roman",20,"italic"), state="readonly")
                self.text_Mc["values"] = ("Mc68", "Mc69","Mc70","Mc71","Mc72")
                # self.text_Mc.current(0)
                self.text_Mc.place(x=120,y=150,width=200)

                
                idbobina=Label(gestion_fram,text="APN :",font=("times new roman",20,"italic"),bg="floralwhite")
                idbobina.place(x=30,y=200)
                
                self.id_text_bobina=Entry(gestion_fram,textvariable=self.id_text_bobina,font=("times new roman",20,"italic"),bg="floralwhite")
                self.id_text_bobina.place(x=120,y=200,width=200)
                self.id_text_bobina.bind("<Return>",self.ecouter_clavier)
                
                

                commande_btn=Button(self.root,text="Order",command=self.ajouter_en_btn,fg="white",bg="#00B0F0",activebackground="#00B0F0",activeforeground="white",cursor="hand2",font=("times new roman",20,"italic")).place(x=80,y=420,width=230,height=35)
                commande_btn_reini=Button(self.root,text="login",command=self.login,fg="white",bg="grey",activebackground="#00B0F0",activeforeground="white",cursor="hand2",font=("times new roman",20,"italic")).place(x=80,y=540,width=230,height=35)
                commande_btn_reini=Button(self.root,text="Delete",command=self.delete,fg="white",bg="red",activebackground="#00B0F0",activeforeground="white",cursor="hand2",font=("times new roman",20,"italic")).place(x=80,y=480,width=230,height=35)


                
                # //affichage table
                details_frame=Frame(self.root,bd=5,relief=GROOVE,bg="lightcyan",borderwidth=1)
                details_frame.place(x=380,y=150,width=800,height=650)
                titre_commande=Label(details_frame,text="Order:",font=("times new roman",25,"bold"),bg="floralwhite")
                titre_commande.place(x=50,y=50)

                message_label = Label(details_frame, text=f"Matricule :{self.matri} Welcome to the app by M.yassin Elasri  !", font=('Courier', 12, 'underline'), bg='lightcyan', fg='black')
                message_label.place(x=280,y=630,height=35, anchor="center")
                                
                # Affichage de l'heure
                self.titre_commande_time = Label(details_frame, text="", font=("times new roman",15,"bold"), bg="lightcyan")
                self.titre_commande_time.place(x=570,y=1)
                self.update_time()


                commande_btn=Button(details_frame,text="Export",command=self.export,fg="white",bg="black",activebackground="#00B0F0",activeforeground="white",cursor="hand2",font=("times new roman",15,"italic")).place(x=190,y=100,width=150,height=30)
                commande_btn=Button(details_frame,command=self.Refresh,text="Refresh",fg="white",bg="aqua",activebackground="#00B0F0",activeforeground="white",cursor="hand2",font=("times new roman",15,"italic")).place(x=370,y=100,width=150,height=30)
                commande_btn=Button(details_frame,command=self.Logout,text="Logout",fg="white",bg="tomato",activebackground="#00B0F0",activeforeground="white",cursor="hand2",font=("times new roman",15,"italic")).place(x=550,y=100,width=150,height=30)

                result_fram=Frame(details_frame,bd=5,relief=GROOVE,bg="white" ,borderwidth=1)
                result_fram.place(x=10,y=150,width=775,height=420)
                
                
                scroll_x=Scrollbar(result_fram,orient=HORIZONTAL)
                scroll_y=Scrollbar(result_fram,orient=VERTICAL)

                # self.table=ttk.Treeview(result_fram,columns=("Apn","Status","Date_demande","Date_livre"),xscrollcommand=scroll_x.set,yscrollcommand=scroll_y.set)
                self.table=ttk.Treeview(result_fram,columns=("id","Apn","Status","Date_demande","Date_livre","Matricule","Heurs","Mc"),xscrollcommand=scroll_x.set,yscrollcommand=scroll_y.set,show="headings",selectmode="browse")
                scroll_x.pack(side=BOTTOM,fill=X)
                scroll_y.pack(side=RIGHT,fill=Y)

                self.table.heading("id",text="Id",anchor="center")
                self.table.heading("Apn",text="Apn",anchor="center")
                self.table.heading("Status",text="Status",anchor="center")
                self.table.heading("Date_demande",text="Date_demande",anchor="center")
                self.table.heading("Date_livre",text="Date_livre",anchor="center")
                self.table.heading("Matricule",text="Matricule",anchor="center")
                self.table.heading("Heurs",text="Heurs",anchor="center")
                self.table.heading("Mc",text="Mc",anchor="center")

                self.table["show"]="headings"

                self.table.column("id",width=20,anchor="center")
                self.table.column("Apn",width=40,anchor="center")
                self.table.column("Status",width=40,anchor="center")
                self.table.column("Date_demande",width=50,anchor="center")
                self.table.column("Date_livre",width=50,anchor="center")
                self.table.column("Matricule",width=30,anchor="center")
                self.table.column("Heurs",width=60,anchor="center")                
                self.table.column("Mc",width=20,anchor="center")                


                

                self.table.pack(fill="both",expand=1)
                
                self.table.tag_configure("oddrow",background="white")
                self.table.tag_configure("evenrow",background="lightblue")
                self.table.tag_configure("redrow",background="red")
                self.table.tag_configure("aquarow",background="bisque")
                self.table.bind("<ButtonRelease-1>")
                self.affiche_commande()
        def ecouter_clavier(self,event):
                if self.text_Mc.get() != "":
                     self.ajouter_commande()
                else:
                        messagebox.showerror("Erreur", "Mc obligatory", parent=self.root)
                        self.vide()
        def ajouter_en_btn(self):
                if self.text_Mc.get() != "":
                     self.ajouter_commande()
                else:
                        messagebox.showerror("Erreur", "Mc obligatory", parent=self.root)
                        self.vide()




        def update_time(self):
                time_string = strftime("%I:%M:%S %p")
                self.titre_commande_time.config(text=time_string)
                self.root.after(1000, self.update_time)

        def ajouter_commande(self):
                try:
                        # Establish a single database connection
                        connect = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};'
                                                'SERVER=130.171.181.100;'
                                                'DATABASE=Précontrole;'
                                                'UID=Cutting;'
                                                'PWD=Cutting')
                        cur = connect.cursor()

                        # Fetch the maximum ID from the database
                        cur.execute("SELECT MAX(id) FROM LP_Spool_demand_archife")
                        max_id = cur.fetchone()[0]
                        id_commande = max_id + 1 if max_id is not None else 1

                        # Prepare other data
                        maintenant = datetime.now()
                        Date_demande = maintenant.strftime("%d/%m/%Y %H:%M:%S")
                        statu = "en attende"
                        Date_livre = ""
                        heurs = 60

                        # Check for empty bobina ID
                        if not self.id_text_bobina.get():
                                messagebox.showerror("Erreur", "You have not completed the required fields", parent=self.root)
                                return

                        # Check if APN Bobina is valid
                        cur.execute("SELECT * FROM LP_Spool_demand_bobine WHERE apn = ?", (self.id_text_bobina.get(),))
                        if cur.fetchone() is None:
                                messagebox.showerror("Erreur", "Invalid APN Bobina ", parent=self.root)
                                return

                        # Check if the order already exists
                        cur.execute("SELECT * FROM LP_Spool_demand_commande WHERE apn = ?", (self.id_text_bobina.get(),))
                        if cur.fetchone() is not None:
                                messagebox.showerror("Erreur", "Order already existing", parent=self.root)
                                return

                        # Insert new records if the number of orders is less than 4
                        cur.execute("SELECT COUNT(*) FROM LP_Spool_demand_commande")
                        nbr_lignes = cur.fetchone()[0]
                        if nbr_lignes < 20:
                                cur.execute("INSERT INTO LP_Spool_demand_archife (apn,status,Date_demande,date_livre,matricule,Mc) VALUES (?, ?, ?, ?, ?, ?)",
                                                (self.id_text_bobina.get(), statu, Date_demande, Date_livre, self.matri,self.text_Mc.get()))
                                cur.execute("INSERT INTO LP_Spool_demand_commande (id,apn,status,Date_demande,date_livre,matricule,Mc) VALUES (?, ?, ?, ?, ?, ?, ?)",
                                                (id_commande,self.id_text_bobina.get(), statu, Date_demande, Date_livre, self.matri,self.text_Mc.get()))
                                connect.commit()
                                self.affiche_commande()
                                self.vide()
                                self.id_text_bobina.set("")
                        else:
                                messagebox.showerror("Erreur", "Orders larger than 20", parent=self.root)

                except pyodbc.Error as e:
                        messagebox.showerror("Database Error", str(e), parent=self.root)
                finally:
                        # Clean up the cursor and connection
                        cur.close()
                        connect.close()


        def affiche_commande(self):
                maintenantm = datetime.now()
                Date_main =maintenantm.strftime("%d/%m/%Y %H:%M:%S")
                connect = pyodbc.connect(
                                    'DRIVER={ODBC Driver 17 for SQL Server};'
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
                                # if deff>"01:00:00":
                                #         self.table.tag_configure("red",background="red")

                                self.table.insert("", END, values=(id,Apn,Status,Date_demande,Date_livre,Matricule,deff,Mc), tags=tags)
                                # self.table.insert('','end',values=('','','mmmmmmmmmmmm','','','','','mmmmmmmmmmmm'))

                connect.commit()
                connect.close()


# delete
        def delete(self):
                if self.id_text_bobina.get()=="":
                        messagebox.showerror("Erreur","You have not completed the required fields",parent=self.root)
                else:
                        connect = pyodbc.connect(
                                        'DRIVER={ODBC Driver 17 for SQL Server};'
                                        'SERVER=130.171.181.100;'
                                        'DATABASE=Précontrole;'
                                        'UID=Cutting;'
                                        'PWD=Cutting')
                        cur=connect.cursor()
                        cur.execute("select * from LP_Spool_demand_commande where apn=?",(self.id_text_bobina.get()))
                        row=cur.fetchone()
                        if row==None:
                                messagebox.showerror("Erreur","APN Invalid",parent=self.root)
                        else:
                                status_m = "Delete"
                                maintenant = datetime.now()
                                date_livre = maintenant.strftime("%d/%m/%Y %H:%M:%S")
                                connect = pyodbc.connect(
                                                        'DRIVER={ODBC Driver 17 for SQL Server};'
                                                        'SERVER=130.171.181.100;'
                                                        'DATABASE=Précontrole;'
                                                        'UID=Cutting;'
                                                        'PWD=Cutting')
                                cur=connect.cursor()
                                cur.execute("delete from LP_Spool_demand_commande where apn =? and matricule=?",(self.id_text_bobina.get(),self.matri))
                                cur.execute("update LP_Spool_demand_archife set status=?, date_livre=? where apn=? and matricule=? and (status=? OR status=?)", (status_m,date_livre, self.id_text_bobina.get(),self.matri,"en attende","Missing cable"))

                                connect.commit()
                                self.affiche_commande()
                                self.vide()
                                connect.close()
                                # messagebox.showinfo("Success","delete",parent=self.root)
                                
                                


        def Refresh(self):

                for widget in root.winfo_children():
                        widget.destroy()
                self.__init__(root)

                
        def Logout(self):
                root.withdraw()
                # call(["python","form_mysql.py"])
        def login(self):
                root.destroy()
                call(["python","form_mysql.py"])
        def vide(self):
                if self.id_text_bobina.get():
                        self.id_text_bobina.delete(0,"end")
        def export(self):
                # Connexion à la base de données
                try:
                        connection = pyodbc.connect(
                                'DRIVER={ODBC Driver 17 for SQL Server};'
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
                                # Créer le dossier 'export' sur le bureau s'il n'existe pas
                                desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
                                export_folder = os.path.join(desktop_path, 'export')
                                if not os.path.exists(export_folder):
                                        os.makedirs(export_folder)

                                        # Nom du fichier CSV
                                        file_path = os.path.join(export_folder, 'exported_data.csv')

                                        # Écrire les données dans le fichier CSV
                                with open(file_path, "w", newline="") as file:
                                        writer = csv.writer(file)
                                        writer.writerow([i[0] for i in cursor.description])  # Écrire les noms de colonnes
                                        writer.writerows(rows)  # Écrire les données

                                        messagebox.showinfo("Export terminé", f"Les données ont été exportées avec succès vers {file_path}")
                        else:
                                messagebox.showwarning("Avertissement", "La table est vide, aucune donnée à exporter.")

                except pyodbc.Error as e:
                        print("Erreur lors de l'exportation des données:", e)
                        messagebox.showerror("Erreur", "Une erreur est survenue lors de l'exportation des données.")

                finally:
                        # Fermer la connexion à la base de données
                        if connection:
                                connection.close()





root = Tk()
obj = commande(root)
root.mainloop()