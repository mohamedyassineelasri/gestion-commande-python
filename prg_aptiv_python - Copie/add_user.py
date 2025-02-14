from tkinter import *
from PIL import ImageTk
from tkinter import ttk,messagebox
import pyodbc
from datetime import datetime
import tkinter as tk
from subprocess import call
from time import strftime
import time


class commande :
        # matri = FormMysql.matri
        def disable_close_button():
                pass  # Cette fonction ne fait rien lorsqu'elle est appelée


        def __init__(self,root):
                self.root=root
                self.root.title("_Add/Delete Bobina_")
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
                                background="#D3D3D3",
                                foreground="black",
                                rowheight=25,
                                fieldbackground="white"
                                )
                style.map('Treeview',background=[('selected',"#347083")])
                
                                
                # //navbarbottom
                self.logo = PhotoImage(file="images/logo.png")
                navbar_frame = tk.Frame(root, bg="lightgrey", padx=10, pady=10)
                navbar_frame.pack(side="bottom", fill="x")
                section1_button = Label(navbar_frame, image=self.logo, width=2105, height=65)
                section1_button.place(relx=0.5, rely=0.5, anchor="center")
                section1_button.pack(side="left", padx=10)
                

                # //navbartop
                navbar_frame2 = tk.Frame(root, bg="azure", padx=27, pady=27)
                navbar_frame2.pack(side="top", fill="x")
                Archife_button = Button(navbar_frame2, text="Archive", command=self.btn_archife, font=("times new roman",15,"italic"),bg="azure",fg="black",activebackground="azure", activeforeground="black", cursor="hand2")
                Archife_button.place(relx=0.5, rely=0.5, anchor="center")
                Archife_button.pack(side="left", padx=40)

                Add_bobine = Button(navbar_frame2, text="Crud_Cable", command=self.btn_crudbd, font=("times new roman",15,"italic"),bg="azure",fg="black",activebackground="azure", activeforeground="black", cursor="hand2")
                Add_bobine.place(relx=0.5, rely=0.5, anchor="center")
                Add_bobine.pack(side="left", padx=40)

                add_user = Button(navbar_frame2, text="Crud_user",command=self.Add_user,font=("times new roman",15,"italic"),bg="azure",fg="black",activebackground="azure", activeforeground="black", cursor="hand2")
                add_user.place(relx=0.5, rely=0.5, anchor="center")
                add_user.pack(side="left", padx=40)




                # les variable
                

                gestion_fram=Frame(self.root,bd=5,relief=GROOVE,bg="lightcyan" ,borderwidth=1)
                gestion_fram.place(x=35,y=150,width=400,height=600)
                gestion_title=Label(gestion_fram,text="Add User :",font=("times new roman",25,"bold"),bg="floralwhite")
                gestion_title.place(x=50,y=50)

                Matricule=Label(gestion_fram,text="Matricule :",font=("times new roman",20,"italic"),bg="floralwhite")
                Matricule.place(x=10,y=150)
                self.id_text_Matricule=Entry(gestion_fram,font=("times new roman",20,"italic"),bg="floralwhite")
                self.id_text_Matricule.place(x=140,y=150,width=220)

                Password=Label(gestion_fram,text="Password :",font=("times new roman",20,"italic"),bg="floralwhite")
                Password.place(x=10,y=200)
                self.text_Password=Entry(gestion_fram,font=("times new roman",20,"italic"),bg="floralwhite")
                self.text_Password.place(x=140,y=200,width=220)

                Role=Label(gestion_fram,text="Role :",font=("times new roman",20,"italic"),bg="floralwhite")
                Role.place(x=10,y=250)
                # self.text_Role=Entry(gestion_fram,font=("times new roman",20,"italic"),bg="floralwhite")
                # self.text_Role.place(x=140,y=250,width=220)

                self.text_Role = ttk.Combobox(gestion_fram,font=("times new roman",20,"italic"), state="readonly")
                self.text_Role["values"] = ("HG", "logistique","admin")
                # self.text_Role.current(0)
                self.text_Role.place(x=140,y=250,width=220)



                btn_Add=Button(gestion_fram,text="Add",command=self.ajouter_user,fg="white",bg="#00B0F0",activebackground="#00B0F0",activeforeground="white",cursor="hand2",font=("times new roman",20,"italic")).place(x=80,y=400,width=250,height=35)
                btn_delete=Button(gestion_fram,text="delete",command=self.delete,fg="white",bg="red",activebackground="#00B0F0",activeforeground="white",cursor="hand2",font=("times new roman",20,"italic")).place(x=80,y=450,width=250,height=35)
                
                # //affichage table
                details_frame=Frame(self.root,bd=5,relief=GROOVE,bg="lightcyan",borderwidth=1)
                details_frame.place(x=470,y=150,width=700,height=600)
                titre_commande=Label(details_frame,text="Users :",font=("times new roman",25,"bold"),bg="floralwhite")
                titre_commande.place(x=50,y=50)
                                
                # Affichage de l'heure
                self.titre_commande_time = Label(details_frame, text="", font=("times new roman", 30, "bold"), bg="lightcyan")
                self.titre_commande_time.place(x=1050,y=30)
                self.update_time()

                commande_btn=Button(details_frame,command=self.Refresh,text="Refresh",fg="white",bg="aqua",activebackground="#00B0F0",activeforeground="white",cursor="hand2",font=("times new roman",20,"italic")).place(x=350,y=100,width=150,height=30)
                commande_btn=Button(details_frame,command=self.Logout,text="Logout",fg="white",bg="tomato",activebackground="#00B0F0",activeforeground="white",cursor="hand2",font=("times new roman",20,"italic")).place(x=520,y=100,width=150,height=30)

                result_fram=Frame(details_frame,bd=5,relief=GROOVE,bg="white" ,borderwidth=1)
                result_fram.place(x=10,y=150,width=680,height=420)
                
                scroll_x=Scrollbar(result_fram,orient=HORIZONTAL)
                scroll_y=Scrollbar(result_fram,orient=VERTICAL)

                # self.table=ttk.Treeview(result_fram,columns=("Apn","Status","Date_demande","Date_livre"),xscrollcommand=scroll_x.set,yscrollcommand=scroll_y.set)
                self.table=ttk.Treeview(result_fram,columns=("id","matricule","password","role"),xscrollcommand=scroll_x.set,yscrollcommand=scroll_y.set,show="headings",selectmode="browse")
                scroll_x.pack(side=BOTTOM,fill=X)
                scroll_x.pack(side=BOTTOM,fill=X)
                scroll_y.pack(side=RIGHT,fill=Y)

                self.table.heading("id",text="id",anchor="center")
                self.table.heading("matricule",text="matricule",anchor="center")
                self.table.heading("password",text="password",anchor="center")
                self.table.heading("role",text="role",anchor="center")


                self.table["show"]="headings"

                self.table.column("id",width=50,anchor="center")
                self.table.column("matricule",width=50,anchor="center")
                self.table.column("password",width=100,anchor="center")
                self.table.column("role",width=100,anchor="center")

                self.table.pack(fill="both",expand=1)
                self.table.tag_configure("oddrow",background="white")
                self.table.tag_configure("evenrow",background="lightblue")
                self.table.bind("<ButtonRelease-1>")
                self.affiche_user()

        def update_time(self):
                time_string = strftime("%I:%M:%S %p")
                self.titre_commande_time.config(text=time_string)
                self.root.after(1000, self.update_time)

        def ajouter_user(self):

                connect = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};'
                        'SERVER=130.171.181.100;'
                        'DATABASE=Précontrole;'
                        'UID=Cutting;'
                        'PWD=Cutting')
                cur=connect.cursor()
                cur.execute("select * from LP_Spool_demand_users")
                rows=cur.fetchall()
                
                print(len(rows))
                
                # id_user=len(rows)+1
                id_user=1
                cur.execute("SELECT MAX(id) FROM LP_Spool_demand_users")
                max_id = cur.fetchone()[0]
                if max_id is not None:
                        id_user = max_id + 1

                maintenant = datetime.now()
                Date_ajouter =maintenant.strftime("%d/%m/%Y %H:%M:%S")
                




                if self.id_text_Matricule.get()=="" or self.text_Password.get()=="" or self.text_Role.get()=="":
                        messagebox.showerror("Erreur","You have not completed the required fields",parent=self.root)
                else:
                        connect = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};'
                                'SERVER=130.171.181.100;'
                                'DATABASE=Précontrole;'
                                'UID=Cutting;'
                                'PWD=Cutting')
                        cur=connect.cursor()
                        cur.execute("select * from LP_Spool_demand_users where matricule=?",(self.id_text_Matricule.get()))
                        row=cur.fetchone()
                        if row==None:
                                connect = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};'
                                        'SERVER=130.171.181.100;'
                                        'DATABASE=Précontrole;'
                                        'UID=Cutting;'
                                        'PWD=Cutting')
                                cur=connect.cursor()
                                cur.execute("SET IDENTITY_INSERT LP_Spool_demand_users ON")
                                cur.execute("insert into LP_Spool_demand_users (id,matricule,password,role) values(?,?,?,?)",(id_user,self.id_text_Matricule.get(),self.text_Password.get(),self.text_Role.get()))
                                connect.commit()
                                self.affiche_user()
                                self.vide()
                                connect.close()
                                messagebox.showinfo("Success","Good Add")
                        else:
                                messagebox.showerror("Erreur","users déja exsite",parent=self.root)

                                
        

        def affiche_user(self):
                try:
                        # Établir la connexion à la base de données
                        with pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};'
                                        'SERVER=130.171.181.100;'
                                        'DATABASE=Précontrole;'
                                        'UID=Cutting;'
                                        'PWD=Cutting') as connect:
                                cur = connect.cursor()
                                cur.execute("SELECT * FROM LP_Spool_demand_users")
                                rows = cur.fetchall()

                        if rows:
                                self.table.delete(*self.table.get_children())
                                for i, row in enumerate(rows, start=1):
                                # Formatage des données de la ligne si nécessaire
                                        formatted_row = [str(item) for item in row]
                                        tags = ("oddrow",) if i % 2 == 1 else ("evenrow",)
                                        self.table.insert("", 'end', values=formatted_row, tags=tags)

                except pyodbc.Error as e:
                        messagebox.showerror("Erreur de base de données", str(e), parent=self.root)


# delete
        def delete(self):
                if self.id_text_Matricule.get()=="":
                        messagebox.showerror("Erreur","Matricule obligatory",parent=self.root)
                else:
                        connect = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};'
                                'SERVER=130.171.181.100;'
                                'DATABASE=Précontrole;'
                                'UID=Cutting;'
                                'PWD=Cutting')
                        cur=connect.cursor()
                        cur.execute("select * from LP_Spool_demand_users where matricule=?",(self.id_text_Matricule.get()))
                        row=cur.fetchone()
                        if row==None:
                                messagebox.showerror("Erreur","Matricule Invalid",parent=self.root)
                        else:
                                connect = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};'
                                        'SERVER=130.171.181.100;'
                                        'DATABASE=Précontrole;'
                                        'UID=Cutting;'
                                        'PWD=Cutting')
                                cur=connect.cursor()
                                cur.execute("delete from LP_Spool_demand_users where matricule =? ",(self.id_text_Matricule.get()))

                                connect.commit()
                                self.affiche_user()
                                connect.close()
                                messagebox.showinfo("Success","delete",parent=self.root)
                                

        def Refresh(self):
                for widget in root.winfo_children():
                        widget.destroy()
                self.__init__(root)
        def Logout(self):

                root.withdraw()
        
        def btn_archife(self):
                root.destroy()
                call(["python", "archife.py"])

        def btn_crudbd(self):
                root.destroy()
                call(["python", "Crud_bd.py"])
                
        def Add_user(self):
                root.destroy()
                call(["python", "add_user.py"])
                
        def vide(self):
                if self.id_text_Matricule.get():
                        self.id_text_Matricule.delete(0,"end")
                        self.text_Password.delete(0,"end")
                        self.text_Role.delete(0,"end")




root = Tk()
obj = commande(root)
root.mainloop()