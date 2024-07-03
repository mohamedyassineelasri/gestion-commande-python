from tkinter import *
from tkinter import messagebox
import pyodbc
from PIL import ImageTk
from subprocess import call
import add_commande

class FormMysql:
    def __init__(self, root=None):
        self.root = root
        if self.root is None:
            self.root = Tk()
        self.root.title("Login")
        self.root.geometry("1199x600+100+50")
        root.minsize(1199, 600)  
        self.root.resizable(False, False)
        # root.overrideredirect(not root.overrideredirect())

        screen_width=self.root.winfo_screenwidth()
        screen_height=self.root.winfo_screenheight()
        window_width=1199
        window_height=600
        position_x=(screen_width//2)-(window_width//2)
        position_y=(screen_height//2)-(window_height//2)
        self.root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")
        self.root.rowconfigure(0,weight=1)
        self.root.columnconfigure(0,weight=1)


        self.logo = PhotoImage(file="images/logo.png")
        self.bg = ImageTk.PhotoImage(file="images/n.jpg")
        self.bg_images = Label(self.root, image=self.bg)
        self.bg_images.place(x=0, y=0, relwidth=1, relheight=1)

        Frame_login = Frame(self.root, bg="white")
        Frame_login.place(x=150, y=150, height=340, width=500)

        title = Label(Frame_login, text="Login", font=('Elephant', 30, "bold"), bg="white", fg="black")
        title.place(x=90, y=30)

        lbl_matricule = Label(Frame_login, text='Matricule', font=("Goudy old style", 15, "bold"), bg="white", fg="#767171")
        lbl_matricule.place(x=90, y=110)
        self.text_matricule = Entry(Frame_login, font=("times new roman", 15), bg="#ECECEC")
        self.text_matricule.place(x=90, y=140, width=350, height=35)

        lbl_password = Label(Frame_login, text='Password', font=("Goudy old style", 15, "bold"), bg="white", fg="#767171")
        lbl_password.place(x=90, y=180)
        self.text_password = Entry(Frame_login, show="*", font=("times new roman", 15), bg="#ECECEC")
        self.text_password.place(x=90, y=210, width=350, height=35)

        login_btn = Button(self.root, command=self.login, text="Log In", fg="white", bg="#00B0F0", activebackground="#00B0F0", activeforeground="white", cursor="hand2", font=("Arial Rounded MT Bold", 20))
        login_btn.place(x=250, y=470, width=250, height=35)

        hr = Label(Frame_login, bg="lightgray")
        hr.place(x=89, y=270, width=350, height=2)

        matri = None
        # global_matricule=None
    
    def login(self):
        global global_matricule
        matri = self.text_matricule.get()
        global_matricule = matri 
        if self.text_matricule.get() == "" or self.text_password.get() == "":
            messagebox.showerror("Erreur", "You have not completed the required fields", parent=self.root)
        else:
            try:
                connect = pyodbc.connect(
                                    'DRIVER={ODBC Driver 17 for SQL Server};'
                                    'SERVER=130.171.181.100;'
                                    'DATABASE=Précontrole;'
                                    'UID=Cutting;'
                                    'PWD=Cutting')
                cur = connect.cursor()
                cur.execute("select * from LP_Spool_demand_users where password=? and matricule=?", (self.text_password.get(), self.text_matricule.get()))
                row = cur.fetchone()
                if row is None:
                    messagebox.showerror("Erreur", "Matricule ou mot de passe invalide", parent=self.root)
                else:
                    self.root.destroy()
                    FormMysql.matri = matri  # Stocker le matricule dans la variable de classe
                    if row[3] == "HG":
                        # add_commande(global_matricule)
                        # call(["python", "add_commande.py"])
                        root.withdraw() #cache
                        # open_new_
                        add_commande
                    elif row[3] == "logistique":
                        # import listcommande
                        # root.withdraw()
                        call(["python", "listcommande.py"])
                    elif row[3] == "admin":
                        import archife
                    connect.close()
            except Exception as es:
                messagebox.showerror("Erreur", f"Erreur de connexion:{str(es)}", parent=self.root)

root = Tk()
obj = FormMysql(root)
root.mainloop()
