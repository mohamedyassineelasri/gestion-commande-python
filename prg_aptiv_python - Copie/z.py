import tkinter as tk

def disable_close_button():
    pass  # Cette fonction ne fait rien lorsqu'elle est appelée

# Création de la fenêtre principale
root = tk.Tk()
root.title("Application sans bouton de fermeture")

# Désactivation du bouton de fermeture
root.protocol("WM_DELETE_WINDOW", disable_close_button)

# Configuration du style de la fenêtre
root.configure(bg='gray')
root.geometry('300x200')

# Ajout d'une étiquette comme contenu de la fenêtre
label = tk.Label(root, text="Contenu de la fenêtre", bg='gray', fg='white')
label.pack(expand=True)

# Démarrage de l'application Tkinter
root.mainloop()
