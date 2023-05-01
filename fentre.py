import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import paretoù
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class ImportFichier(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Import de fichier Excel")
        self.geometry("500x300")

        # Création du label "Choisir le fichier"
        label_fichier = tk.Label(self, text="Choisir le fichier")
        label_fichier.pack(pady=10)
        fichier_path= None
        # Création du bouton "Importer un fichier"
        bouton_importer_fichier = tk.Button(self, text="Importer un fichier", command=self.importer_fichier)
        bouton_importer_fichier.pack()

        # Initialisation des variables de contrôle pour les listes déroulantes
        self.liste_feuilles = []
        self.liste_colonnes = []
        self.liste_deroulante_feuilles_variable = tk.StringVar(self)
        self.liste_deroulante_feuilles_variable.set("")
        self.liste_deroulante_colonnes_variable = tk.StringVar(self)
        self.liste_deroulante_colonnes_variable.set("")

    def importer_fichier(self):
        # Ouvrir la boîte de dialogue pour choisir un fichier Excel
        fichier_path = filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xlsx;*.xls")])
        if fichier_path:
            try:
                # Charger le fichier Excel en utilisant pandas
                excel_data = pd.read_excel(fichier_path, sheet_name=None)

                # Mettre à jour la liste déroulante des feuilles
                self.liste_feuilles = list(excel_data.keys())
                self.liste_deroulante_feuilles_variable.set(self.liste_feuilles[0])

                # Créer la liste déroulante des feuilles
                label_feuilles = tk.Label(self, text="Choisir la feuille")
                label_feuilles.pack(pady=10)
                self.liste_deroulante_feuilles = tk.OptionMenu(self, self.liste_deroulante_feuilles_variable, *self.liste_feuilles)
                self.liste_deroulante_feuilles.pack()

                # Mettre à jour la liste déroulante des colonnes
                feuille_selectionnee = self.liste_deroulante_feuilles_variable.get()
                self.liste_colonnes = list(excel_data[feuille_selectionnee].columns)
                self.liste_deroulante_colonnes_variable.set(self.liste_colonnes[0])

                # Créer la liste déroulante des colonnes
                label_colonnes = tk.Label(self, text="Choisir la colonne")
                label_colonnes.pack(pady=10)
                self.liste_deroulante_colonnes = tk.OptionMenu(self, self.liste_deroulante_colonnes_variable, *self.liste_colonnes)
                self.liste_deroulante_colonnes.pack()

                # Créer le bouton "Submit"
                bouton_submit = tk.Button(self, text="Submit", command=self.afficher_informations)
                bouton_submit.pack()
                self.fichier_path=fichier_path

            except Exception as e:
                tk.messagebox.showerror("Erreur", f"Erreur lors de la lecture du fichier Excel : {str(e)}")

    def afficher_informations(self):
        # Récupérer les informations sélectionnées dans les listes déroulantes et affichage
        feuille = self.liste_deroulante_feuilles_variable.get()
        colonne = self.liste_deroulante_colonnes_variable.get()
        fichier_path = self.fichier_path
        paretoù.Pareto(fichier_path, feuille, colonne).create_pareto()
     

ImportFichier().mainloop()