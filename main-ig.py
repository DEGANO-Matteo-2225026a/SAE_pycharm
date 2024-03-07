import tkinter as tk
from tkinter import filedialog, messagebox
import sys
import os
import shutil

from generationRapport import RapportErreurGenerator
from generationRessources import GenerationRessources
from generationFicheProf import GenerationFicheProf
from bddInsertion import BddInsertion
from ExtractionPlanning import *
from io import StringIO

class TextRedirector(object):
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, str):
        self.widget.insert(tk.END, str, (self.tag,))

    def flush(self):
        pass

def executer_code():
    for widget in page_principale.winfo_children():
        widget.destroy()

    bouton_retour = tk.Button(page_principale, text="Retour au Menu Principal", command=afficher_page_principale, bg="blue", fg="white", font=("Helvetica", 12, "bold"), padx=20, pady=10)
    bouton_retour.pack(pady=20)

    sortie_texte = tk.Text(page_principale, bg="white", fg="black", font=("Helvetica", 10))
    sortie_texte.pack(fill=tk.BOTH, expand=True)

    sys.stdout = TextRedirector(sortie_texte, "stdout")

    # Exécuter le contenu des fichiers sélectionnés
    try:
        exec(open("bddInsertion.py").read())
        print("Succès : Base de données")
    except Exception as e:
        print(f"Erreur lors de l'exécution du fichier bddInsertion.py : {e}")

    print("-----------------")
    print("Génération Rapport Erreur : ")
    print("-----------------")
    try:
        exec(open("generationRapport.py").read())
        print("Succès : Génération Rapport Erreur")
        print("-----------------")
    except Exception as e:
        print(f"Erreur lors de l'exécution du fichier generationRapport.py : {e}")

    print("Contenu du fichier rapportErreurs.txt : ")
    print("-----------------")
    try:
        with open("rapportErreurs.txt", "r") as rapport_file:
            contenu_rapport = rapport_file.read()
            print(contenu_rapport)
    except Exception as e:
        print(f"Erreur lors de la lecture du fichier rapportErreurs.txt : {e}")

    print("-----------------")
    print("Génération Ressources")
    print("-----------------")
    try:
        exec(open("generationRessources.py").read())
        print("Succès : Génération Ressource")
        print("-----------------")
    except Exception as e:
        print(f"Erreur lors de l'exécution du fichier generationRessources.py : {e}")
    try:
        exec(open("generationFicheProf.py").read())
        print("Succès : Génération FicheProf")
        print("-----------------")
    except Exception as e:
        print(f"Erreur lors de l'exécution du fichier generationFicheProf.py : {e}")
        
    bouton_telecharger = tk.Button(page_principale, text="Télécharger les fichiers", command=telecharger_fichiers, bg="green", fg="white", font=("Helvetica", 12, "bold"), padx=20, pady=10)
    bouton_telecharger.pack(pady=20)

fichiers_selectionnes = []
fichiers_coche = []

def telecharger_fichiers():
    try:
        destination_folder = filedialog.askdirectory(title="Sélectionner le dossier de destination")

        if destination_folder:
            # Copie des fichiers vers le dossier de destination
            for fichier in fichiers_selectionnes:
                shutil.copy(fichier, destination_folder)

            messagebox.showinfo("Téléchargement réussi", "Les fichiers ont été téléchargés avec succès.")
    except Exception as e:
        messagebox.showerror("Erreur de téléchargement", f"Une erreur est survenue lors du téléchargement des fichiers : {e}")

def afficher_selection_fichiers():
    for widget in page_principale.winfo_children():
        widget.destroy()

    bouton_retour = tk.Button(page_principale, text="Retour au Menu Principal", command=afficher_page_principale, bg="blue", fg="white", font=("Helvetica", 12, "bold"), padx=20, pady=10)
    bouton_retour.pack(pady=20)

    bouton_selectionner = tk.Button(page_principale, text="Sélectionner des fichiers", command=selectionner_fichiers, bg="green", fg="white", font=("Helvetica", 12, "bold"), padx=20, pady=10)
    bouton_selectionner.pack(pady=20)

    bouton_executer = tk.Button(page_principale, text="Exécuter les fichiers sélectionnés", command=executer_fichiers_selectionnes, bg="orange", fg="white", font=("Helvetica", 12, "bold"), padx=20, pady=10)
    bouton_executer.pack(pady=20)

def selectionner_fichiers():
    global fichiers_selectionnes, fichiers_coche
    fichiers_selectionnes = []
    fichiers_coche = []

    fenetre_selection = tk.Toplevel()
    fenetre_selection.title("Sélectionner les fichiers à lancer")

    check_bdd = tk.Checkbutton(fenetre_selection, text="bddInsertion.py")
    fichiers_coche.append(tk.BooleanVar())
    check_bdd.config(variable=fichiers_coche[-1])
    check_bdd.pack()

    check_rapport = tk.Checkbutton(fenetre_selection, text="generationRapport.py")
    fichiers_coche.append(tk.BooleanVar())
    check_rapport.config(variable=fichiers_coche[-1])
    check_rapport.pack()

    check_ressources = tk.Checkbutton(fenetre_selection, text="generationRessources.py")
    fichiers_coche.append(tk.BooleanVar())
    check_ressources.config(variable=fichiers_coche[-1])
    check_ressources.pack()

    check_ressources = tk.Checkbutton(fenetre_selection, text="generationFicheProf.py")
    fichiers_coche.append(tk.BooleanVar())
    check_ressources.config(variable=fichiers_coche[-1])
    check_ressources.pack()

    def ajouter_fichier(nom_fichier):
        if nom_fichier in fichiers_selectionnes:
            fichiers_selectionnes.remove(nom_fichier)
        else:
            fichiers_selectionnes.append(nom_fichier)

    bouton_valider = tk.Button(fenetre_selection, text="Valider", command=lambda: [ajouter_fichier("bddInsertion.py" if fichiers_coche[0].get() else ""), ajouter_fichier("generationRapport.py" if fichiers_coche[1].get() else ""), ajouter_fichier("generationRessources.py" if fichiers_coche[2].get() else ""), ajouter_fichier("generationFicheProf.py" if fichiers_coche[3].get() else ""), fenetre_selection.destroy()], bg="blue", fg="white", font=("Helvetica", 12, "bold"), padx=20, pady=10)
    bouton_valider.pack(pady=20)

    fenetre_selection.mainloop()

def executer_fichiers_selectionnes():
    for widget in page_principale.winfo_children():
        widget.destroy()

    sortie_texte = tk.Text(page_principale, bg="white", fg="black", font=("Helvetica", 10))
    sortie_texte.pack(fill=tk.BOTH, expand=True)

    # Redirection de la sortie standard vers un objet de capture
    capture_sortie = StringIO()
    sys.stdout = capture_sortie

    for fichier in fichiers_selectionnes:
        if fichier == "bddInsertion.py":
            print("-----------------")
            print("Lancement de bddInsertion.py")
            print("-----------------")
            try:
                exec(open(fichier).read())
                print(f"Succès : {fichier}")
            except Exception as e:
                print(f"Erreur lors de l'exécution du fichier {fichier} : {e}")

        elif fichier == "generationRapport.py":
            print("-----------------")
            print("Lancement de generationRapport.py")
            print("-----------------")
            try:
                exec(open(fichier).read())
                print(f"Succès : {fichier}")
                print("-----------------")
            except Exception as e:
                print(f"Erreur lors de l'exécution du fichier {fichier} : {e}")
            print("Contenu du fichier rapportErreurs.txt : ")
            print("-----------------")
            try:
                with open("rapportErreurs.txt", "r") as rapport_file:
                    contenu_rapport = rapport_file.read()
                    print(contenu_rapport)
            except Exception as e:
                print(f"Erreur lors de la lecture du fichier rapportErreurs.txt : {e}")

        elif fichier == "generationRessources.py":
            print("-----------------")
            print("Lancement de generationRessources.py")
            print("-----------------")
            try:
                exec(open(fichier).read())
                print(f"Succès : {fichier}")
                print("-----------------")
            except Exception as e:
                print(f"Erreur lors de l'exécution du fichier {fichier} : {e}")

        elif fichier == "generationFicheProf.py":
            print("-----------------")
            print("Lancement de generationFicheProf.py")
            print("-----------------")
            try:
                exec(open(fichier).read())
                print(f"Succès : {fichier}")
                print("-----------------")
            except Exception as e:
                print(f"Erreur lors de l'exécution du fichier {fichier} : {e}")

    # Restaurer la sortie standard
    sys.stdout = sys.__stdout__

    # Récupérer le contenu de la sortie capturée
    contenu_resultat = capture_sortie.getvalue()

    # Afficher le contenu dans le widget texte en utilisant la fonction afficher_resultat
    afficher_resultat(sortie_texte, contenu_resultat)

    # Ajouter un bouton pour revenir en arrière
    bouton_retour = tk.Button(page_principale, text="Retour", command=afficher_selection_fichiers, bg="blue", fg="white", font=("Helvetica", 12, "bold"), padx=20, pady=10)
    bouton_retour.pack(pady=20)

def afficher_resultat(sortie_texte, contenu_resultat):
    sortie_texte.insert(tk.END, contenu_resultat)
    fenetre.update()
def telecharger_fichiers():
    try:
        destination_folder = filedialog.askdirectory(title="Sélectionner le dossier de destination")

        if destination_folder:
            # Copie des fichiers FicheProf.xlsx et Ressources.xlsx vers le dossier de destination
            shutil.copy("Excels ressources/FicheProf.xlsx", destination_folder)
            shutil.copy("Excels ressources/Ressources.xlsx", destination_folder)

            messagebox.showinfo("Téléchargement réussi", "Les fichiers ont été téléchargés avec succès.")
    except Exception as e:
        messagebox.showerror("Erreur de téléchargement", f"Une erreur est survenue lors du téléchargement des fichiers : {e}")

def afficher_page_principale():
    for widget in page_principale.winfo_children():
        widget.destroy()
    
    bouton_lancer_tous = tk.Button(page_principale, text="Lancer tous les fichiers", command=executer_code, bg="green", fg="white", font=("Helvetica", 12, "bold"), padx=20, pady=10)
    bouton_lancer_tous.pack(pady=20)

    bouton_lancer_selection = tk.Button(page_principale, text="Lancer les fichiers sélectionnés", command=afficher_selection_fichiers, bg="blue", fg="white", font=("Helvetica", 12, "bold"), padx=20, pady=10)
    bouton_lancer_selection.pack(pady=20)

fenetre = tk.Tk()
fenetre.title("Gestion des fichiers")
fenetre.geometry("600x400")
fenetre.configure(bg="lightgrey")

page_principale = tk.Frame(fenetre, bg="lightgrey")
page_principale.pack(fill=tk.BOTH, expand=True)  

afficher_page_principale()

if __name__ == "__main__":
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        os.chdir(sys._MEIPASS)
    fenetre.mainloop()
