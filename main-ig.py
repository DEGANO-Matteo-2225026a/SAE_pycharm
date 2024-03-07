import tkinter as tk
from tkinter import filedialog, messagebox
import sys
import os

from generationRapport import RapportErreurGenerator
from generationRessources import GenerationRessources
from bddInsertion import BddInsertion
from ExtractionPlanning import *

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
        with open("rapportErreur.txt", "r") as rapport_file:
            contenu_rapport = rapport_file.read()
            print(contenu_rapport)
    except Exception as e:
        print(f"Erreur lors de l'exécution du fichier generationRessources.py : {e}")

def afficher_selection_fichiers():
    for widget in page_principale.winfo_children():
        widget.destroy()

    bouton_retour = tk.Button(page_principale, text="Retour au Menu Principal", command=afficher_page_principale, bg="blue", fg="white", font=("Helvetica", 12, "bold"), padx=20, pady=10)
    bouton_retour.pack(pady=20)

    bouton_selectionner = tk.Button(page_principale, text="Sélectionner des fichiers", command=selectionner_fichiers, bg="green", fg="white", font=("Helvetica", 12, "bold"), padx=20, pady=10)
    bouton_selectionner.pack(pady=20)

def executer_fichiers_selectionnes():
    for widget in page_principale.winfo_children():
        widget.destroy()

    sortie_texte = tk.Text(page_principale, bg="white", fg="black", font=("Helvetica", 10))
    sortie_texte.pack(fill=tk.BOTH, expand=True)

    sys.stdout = TextRedirector(sortie_texte, "stdout")

    for fichier, coche in zip(fichiers_disponibles, fichiers_coche):
        if coche.get():
            try:
                exec(open(fichier).read())
                print(f"Succès : {fichier}")
            except Exception as e:
                print(f"Erreur lors de l'exécution du fichier {fichier} : {e}")

    bouton_valider = tk.Button(page_principale, text="Valider", command=lambda: afficher_resultat(sortie_texte), bg="green", fg="white", font=("Helvetica", 12, "bold"), padx=20, pady=10)
    bouton_valider.pack(pady=20)

def afficher_resultat(sortie_texte):
    contenu_resultat = sortie_texte.get("1.0", tk.END)
    print("Résultat : ")
    print("-----------------")  
    print(contenu_resultat)

def selectionner_fichiers():
    global fichiers_selectionnes
    fichiers_selectionnes = []

    fenetre_selection = tk.Toplevel()
    fenetre_selection.title("Sélectionner les fichiers à lancer")

    check_bdd = tk.Checkbutton(fenetre_selection, text="bddInsertion.py", command=lambda: ajouter_fichier("bddInsertion.py"))
    check_bdd.pack()

    check_rapport = tk.Checkbutton(fenetre_selection, text="generationRapport.py", command=lambda: ajouter_fichier("generationRapport.py"))
    check_rapport.pack()

    check_ressources = tk.Checkbutton(fenetre_selection, text="generationRessources.py", command=lambda: ajouter_fichier("generationRessources.py"))
    check_ressources.pack()

    def ajouter_fichier(nom_fichier):
        if nom_fichier in fichiers_selectionnes:
            fichiers_selectionnes.remove(nom_fichier)
        else:
            fichiers_selectionnes.append(nom_fichier)

    bouton_valider = tk.Button(fenetre_selection, text="Valider", command=fenetre_selection.destroy, bg="blue", fg="white", font=("Helvetica", 12, "bold"), padx=20, pady=10)
    bouton_valider.pack(pady=20)

    fenetre_selection.mainloop()

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
