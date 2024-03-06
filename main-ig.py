# Le lancement de tt les fichiers " fonctionne ", la selection oui mais ça lance pas encore
# faudra modif les fichiers pour en faire des print  -> genRapport / ressources


import tkinter as tk
from tkinter import filedialog
from tkinter import filedialog, messagebox
import sys
import os
import sys

from generationRapport import RapportErreurGenerator
from generationRessources import GenerationRessources
from bddInsertion import BddInsertion
from ExtractionPlanning import *

def executer_code():
    for widget in page_principale.winfo_children():
        widget.destroy()

    sortie_texte = tk.Text(page_principale)
    sortie_texte.pack(fill=tk.BOTH, expand=True)

    sys.stdout = TextRedirector(sortie_texte, "stdout")

    # Exécuter le contenu de bdd_file.py
    print("Base de donnée : ")
    print("-----------------")
    try:
        exec(open("bddInsertion.py").read())
        print("Succès")
    except Exception as e:
        print(f"Erreur lors de l'exécution du fichier : {e}")
    # Exécuter le contenu de Rapport.py
    print("-----------------")
    print("Génération Rapport Erreur : ")
    print("-----------------")
    try:
        exec(open("generationRapport.py").read())
        print("Succès ")
        print("-----------------")

    except Exception as e:
        print(f"Erreur lors de l'exécution du fichier : {e}")
        print("-----------------")
 
    print("Contenu du fichier rapportErreurs.txt : ")
    print("-----------------")
    try:
        with open("rapportErreurs.txt", "r") as rapport_file:
            contenu_rapport = rapport_file.read()
            print(contenu_rapport)
    except Exception as e:
        print(f"Erreur lors de la lecture du fichier : {e}")
     # Exécuter le contenu de Ressources.py
    print("-----------------")
    print("Génération Ressources")
    print("-----------------")
    try:
        exec(open("generationRessources.py"))
        with open("rapportErreur.txt", "r") as rapport_file:
            contenu_rapport = rapport_file.read()
            print(contenu_rapport)
    except Exception as e:
        print(f"Erreur lors de l'exécution du fichier : {e}")



class TextRedirector(object):
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, str):
        self.widget.insert(tk.END, str, (self.tag,))

    def flush(self):
        pass

def afficher_selection_fichiers():
    # Effacer le contenu de la page principale
    for widget in page_principale.winfo_children():
        widget.destroy()

    # Créer un bouton pour revenir à la première page
    bouton_retour = tk.Button(page_principale, text="Retour", command=afficher_page_principale)
    bouton_retour.pack()

    # Bouton pour sélectionner les fichiers
    bouton_selectionner = tk.Button(page_principale, text="Sélectionner des fichiers", command=selectionner_fichiers)
    bouton_selectionner.pack()

def executer_fichiers_selectionnes():
    # Effacer le contenu de la page principale
    for widget in page_principale.winfo_children():
        widget.destroy()

    # Afficher le résultat dans le widget Text
    sortie_texte = tk.Text(page_principale)
    sortie_texte.pack(fill=tk.BOTH, expand=True)

    sys.stdout = TextRedirector(sortie_texte, "stdout")

    # Parcourir les fichiers sélectionnés
    for fichier, coche in zip(fichiers_disponibles, fichiers_coche):
        if coche.get():
            try:
                if fichier == "bddInsertion.py":
                    print("Base de donnée : ")
                    print("-----------------")
                    try:
                        exec(open("bddInsertion.py").read())
                        print("Succès")
                    except Exception as e:
                        print(f"Erreur lors de l'exécution du fichier : {e}")
                elif fichier == "generationRapport.py":
                    print("-----------------")
                    print("Génération Rapport Erreur : ")
                    print("-----------------")
                    try:
                        exec(open("generationRapport.py").read())
                        print("Succès ")
                        print("-----------------")

                    except Exception as e:
                        print(f"Erreur lors de l'exécution du fichier : {e}")
                        print("-----------------")

                    print("Contenu du fichier rapportErreurs.txt : ")
                    print("-----------------")
                    try:
                        with open("rapportErreurs.txt", "r") as rapport_file:
                            contenu_rapport = rapport_file.read()
                        print(contenu_rapport)
                    except Exception as e:
                        print(f"Erreur lors de la lecture du fichier : {e}")
                elif fichier == "generationRessources.py":
                    print("-----------------")
                    print("Génération Ressources")
                    print("-----------------")
                    try:
                        exec(open("generationRessources.py"))
                        with open("rapportErreur.txt", "r") as rapport_file:
                            contenu_rapport = rapport_file.read()
                            print(contenu_rapport)
                    except Exception as e:
                        print(f"Erreur lors de l'exécution du fichier : {e}")
            except Exception as e:
                print(f"Erreur lors de l'exécution du fichier {fichier} : {e}")

    # Ajouter un bouton "Valider" pour afficher le résultat
    bouton_valider = tk.Button(page_principale, text="Valider", command=lambda: afficher_resultat(sortie_texte))
    bouton_valider.pack()

def afficher_resultat(sortie_texte):
    # Afficher le résultat dans le widget Text
    contenu_resultat = sortie_texte.get("1.0", tk.END)
    print("Résultat : ")
    print("-----------------")
    print(contenu_resultat)

def selectionner_fichiers():
    global fichiers_selectionnes
    fichiers_selectionnes = []

    # Création de la fenêtre de sélection de fichiers
    fenetre_selection = tk.Toplevel()
    fenetre_selection.title("Sélectionner les fichiers à lancer")

    # Création des cases à cocher pour les fichiers
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

    # Bouton pour valider la sélection
    bouton_valider = tk.Button(fenetre_selection, text="Valider", command=fenetre_selection.destroy)
    bouton_valider.pack()

    fenetre_selection.mainloop()

def afficher_page_principale():
    for widget in page_principale.winfo_children():
        widget.destroy()
    
    # Créer les boutons de la page principale
    bouton_lancer_tous = tk.Button(page_principale, text="Lancer tous les fichiers", command=executer_code)
    bouton_lancer_tous.pack()

    bouton_lancer_selection = tk.Button(page_principale, text="Lancer les fichiers sélectionnés", command=afficher_selection_fichiers)
    bouton_lancer_selection.pack()

# Création de la fenêtre principale
fenetre = tk.Tk()
fenetre.title("Gestion des fichiers")
fenetre.geometry("400x300")

# Création de la page principale
page_principale = tk.Frame(fenetre)
page_principale.pack(fill=tk.BOTH, expand=True)  

afficher_page_principale()

if __name__ == "__main__":
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        os.chdir(sys._MEIPASS)
    fenetre.mainloop()
