# Le lancement de tt les fichiers " fonctionne ", la selection oui mais ça lance pas encore
# faudra modif les fichiers pour en faire des print  -> genRapport / ressources


import tkinter as tk
from tkinter import filedialog
import sys
import os

# Fonction pour exécuter le contenu de a.py
def executer_code():
    # Effacer le contenu actuel de la page principale
    for widget in page_principale.winfo_children():
        widget.destroy()

    # Créer un widget Text pour afficher la sortie
    sortie_texte = tk.Text(page_principale)
    sortie_texte.pack(fill=tk.BOTH, expand=True)

    # Rediriger stdout vers le widget Text
    sys.stdout = TextRedirector(sortie_texte, "stdout")

    # Exécuter le contenu de bdd_file.py
    print("Base de donnée : ")
    print("-----------------")
    try:
        exec(open("BDD_file.py").read())
        print("Succès")
    except Exception as e:
        print(f"Erreur lors de l'exécution du fichier : {e}")
    # Exécuter le contenu de Rapport.py
    print("-----------------")
    print("Génération Rapport Erreur : ")
    print("-----------------")
    try:
        exec(open("generationRapport.py").read())
    except Exception as e:
        print(f"Erreur lors de l'exécution du fichier : {e}")
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



# Classe pour rediriger stdout vers un widget Text
class TextRedirector(object):
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, str):
        self.widget.insert(tk.END, str, (self.tag,))

    def flush(self):
        pass

# Fonction pour afficher la sélection des fichiers
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
    for fichier in fichiers_selectionnes:
        try:
            contenu_resultat = ''
            with open(fichier, "r") as file:
                contenu_resultat = file.read()
            # Afficher le nom du fichier
            label_nom_fichier = tk.Label(page_principale, text=f"Fichier : {fichier}")
            label_nom_fichier.pack()
            # Afficher le contenu du fichier
            label_contenu_resultat = tk.Label(page_principale, text=contenu_resultat)
            label_contenu_resultat.pack()
            print(f"Succès : {fichier}")
        except Exception as e:
            print(f"Erreur lors de l'exécution du fichier {fichier} : {e}")
# Fonction pour sélectionner les fichiers
def selectionner_fichiers():
    global fichiers_selectionnes
    fichiers = filedialog.askopenfilenames()
    fichiers_selectionnes = list(fichiers)  # Stocker les chemins des fichiers sélectionnés
    executer_fichiers_selectionnes()  # Exécuter les fichiers sélectionnés après la sélection

# Fonction pour afficher la page principale
def afficher_page_principale():
    # Effacer le contenu actuel de la page principale
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
fenetre.geometry("400x300")  # Définir la taille de la fenêtre

# Création de la page principale
page_principale = tk.Frame(fenetre)
page_principale.pack(fill=tk.BOTH, expand=True)  # Remplir la fenêtre principale

# Afficher la page principale au démarrage
afficher_page_principale()

# Si le script est exécuté directement en Python, exécuter la boucle principale d'événements
if __name__ == "__main__":
    # Si un fichier exécutable est créé à partir de ce script
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        # Définir le répertoire de travail sur le répertoire contenant le fichier exécutable
        os.chdir(sys._MEIPASS)
    fenetre.mainloop()
