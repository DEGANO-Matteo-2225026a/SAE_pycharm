import openpyxl
from openpyxl import load_workbook
from pathlib import Path
import sqlite3


#récupérer un fichier excel de ressource de base créé par un autre code
base = Path('Ressources.xlsx')
#si le fichier ecxel résultant de ce code existe déjà on l'utilise directement
if base.is_file():
    # Charger le classeur Excel existant
    classeur_existant = load_workbook('Ressources.xlsx')
# si il n'existe pas on éxécute le code qui va créer le fichier de base
else:
    exec(open('baseRessources.py').read())
    classeur_existant = load_workbook('Ressources.xlsx')

def Remplissage():
    # Sélectionner la feuille existante (ou créer une nouvelle feuille si elle n'existe pas)
    #nom_feuille = 'Sheet'  # Remplacez par le nom de votre feuille
    #feuille = classeur_existant[
    #    nom_feuille] if nom_feuille in classeur_existant.sheetnames else classeur_existant.create_sheet(nom_feuille)
    nouvelle_feuille = classeur_existant.create_sheet("R...")

    # Copier le contenu de la feuille existante dans la nouvelle feuille
    nom_feuille_base = "Sheet"  # ou spécifiez le nom de la feuille existante
    feuille_base = classeur_existant[nom_feuille_base]
    for ligne in feuille_base.iter_rows(min_row=1, values_only=True):
        nouvelle_feuille.append(ligne)

    cellule_selectionnee = feuille_base['C1']
    cellule_selectionnee.value = 'Nomres'

    cellule_selectionnee = nouvelle_feuille['H1']
    cellule_selectionnee.value = 'Nomresp'

    cellule_selectionnee = nouvelle_feuille['B4']
    cellule_selectionnee.value = '...h'

    cellule_selectionnee = nouvelle_feuille['C4']
    cellule_selectionnee.value = '...h'

    cellule_selectionnee = nouvelle_feuille['D4']
    cellule_selectionnee.value = '...h'

    ligne_insertion = 7
    nouvelle_feuille.insert_rows(ligne_insertion)
    nouvelle_feuille.cell(row=ligne_insertion, column=1, value='prof1')

    ligne_insertion = 8
    nouvelle_feuille.insert_rows(ligne_insertion)
    nouvelle_feuille.cell(row=ligne_insertion, column=1, value='prof2')

    ligne_insertion = 9
    nouvelle_feuille.insert_rows(ligne_insertion)
    nouvelle_feuille.cell(row=ligne_insertion, column=1, value='prof3')

    # Enregistrer le classeur modifié dans un nouveau fichier
    classeur_existant.save('Ressources.xlsx')

Remplissage()