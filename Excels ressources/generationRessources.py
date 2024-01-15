import openpyxl
from openpyxl import load_workbook
from pathlib import Path
import sqlite3

# connection à la base de donnée
db_connection = sqlite3.connect("../SAE.db")
cursor = db_connection.cursor()

#récupérer un fichier excel de ressource de base créé par un autre code
base = Path('Ressources.xlsx')
#si le fichier excel résultant de ce code existe déjà on l'utilise directement
if base.is_file():
    # Charger le classeur Excel existant
    classeur_existant = load_workbook('Ressources.xlsx')
# si il n'existe pas on éxécute le code qui va créer le fichier de base
else:
    exec(open('baseRessources.py').read())
    classeur_existant = load_workbook('Ressources.xlsx')


def Remplissage(res):
    # Créer une nouvelle feuille pour chaque matière
    nouvelle_feuille = classeur_existant.create_sheet(res)

    # Copier le contenu de la feuille existante dans la nouvelle feuille
    nom_feuille_base = "Sheet"  # ou spécifiez le nom de la feuille existante
    feuille_base = classeur_existant[nom_feuille_base]
    for ligne in feuille_base.iter_rows(min_row=1, values_only=True):
        nouvelle_feuille.append(ligne)

    cellule_selectionnee = nouvelle_feuille['C1']
    cellule_selectionnee.value = res

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

#récupération des noms de chaque ressource de la base de donnée
cursor.execute("SELECT libelle_simple FROM BIBLE")
liste_libelle = cursor.fetchall()
for i in range(len(liste_libelle)):
    Remplissage(liste_libelle[i][0])