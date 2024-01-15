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


def RemplissageBible(res, cm, td, tp):
    # Créer une nouvelle feuille pour chaque matière
    nouvelle_feuille = classeur_existant.create_sheet(res)

    # Copier le contenu de la feuille existante dans la nouvelle feuille
    nom_feuille_base = "Sheet"  # ou spécifiez le nom de la feuille existante
    feuille_base = classeur_existant[nom_feuille_base]
    for ligne in feuille_base.iter_rows(min_row=1, values_only=True):
        nouvelle_feuille.append(ligne)

    cellule_selectionnee = nouvelle_feuille['C1']
    cellule_selectionnee.value = res

    cellule_selectionnee = nouvelle_feuille['B4']
    cellule_selectionnee.value = cm

    cellule_selectionnee = nouvelle_feuille['C4']
    cellule_selectionnee.value = td

    cellule_selectionnee = nouvelle_feuille['D4']
    cellule_selectionnee.value = tp



    ligne_insertion = 7
    nouvelle_feuille.insert_rows(ligne_insertion)
    nouvelle_feuille.cell(row=ligne_insertion, column=1, value='prof1')

def RemplissageProfs(feuille, resp):
    feuilleRes = classeur_existant[feuille]

    cellule_selectionnee = feuilleRes['H1']
    cellule_selectionnee.value = resp

#récupération des noms de chaque ressource de la base de donnée
cursor.execute("SELECT libelle_simple FROM BIBLE")
liste_libelle = cursor.fetchall()

#récupération des horaires
cursor.execute("SELECT total_cm FROM BIBLE")
liste_cm = cursor.fetchall()
cursor.execute("SELECT total_td FROM BIBLE")
liste_td = cursor.fetchall()
cursor.execute("SELECT total_tp FROM BIBLE")
liste_tp = cursor.fetchall()

# récupération des noms de profs
cursor.execute("SELECT AlerteProf, Intervenant FROM DONNEEPROF")
liste_profs = cursor.fetchall()

for i in range(len(liste_libelle)):
    RemplissageBible(liste_libelle[i][0], liste_cm[i][0], liste_td[i][0], liste_tp[i][0])
    for j in range(len(liste_profs)):
        RemplissageProfs(liste_libelle[i][0], liste_profs[j][1])









# Enregistrer le classeur modifié dans un nouveau fichier
classeur_existant.save('Ressources.xlsx')