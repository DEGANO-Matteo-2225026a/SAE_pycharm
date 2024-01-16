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


def Remplissage(res,resp):
    # Créer une nouvelle feuille pour chaque matière
    nouvelle_feuille = classeur_existant.create_sheet(res)

    # Copier le contenu de la feuille existante dans la nouvelle feuille
    nom_feuille_base = "Sheet"  # ou spécifiez le nom de la feuille existante
    feuille_base = classeur_existant[nom_feuille_base]
    for ligne in feuille_base.iter_rows(min_row=1, values_only=True):
        nouvelle_feuille.append(ligne)

    cellule_selectionnee = nouvelle_feuille['H1']
    cellule_selectionnee.value = resp

    cellule_selectionnee = nouvelle_feuille['C1']
    cellule_selectionnee.value = res

    requete = "SELECT Intervenant, Acronyme, CM, TDNonD, TPD FROM DONNEEPROF WHERE MatiereActuelle = ?"
    cursor.execute(requete, (res,))
    liste_inter = cursor.fetchall()

    # insertion des noms des profs dans la feuille
    for i in range(len(liste_inter)):
        ligne_insertion = 7 + i
        nouvelle_feuille.insert_rows(ligne_insertion)
        nouvelle_feuille.cell(row=ligne_insertion, column=1, value=liste_inter[i][0])
        nouvelle_feuille.cell(row=ligne_insertion, column=2, value=liste_inter[i][1])

        ligne_insertion = 11 + i
        nouvelle_feuille.insert_rows(ligne_insertion)
        nouvelle_feuille.cell(row=ligne_insertion, column=1, value=liste_inter[i][1])
        nouvelle_feuille.cell(row=ligne_insertion, column=2, value=liste_inter[i][2])

    cellule_selectionnee = nouvelle_feuille['B4']
    cellule_selectionnee.value = "heures"

    cellule_selectionnee = nouvelle_feuille['C4']
    cellule_selectionnee.value = "heures"

    cellule_selectionnee = nouvelle_feuille['D4']
    cellule_selectionnee.value = "heures"


cursor.execute("SELECT MatiereActuelle, Intervenant FROM DONNEEPROF GROUP BY MatiereActuelle")
liste_resp = cursor.fetchall()

for i in range(len(liste_resp)):
     Remplissage(liste_resp[i][0], liste_resp[i][1])





# Enregistrer le classeur modifié dans un nouveau fichier
classeur_existant.save('Ressources.xlsx')