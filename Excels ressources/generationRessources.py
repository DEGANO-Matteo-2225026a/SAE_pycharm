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


def Remplissage(res, cmtot, tdtot, tptot, resp):
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

    cellule_selectionnee = nouvelle_feuille['B4']
    cellule_selectionnee.value = cmtot

    cellule_selectionnee = nouvelle_feuille['C4']
    cellule_selectionnee.value = tdtot

    cellule_selectionnee = nouvelle_feuille['D4']
    cellule_selectionnee.value = tptot



    requete = "SELECT Intervenant, Acronyme, CM, TDNonD, TPD FROM DONNEEPROF WHERE SUBSTR(MatiereActuelle, 1, INSTR(MatiereActuelle, ' ') - 1) = ?"
    cursor.execute(requete, (res,))
    liste_inter = cursor.fetchall()


    # insertion des noms des profs dans la feuille
    for i in range(len(liste_inter)):

        inter = trouver_ligne(nouvelle_feuille, "Intervenants")

        ligne_insertion = 7 + i
        nouvelle_feuille.insert_rows(ligne_insertion)
        nouvelle_feuille.cell(row=ligne_insertion, column=1, value=liste_inter[i][0])
        nouvelle_feuille.cell(row=ligne_insertion, column=2, value=liste_inter[i][1])

        cms = trouver_ligne(nouvelle_feuille, "CM")

        ligne_insertion = cms + 1 + i
        nouvelle_feuille.insert_rows(ligne_insertion)
        nouvelle_feuille.cell(row=ligne_insertion, column=1, value=liste_inter[i][1])
        nouvelle_feuille.cell(row=ligne_insertion, column=2, value=liste_inter[i][2])

        tds = trouver_ligne(nouvelle_feuille, "TD")

        ligne_insertion = tds + 1 + i
        nouvelle_feuille.insert_rows(ligne_insertion)
        nouvelle_feuille.cell(row=ligne_insertion, column=1, value=liste_inter[i][1])
        nouvelle_feuille.cell(row=ligne_insertion, column=2, value=liste_inter[i][3])

        tps = trouver_ligne(nouvelle_feuille, "TP non dedoubles")

        ligne_insertion = tps + 1 + i
        nouvelle_feuille.insert_rows(ligne_insertion)
        nouvelle_feuille.cell(row=ligne_insertion, column=1, value=liste_inter[i][1])
        nouvelle_feuille.cell(row=ligne_insertion, column=2, value=liste_inter[i][4])

    cours = "SELECT Semaine, TypeCours, COUNT(*) AS NombreCours FROM PLANINFO WHERE Ressource = ? GROUP BY Semaine, TypeCours "
    cursor.execute(cours, (res,))
    liste_cours = cursor.fetchall()


    for i in range(len(liste_cours)):

        hcm = trouver_ligne(nouvelle_feuille, "CM heures")

        ligne_insertion = hcm + 1 + i
        nouvelle_feuille.insert_rows(ligne_insertion)
        nouvelle_feuille.cell(row=ligne_insertion, column=1, value=liste_cours[i][0])

        ligne_date = trouver_ligne(nouvelle_feuille, liste_cours[i][0])
        if liste_cours[i][1] == "Cours" :
            nouvelle_feuille.cell(row=ligne_date, column=2, value=(liste_cours[i][2]) * 1.5)
        elif liste_cours[i][1] == "TD":
            nouvelle_feuille.cell(row=ligne_date, column=3, value=(liste_cours[i][2] * 2))
        elif liste_cours[i][1] == "TP":
            nouvelle_feuille.cell(row=ligne_date, column=4, value=(liste_cours[i][2] * 2))
        else:
            nouvelle_feuille.cell(row=ligne_date, column=5, value=(liste_cours[i][2] * 2))












def trouver_ligne(feuille, contenu_cible):
    for row_number, row in enumerate(feuille.iter_rows(values_only=True), start=1):
        if contenu_cible in row:
            return row_number

    return None

requeteRemplissage = "SELECT Ressource, CM, TD, TP, Acronyme FROM PLANRESSOURCE WHERE Ressource NOT LIKE '%SAE%' GROUP BY Ressource ORDER BY Ressource "
cursor.execute(requeteRemplissage)
liste_res = cursor.fetchall()

for i in range(len(liste_res)):
     Remplissage(liste_res[i][0], liste_res[i][1], liste_res[i][2], liste_res[i][3], liste_res[i][4])


# Enregistrer le classeur modifié dans un nouveau fichier
classeur_existant.save('Ressources.xlsx')