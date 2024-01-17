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

# fonction qui crée une nouvelle feuille pour chaque ressource
def Remplissage(res, cmtot, tdtot, tptot, resp):
    # Créer une nouvelle feuille pour chaque matière
    nouvelle_feuille = classeur_existant.create_sheet(res)

    # Copier le contenu de la feuille existante dans la nouvelle feuille
    nom_feuille_base = "Sheet"  # ou spécifiez le nom de la feuille existante
    feuille_base = classeur_existant[nom_feuille_base]


    for ligne in feuille_base.iter_rows(min_row=1, values_only=True):
        nouvelle_feuille.append(ligne)

    #nom responsable de ressource
    cellule_selectionnee = nouvelle_feuille['H1']
    cellule_selectionnee.value = resp

    #nom ressource
    cellule_selectionnee = nouvelle_feuille['C1']
    cellule_selectionnee.value = res

    #volume horaire total
    cellule_selectionnee = nouvelle_feuille['B4']
    cellule_selectionnee.value = cmtot

    cellule_selectionnee = nouvelle_feuille['C4']
    cellule_selectionnee.value = tdtot

    cellule_selectionnee = nouvelle_feuille['D4']
    cellule_selectionnee.value = tptot

    #Récupération des donnees des intervenants
    requete = "SELECT Intervenant, Acronyme, CM, TD, TPD, TDNonD FROM DONNEEPROF WHERE SUBSTR(MatiereActuelle, 1, INSTR(MatiereActuelle, ' ') - 1) = ?"
    cursor.execute(requete, (res,))
    liste_inter = cursor.fetchall()


    # insertion données dans la feuille
    for i in range(len(liste_inter)):

        #nom et acronyme
        inter = trouver_ligne(nouvelle_feuille, "Intervenants")
        
        ligne_insertion = 7 + i
        nouvelle_feuille.insert_rows(ligne_insertion)
        nouvelle_feuille.cell(row=ligne_insertion, column=1, value=liste_inter[i][0])
        nouvelle_feuille.cell(row=ligne_insertion, column=2, value=liste_inter[i][1])

        # pourcentage du total de cm effectué par le professeur
        cms = trouver_ligne(nouvelle_feuille, "CM")

        ligne_insertion = cms + 1 + i
        nouvelle_feuille.insert_rows(ligne_insertion)
        nouvelle_feuille.cell(row=ligne_insertion, column=1, value=liste_inter[i][1])
        nouvelle_feuille.cell(row=ligne_insertion, column=2, value=liste_inter[i][2])

        tds = trouver_ligne(nouvelle_feuille, "TD")

        #nombre de groupes en TD
        ligne_insertion = tds + 1 + i
        nouvelle_feuille.insert_rows(ligne_insertion)
        nouvelle_feuille.cell(row=ligne_insertion, column=1, value=liste_inter[i][1])
        nouvelle_feuille.cell(row=ligne_insertion, column=2, value=liste_inter[i][3])

        tpsD = trouver_ligne(nouvelle_feuille, "TP dedoubles")

        #nombre de groupes en TP dédoublés
        ligne_insertion = tpsD + 1 + i
        nouvelle_feuille.insert_rows(ligne_insertion)
        nouvelle_feuille.cell(row=ligne_insertion, column=1, value=liste_inter[i][1])
        nouvelle_feuille.cell(row=ligne_insertion, column=2, value=liste_inter[i][4])

        tps = trouver_ligne(nouvelle_feuille, "TP non dedoubles")

        #nombre de groupes en TP non dédoublés
        ligne_insertion = tps + 1 + i
        nouvelle_feuille.insert_rows(ligne_insertion)
        nouvelle_feuille.cell(row=ligne_insertion, column=1, value=liste_inter[i][1])
        nouvelle_feuille.cell(row=ligne_insertion, column=2, value=liste_inter[i][5])


    #Récupération des données de cours par semaine 
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
            nouvelle_feuille.cell(row=ligne_date, column=2, value=(liste_cours[i][2]) * 2)
        elif liste_cours[i][1] == "TD":
            nouvelle_feuille.cell(row=ligne_date, column=3, value=(liste_cours[i][2] * 2))
        elif liste_cours[i][1] == "TP":
            nouvelle_feuille.cell(row=ligne_date, column=4, value=(liste_cours[i][2] * 2))
        else:
            nouvelle_feuille.cell(row=ligne_date, column=5, value=(liste_cours[i][2] * 2))


    #Récupération des données pour remplir tableau titulaire
    titu = "SELECT Intervenant, Feuille_title, SUBSTR(MatiereActuelle, 1, INSTR(MatiereActuelle, ' ') - 1) AS matiere FROM DONNEEPROF WHERE AlerteProf == 1 AND matiere = ?"
    cursor.execute(titu, (res,))
    liste_titu = cursor.fetchall()

    ligne_titu = trouver_ligne(nouvelle_feuille, "Service previsionnel titulaires")


    for i in range(len(liste_titu)):
        ligne_insertion = ligne_titu + i + 3
        nouvelle_feuille.insert_rows(ligne_insertion)
        nouvelle_feuille.cell(row=ligne_insertion, column=1, value=(liste_titu[i][0]))
        if "S1" or "S2" in liste_titu[i][1]:
            nouvelle_feuille.cell(row=ligne_insertion, column=3, value="BUT1")
        elif "S3" or "S4" in liste_titu[i][1]:
            nouvelle_feuille.cell(row=ligne_insertion, column=3, value="BUT2")
        else:
            nouvelle_feuille.cell(row=ligne_insertion, column=3, value="BUT3")
        if "A" in liste_titu[i][1]:
            nouvelle_feuille.cell(row=ligne_insertion, column=4, value="A")
        elif "B" in liste_titu[i][1]:
            nouvelle_feuille.cell(row=ligne_insertion, column=4, value="B")

    # Récupération des données pour remplir tableau vacataire
    vac = "SELECT Intervenant, Feuille_title, SUBSTR(MatiereActuelle, 1, INSTR(MatiereActuelle, ' ') - 1) AS matiere,d.CM, p.CM FROM DONNEEPROF d JOIN PLANRESSOURCE p ON matiere = Ressource WHERE AlerteProf == 0 AND matiere = ?"
    cursor.execute(vac, (res,))
    liste_vac = cursor.fetchall()

    ligne_vac = trouver_ligne(nouvelle_feuille, "Service previsionnel vacataires")

    for i in range(len(liste_vac)):
        ligne_insertion = ligne_vac + i + 3
        nouvelle_feuille.insert_rows(ligne_insertion)
        nouvelle_feuille.cell(row=ligne_insertion, column=1, value=(liste_vac[i][0]))
        if "S1" or "S2" in liste_titu[i][1]:
            nouvelle_feuille.cell(row=ligne_insertion, column=3, value="BUT1")
        elif "S3" or "S4" in liste_titu[i][1]:
            nouvelle_feuille.cell(row=ligne_insertion, column=3, value="BUT2")
        else:
            nouvelle_feuille.cell(row=ligne_insertion, column=3, value="BUT3")
        if "A" in liste_vac[i][1]:
            nouvelle_feuille.cell(row=ligne_insertion, column=4, value="A")
        elif "B" in liste_vac[i][1]:
            nouvelle_feuille.cell(row=ligne_insertion, column=4, value="B")
        nouvelle_feuille.cell(row=ligne_insertion, column=6, value=((liste_vac[i][3]) * (liste_vac[i][4])))




#fonction qui trouve la première ligne dans laquelle se trouve une cellule possédant la donnée recherchée
def trouver_ligne(feuille, contenu_cible):
    for row_number, row in enumerate(feuille.iter_rows(values_only=True), start=1):
        if contenu_cible in row:
            return row_number
    return None

# fonction permettant de rajouter les mots en gras, ou les cellules fusionnées qui n'apparaissent pas dans la nouvelle feuille
def applicationForme(feuille):
    feuille['A1'].font = Font(bold=True)
    feuille['F1'].font = Font(bold=True)
    feuille['A6'].font = Font(bold=True)
    feuille['A8'].font = Font(bold=True)
    feuille['A17'].font = Font(bold=True)
    feuille.merge_cells('A22:B23')
    feuille.merge_cells('C22:C23')
    feuille.merge_cells('D22:D23')
    feuille.merge_cells('E22:E23')
    feuille.merge_cells('F22:K22')
    feuille.merge_cells('L22:Q22')
    feuille.merge_cells('A26:B27')
    feuille.merge_cells('C26:C27')
    feuille.merge_cells('D26:D27')
    feuille.merge_cells('E26:E27')
    feuille.merge_cells('F26:K26')
    feuille.merge_cells('L26:Q26')

#recuperation des données pour chaque ressource existante
requeteRemplissage = "SELECT Ressource, CM, TD, TP, Acronyme FROM PLANRESSOURCE WHERE Ressource NOT LIKE '%SAE%' GROUP BY Ressource ORDER BY Ressource "
cursor.execute(requeteRemplissage)
liste_res = cursor.fetchall()

for i in range(len(liste_res)):
     Remplissage(liste_res[i][0], liste_res[i][1], liste_res[i][2], liste_res[i][3], liste_res[i][4])


# Enregistrer le classeur modifié dans un nouveau fichier
classeur_existant.save('Ressources.xlsx')
