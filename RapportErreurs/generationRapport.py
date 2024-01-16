import sqlite3 as sql

import sys
sys.path.insert(0, '../Planning')
from ExtractionPlanning import *


sqlConnection = sql.connect("../SAE.db")
cursor = sqlConnection.cursor()

fichierSortie = open("../Sortie/rapportErreurs.txt", 'w')
fichierSortie.write("Début rapport d'erreurs.\n\n\n")


def fusionRessourcesDivisees(dict_ress):

    ressAFusionner = []

    for key in dict_ress.keys():
        if (key[len(key)-2:] == "-1"):
            ressAFusionner.append(key[:len(key)-2])

    # print(ressAFusionner)

    temp_cm = 0
    temp_td = 0
    temp_tp = 0

    aDetruire = []

    for ress in ressAFusionner:
        for key in dict_ress.keys():
            if (key[:len(key)-2] == ress):
                temp_cm += dict_ress[key][0]
                temp_td += dict_ress[key][1]
                temp_tp += dict_ress[key][2]
                aDetruire.append(key)

        dict_ress[ress] = [temp_cm,temp_td,temp_tp]
        temp_cm = 0
        temp_td = 0
        temp_tp = 0

    for detritus in aDetruire:
        del dict_ress[detritus]


# Execution de la query en sql visant à récupérer les données de références depuis la bible
cursor.execute("SELECT libelle_simple, total_cm, total_td, total_tp  FROM BIBLE;")

# On récupère le nom de chaque ressources depuis la bible que l'on associe à une clef dans une biblothèque
# puis les données du total de CM, TD et TP associés a cette resource
ressourcesComparateur = {}
for row in cursor.fetchall():
    ressourcesComparateur[row[0]] = (row[1], row[2], row[3])

PlanningInfo = op.load_workbook('../Documents/Planning_2023-2024.xlsx', data_only=True)
PurgeFeuille(PlanningInfo)
TableauDonnees = RecuperationParFeuille(PlanningInfo)

ressourcesAComparer = {}
for semestre in TableauDonnees[0]:
    for ressource in semestre:
        if (isinstance(ressource[3], str) or ressource[1][:3] == 'SAE'):
            continue
        else:
            ressourcesAComparer[ressource[1]] = [ressource[2], ressource[3], ressource[4]]

fusionRessourcesDivisees(ressourcesAComparer)
ressourcesAComparer = dict(sorted(ressourcesAComparer.items()))

erreurs = {}
totalErreurs = 0

for ressource in ressourcesAComparer.keys():
    if ressourcesAComparer[ressource][0] != ressourcesComparateur[ressource][0] :
        erreurs[ressource + " total cm : "] = (ressourcesAComparer[ressource][0], ressourcesComparateur[ressource][0])
        totalErreurs += 1
    if ressourcesAComparer[ressource][1] != ressourcesComparateur[ressource][1] :
        erreurs[ressource + " total td : "] = (ressourcesAComparer[ressource][1], ressourcesComparateur[ressource][1])
        totalErreurs += 1
    if ressourcesAComparer[ressource][2] != ressourcesComparateur[ressource][2]:
        erreurs[ressource + " total tp : "] = (ressourcesAComparer[ressource][2], ressourcesComparateur[ressource][2])
        totalErreurs += 1

sb = "Erreur(s) Incohérence Planning / Total cours par matière : " + str(totalErreurs) + "\n\n"
fichierSortie.write(sb)

if totalErreurs == 0 :
    fichierSortie.write("Rien à signaler.")
else :
    for erreur in erreurs.keys() :
        sb = erreur + "Attendu : " + str(erreurs[erreur][1]) + ", Trouvé :" + str(erreurs[erreur][0]) + "\n"
        fichierSortie.write(sb)


    """
    TODO :
    - Pour la comparaison, utiliser une boucle de comparaison et généraliser les noms de ressource
    """

fichierSortie.write("\n\nFin rapport d'erreurs.\n\n")
fichierSortie.close()
sqlConnection.close()
