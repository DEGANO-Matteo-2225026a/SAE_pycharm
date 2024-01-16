import sqlite3 as sql

import sys
sys.path.insert(0, '../Planning')
from ExtractionPlanning import *


sqlConnection = sql.connect("../SAE.db")
cursor = sqlConnection.cursor()

fichierSortie = open("../Sortie/rapportErreurs.txt", 'w')
fichierSortie.write("Début rapport d'erreurs.\n")


def fusionRessourcesDivisees(dict_ress):
    for key in dict_ress.keys():
        if (key[len(key)-2:] == "-1"):
            ressAFusionner.append(key[:len(key)-2])

    for ress in ressAFusionner:





def detectionProblemeTotalCours():

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
    for semestre in TableauDonnees:
        for ressource in semestre:
            if (isinstance(ressource[3], str) or ressource[1] == 'SAE'):
                continue
            else:
                ressourcesAComparer[ressource[1]] = [ressource[2], ressource[3], ressource[4]]

    ressourcesAComparer = dict(sorted(ressourcesAComparer.items()))

    # print(ressourcesComparateur)
    print(ressourcesAComparer)
    fusionRessourcesDivisees(ressourcesAComparer)

    """
    TODO :
    - Récupérer les heures effectuées dans le planning a droite des noms de ressource dans la table Planning
    - Utiliser des fonctions provenant de ExtractionPlanning pour récupérer les données
    - Stocker les informations récupérées dans une bibliothèque pour comparaison
    - Pour la comparaison, utiliser une boucle de comparaison et généraliser les noms de ressource
    """


detectionProblemeTotalCours()

fichierSortie.write("\n\nFin rapport d'erreurs.")
fichierSortie.close()
sqlConnection.close()
