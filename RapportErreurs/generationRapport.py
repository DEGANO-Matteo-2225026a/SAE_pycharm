import sqlite3 as sql

"""
import sys
sys.path.insert(0, '../Planning')
from ExtractionPlanning import LocateRessource
"""


sqlConnection = sql.connect("../SAE.db")
cursor = sqlConnection.cursor()

fichierSortie = open("../Sortie/rapportErreurs.txt", 'w')
fichierSortie.write("Début rapport d'erreurs.\n")


def label(code_apogee):
    return code_apogee[3:len(code_apogee)-1]


def detectionProblemeTotalCours():

    cursor.execute("SELECT code_apogee, total_cm, total_td, total_tp  FROM RESSOURCE;")

    # On récupère le nom de chaque ressources depuis la bible que l'on associe à une clef dans une biblothèque
    # puis les données du total de CM, TD et TP associés a cette resource
    ressources = {}
    for row in cursor.fetchall():
        ressources[label(row[0])] = (row[1], row[2], row[3])

    """
    TODO :
    - Récupérer les heures effectuées dans le planning a droite des noms de ressource dans la table Planning
    - Utiliser des fonctions provenant de ExtractionPlanning pour récupérer les données
    - Stocker les informations récupérées dans une bibliothèque pour comparaison
    - Pour la comparaison, utiliser une boucle de comparaison et généraliser les noms de ressource (Possible problème
      de difficultés dans la comparaison des donnés à cause de l'utilisation du code apogée dans la première
      récupération)
    """


detectionProblemeTotalCours()

fichierSortie.write("\n\nFin rapport d'erreurs.")
fichierSortie.close()
sqlConnection.close()
