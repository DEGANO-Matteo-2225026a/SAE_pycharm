# Importation des modules nécessaires
import sqlite3 as sql
import sys
sys.path.insert(0, '../Planning')  # Ajout du chemin d'accès pour le module ExtractionPlanning
from ExtractionPlanning import *

# Connexion à la base de données SQLite
sqlConnection = sql.connect("../SAE.db")
cursor = sqlConnection.cursor()

# Ouverture du fichier de sortie pour le rapport d'erreurs
fichierSortie = open("../Sortie/rapportErreurs.txt", 'w')
fichierSortie.write("Début rapport d'erreurs.\n\n\n")

# Chargement des données depuis le fichier Excel
PlanningInfo = op.load_workbook('../Documents/Planning_2023-2024-2.xlsx', data_only=True)
PurgeFeuille(PlanningInfo)
TableauDonnees = RecuperationParFeuille(PlanningInfo)


# Fonction de fusion des ressources divisées
def fusionRessourcesDivisees(dict_ress):
    ressAFusionner = []

    # Identification des ressources à fusionner
    for key in dict_ress.keys():
        if key[len(key)-2:] == "-1":
            ressAFusionner.append(key[:len(key)-2])

    temp_cm = 0
    temp_td = 0
    temp_tp = 0

    aDetruire = []

    # Fusion des ressources et mise à jour du dictionnaire
    for ress in ressAFusionner:
        for key in dict_ress.keys():
            if key[:len(key)-2] == ress:
                temp_cm += dict_ress[key][0]
                temp_td += dict_ress[key][1]
                temp_tp += dict_ress[key][2]
                aDetruire.append(key)

        dict_ress[ress] = [temp_cm, temp_td, temp_tp]
        temp_cm = 0
        temp_td = 0
        temp_tp = 0

    # Suppression des ressources fusionnées
    for detritus in aDetruire:
        del dict_ress[detritus]


def detruireElements(aDetruire, dict):
    for detritus in aDetruire:
        del dict[detritus]


# Exécution de la requête SQL pour récupérer les données de référence depuis la table 'BIBLE'
cursor.execute("SELECT libelle_simple, total_cm, total_td, total_tp  FROM BIBLE;")

# Création d'un dictionnaire avec les données de référence
ressourcesComparateur = {}
for row in cursor.fetchall():
    ressourcesComparateur[row[0]] = (row[1], row[2], row[3])

# Création d'un dictionnaire avec les données à comparer
ressourcesAComparer = {}
for semestre in TableauDonnees[0]:
    for ressource in semestre:
        if isinstance(ressource[3], str) or ressource[1][:3] == 'SAE':
            continue
        else:
            ressourcesAComparer[ressource[1]] = [ressource[2], ressource[3], ressource[4]]

# Fusion des ressources divisées dans le dictionnaire à comparer
fusionRessourcesDivisees(ressourcesAComparer)
ressourcesAComparer = dict(sorted(ressourcesAComparer.items()))

# Comparaison des ressources et identification des erreurs
erreurs = {}
totalErreurs = 0

for ressource in ressourcesAComparer.keys():
    if ressourcesAComparer[ressource][0] > ressourcesComparateur[ressource][0]:
        erreurs[ressource + " total cm : "] = (ressourcesAComparer[ressource][0], ressourcesComparateur[ressource][0])
        totalErreurs += 1
    if ressourcesAComparer[ressource][1] > ressourcesComparateur[ressource][1]:
        erreurs[ressource + " total td : "] = (ressourcesAComparer[ressource][1], ressourcesComparateur[ressource][1])
        totalErreurs += 1
    if ressourcesAComparer[ressource][2] > ressourcesComparateur[ressource][2]:
        erreurs[ressource + " total tp : "] = (ressourcesAComparer[ressource][2], ressourcesComparateur[ressource][2])

# Écriture du rapport d'erreurs dans le fichier de sortie
sb = "Erreur(s) Dépassements d'heures entre prévisions et actuelles : " + str(totalErreurs) + "\n\n"
fichierSortie.write(sb)

if totalErreurs == 0:
    fichierSortie.write("Rien à signaler.")
else:
    for erreur in erreurs.keys():
        sb = erreur + "Attendu : " + str(erreurs[erreur][1]) + ", Trouvé :" + str(erreurs[erreur][0]) + "\n"
        fichierSortie.write(sb)



planningTotal = {}
for activite in TableauDonnees[1]:
    if activite[1] not in planningTotal:
        print(activite[1])
        planningTotal[activite[1]] = [0, 0, 0]

    if activite[2] == 'Cours':
        planningTotal[activite[1]][0] += 1.5
    if activite[2] == 'TD':
        planningTotal[activite[1]][1] += 2
    if activite[2] == 'TP':
        planningTotal[activite[1]][2] += 2

aDetruire = []
for key in planningTotal.keys():
    if key[0] != 'R':
        aDetruire.append(key)

detruireElements(aDetruire, planningTotal)
fusionRessourcesDivisees(planningTotal)
planningTotal = dict(sorted(planningTotal.items()))
print(planningTotal)

"""
TODO :
- Modifier la couleur de R2.13 a la couleur 51 51 255
- Ajouter une fonction Mettre en forme
- Compare la table QuiFaitQuoi et les infos de Planning si incohérences entre nombre d'heures que font profs de la matière et heures de matière
"""


# Fermeture du fichier de sortie, de la connexion à la base de données et du fichier Excel
fichierSortie.write("\n\nFin rapport d'erreurs.\n")
fichierSortie.close()
sqlConnection.close()
