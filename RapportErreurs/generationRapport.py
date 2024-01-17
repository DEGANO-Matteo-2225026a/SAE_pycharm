# Importation des modules nécessaires
import sqlite3 as sql
import sys
import openpyxl as op

# Ajout du chemin d'accès pour le module ExtractionPlanning
sys.path.insert(0, '../Planning')
from ExtractionPlanning import *

# Connexion à la base de données SQLite
sqlConnection = sql.connect("../SAE.db")
cursor = sqlConnection.cursor()

# Ouverture du fichier de sortie pour le rapport d'erreurs
fichierSortie = open("../Sortie/rapportErreurs.txt", 'w')
fichierSortie.write("Début rapport d'erreurs.\n\n")

# Chargement des données depuis le fichier Excel
PlanningInfo = op.load_workbook('../Documents/Planning_2023-2024-2.xlsx', data_only=True)
PurgeFeuille(PlanningInfo)
TableauDonnees = RecuperationParFeuille(PlanningInfo)


# Fonction de fusion des ressources divisées
def fusionRessourcesDivisees(dico_ress):
    ressAFusionner = []


    # Identification des ressources à fusionner
    for clef in dico_ress.keys():
        if clef[len(clef) - 2] == '-':
            ressAFusionner.append(clef[:len(clef) - 2])

    temp_cm = 0
    temp_td = 0
    temp_tp = 0

    aDetruire = []


    # Fusion des ressources et mise à jour du dictionnaire
    for ress in ressAFusionner:
        for clef in dico_ress.keys():
            if clef[:len(clef) - 2] == ress:
                temp_cm += dico_ress[clef][0]
                temp_td += dico_ress[clef][1]
                temp_tp += dico_ress[clef][2]
                if clef not in aDetruire:
                    aDetruire.append(clef)

        dico_ress[ress] = [temp_cm, temp_td, temp_tp]
        temp_cm = 0
        temp_td = 0
        temp_tp = 0


    # Suppression des ressources fusionnées
    detruireElements(aDetruire, dico_ress)


# Fonction pour détruire des éléments dans un dictionnaire
def detruireElements(aDetruire, dico):
    for detritus in aDetruire:
        del dico[detritus]


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
        erreurs[ressource + " total CM : "] = (ressourcesAComparer[ressource][0], ressourcesComparateur[ressource][0])
        totalErreurs += 1
    if ressourcesAComparer[ressource][1] > ressourcesComparateur[ressource][1]:
        erreurs[ressource + " total TD : "] = (ressourcesAComparer[ressource][1], ressourcesComparateur[ressource][1])
        totalErreurs += 1
    if ressourcesAComparer[ressource][2] > ressourcesComparateur[ressource][2]:
        erreurs[ressource + " total TP : "] = (ressourcesAComparer[ressource][2], ressourcesComparateur[ressource][2])
        totalErreurs += 1


# Écriture du rapport d'erreurs dans le fichier de sortie
sb = "\nErreur(s) Dépassements d'heures entre prévisions et actuelles : " + str(totalErreurs) + "\n\n"
fichierSortie.write(sb)

if totalErreurs == 0:
    fichierSortie.write("Rien à signaler.")
else:
    for erreur in erreurs.keys():
        sb = erreur + "Attendu : " + str(erreurs[erreur][1]) + ", Trouvé :" + str(erreurs[erreur][0]) + "\n"
        fichierSortie.write(sb)


# Calcul du total des heures pour chaque activité
planningTotal = {}

for activite in TableauDonnees[1]:
    if activite[1] not in planningTotal:
        planningTotal[activite[1]] = [0, 0, 0]

    if activite[2] == 'Cours':
        planningTotal[activite[1]][0] += 2
    if activite[2] == 'TD':
        planningTotal[activite[1]][1] += 2
    if activite[2] == 'TP':
        planningTotal[activite[1]][2] += 2


# Suppression des activités non pertinentes
aDetruire = []
for clef in planningTotal.keys():
    if clef[0] != 'R':
        aDetruire.append(clef)

detruireElements(aDetruire, planningTotal)
fusionRessourcesDivisees(planningTotal)
planningTotal = dict(sorted(planningTotal.items()))


# Comparaison des ressources et identification des warnings
warnings = {}
totalWarnings = 0

for ressource in planningTotal:

    if planningTotal[ressource][0] != ressourcesComparateur[ressource][0]:
        warnings[ressource + " total CM : "] = (planningTotal[ressource][0], ressourcesComparateur[ressource][0])
        totalWarnings += 1
    if planningTotal[ressource][1] != ressourcesComparateur[ressource][1]:
        warnings[ressource + " total TD : "] = (planningTotal[ressource][1], ressourcesComparateur[ressource][1])
        totalWarnings += 1
    if planningTotal[ressource][2] != ressourcesComparateur[ressource][2]:
        warnings[ressource + " total TP : "] = (planningTotal[ressource][2], ressourcesComparateur[ressource][2])
        totalWarnings += 1


# Écriture du rapport de warning dans le fichier de sortie
sb = "\n\nWarning(s) Incohérence entre planning et heures prévues : " + str(totalWarnings) + "\n\n"
fichierSortie.write(sb)

if totalWarnings == 0:
    fichierSortie.write("Rien à signaler.")
else:
    for warning in warnings.keys():
        sb = warning + "Attendu : " + str(warnings[warning][1]) + ", Trouvé : " + str(warnings[warning][0]) + "\n"
        fichierSortie.write(sb)


# Fermeture du fichier de sortie, de la connexion à la base de données et du fichier Excel
fichierSortie.write("\nFin rapport d'erreurs.\n")
fichierSortie.close()
sqlConnection.close()
