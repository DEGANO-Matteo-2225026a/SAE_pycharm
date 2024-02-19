from ExtractionPlanning import *
class RapportErreurGenerator:
    def run(self):
        # Importation des modules nécessaires
        import sqlite3 as sql
        import sys
        import openpyxl as op


        # Connexion à la base de données SQLite
        sqlConnection = sql.connect("SAE.db")
        cursor = sqlConnection.cursor()

        # Ouverture du fichier de sortie pour le rapport d'erreurs
        fichierSortie = open("rapportErreurs.txt", 'w')
        fichierSortie.write("Début rapport d'erreurs.\n\n")

        # Chargement des données depuis le fichier Excel
        PlanningInfo = op.load_workbook("Documents/Planning_2023-2024-2.xlsx", data_only=True)
        PurgeFeuille(PlanningInfo)
        TableauDonnees = RecuperationParFeuille(PlanningInfo)

        # Fonction de fusion des ressources divisées
        def fusionRessourcesDivisees(dico_ress, mode):

            ressAFusionner = []

            # Identification des ressources à fusionner
            for clef in dico_ress.keys():
                if clef[len(clef) - 2] == '-':
                    ressAFusionner.append(clef[:len(clef) - 2])

            aDetruire = []

            temp_cm = 0
            temp_td = 0
            temp_tp = 0
            temp_resp = ""

            """
            CE QUI SUIT N'EST CLAIREMENT PAS OPTI ET FRANCHEMENT DEGUELASSE MAIS J'AI PAS EU LE TEMPS DE FAIRE MIEUX
            LA FACTORISATION EST UN VESTIGE D'UNE CIVILISATION PASSEE
            """

            # Condition pour séparer le cas ou la fusion de ressource se fait sur le responsable, ressource ou planning.
            if mode == "ressource":
                # Fusion des ressources et mise à jour du dictionnaire
                for ress in ressAFusionner:
                    for clef in dico_ress.keys():
                        if clef[:len(clef) - 2] == ress:
                            temp_cm += dico_ress[clef][0]
                            temp_td += dico_ress[clef][1]
                            temp_tp += dico_ress[clef][2]
                            temp_resp = dico_ress[clef][3]
                            if clef not in aDetruire:
                                aDetruire.append(clef)

                    dico_ress[ress] = [temp_cm, temp_td, temp_tp, temp_resp]
                    temp_cm = 0
                    temp_td = 0
                    temp_tp = 0
                    temp_resp = ""

            elif mode == "planning":
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

            elif mode == "responsable":
                for ress in ressAFusionner:
                    for clef in dico_ress.keys():
                        if clef[:len(clef) - 2] == ress:
                            temp_resp = dico_ress[clef]
                            if clef not in aDetruire:
                                aDetruire.append(clef)

                    dico_ress[ress] = temp_resp
                    temp_resp = ""

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
                    ressourcesAComparer[ressource[1]] = [ressource[2], ressource[3], ressource[4], ressource[5]]

        # Fusion des ressources divisées dans le dictionnaire à comparer
        fusionRessourcesDivisees(ressourcesAComparer, "ressource")
        ressourcesAComparer = dict(sorted(ressourcesAComparer.items()))

        # Comparaison des ressources et identification des erreurs
        erreurs = {}
        totalErreurs = 0

        for ressource in ressourcesAComparer.keys():
            if ressourcesAComparer[ressource][0] > ressourcesComparateur[ressource][0]:
                erreurs[ressource + " total CM : "] = (
                ressourcesAComparer[ressource][0], ressourcesComparateur[ressource][0])
                totalErreurs += 1
            if ressourcesAComparer[ressource][1] > ressourcesComparateur[ressource][1]:
                erreurs[ressource + " total TD : "] = (
                ressourcesAComparer[ressource][1], ressourcesComparateur[ressource][1])
                totalErreurs += 1
            if ressourcesAComparer[ressource][2] > ressourcesComparateur[ressource][2]:
                erreurs[ressource + " total TP : "] = (
                ressourcesAComparer[ressource][2], ressourcesComparateur[ressource][2])
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

        cursor.execute("SELECT ressource, typecours FROM PLANINFO")

        for activite in cursor.fetchall():
            if activite[0] not in planningTotal:
                planningTotal[activite[0]] = [0, 0, 0]

            if activite[1] == 'Cours':
                planningTotal[activite[0]][0] += 2
            if activite[1] == 'TD':
                planningTotal[activite[0]][1] += 2
            if activite[1] == 'TP':
                planningTotal[activite[0]][2] += 2

        # Suppression des activités non pertinentes
        aDetruire = []
        for clef in planningTotal.keys():
            if clef[0] != 'R':
                aDetruire.append(clef)

        detruireElements(aDetruire, planningTotal)
        fusionRessourcesDivisees(planningTotal, "planning")
        planningTotal = dict(sorted(planningTotal.items()))

        # Creation d'un dictionnaire contenant tout les responsables de matière
        cursor.execute("SELECT ressource, acronyme FROM PLANRESSOURCE")

        responsableMat = {}

        for row in cursor.fetchall():
            if row[0][:3] != "SAE" and row[0] not in responsableMat.keys():
                print(row[0], row[1])
                responsableMat[row[0]] = row[1]

        fusionRessourcesDivisees(responsableMat, "responsable")
        responsableMat = dict(sorted(responsableMat.items()))

        # Comparaison des ressources ainsi que des responsables et identification des warnings
        warnings = {}
        totalWarnings = 0

        for ressource in planningTotal:

            if planningTotal[ressource][0] != ressourcesComparateur[ressource][0]:
                warnings[ressource + " total CM : "] = (
                planningTotal[ressource][0], ressourcesComparateur[ressource][0])
                totalWarnings += 1
            if planningTotal[ressource][1] != ressourcesComparateur[ressource][1]:
                warnings[ressource + " total TD : "] = (
                planningTotal[ressource][1], ressourcesComparateur[ressource][1])
                totalWarnings += 1
            if planningTotal[ressource][2] != ressourcesComparateur[ressource][2]:
                warnings[ressource + " total TP : "] = (
                planningTotal[ressource][2], ressourcesComparateur[ressource][2])
                totalWarnings += 1

        for ressource in responsableMat:
            if responsableMat[ressource] != ressourcesAComparer[ressource][3]:
                warnings[ressource + " responsable ressource : "] = (
                ressourcesAComparer[ressource][3], responsableMat[ressource])
                totalWarnings += 1

        # Écriture du rapport de warning dans le fichier de sortie
        sb = "\n\nWarning(s) Incohérence planning / heures prévues, responsable de matière : " + str(
            totalWarnings) + "\n\n"
        fichierSortie.write(sb)

        if totalWarnings == 0:
            fichierSortie.write("Rien à signaler.")
        else:
            for warning in warnings.keys():
                sb = warning + "Attendu : " + str(warnings[warning][1]) + ", Trouvé : " + str(
                    warnings[warning][0]) + "\n"
                fichierSortie.write(sb)

        # Fermeture du fichier de sortie, de la connexion à la base de données et du fichier Excel
        fichierSortie.write("\nFin rapport d'erreurs.\n")
        fichierSortie.close()
        sqlConnection.close()

def main():
    # Instanciation de la classe RapportErreurGenerator
    rapport_generator = RapportErreurGenerator()
    # Appel de la méthode run()
    rapport_generator.run()

if __name__ == "__main__":
    main()