from ExtractionPlanning import *

# Importation des modules nécessaires
import sqlite3 as sql
import sys
import openpyxl as op


class RapportErreurGenerator:
    def genRapport(self):
        # Connexion à la base de données SQLite
        sqlConnection = sql.connect("SAE.db")
        cursor = sqlConnection.cursor()

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
            CE QUI SUIT N'EST CLAIREMENT PAS OPTI,
            VOS YEUX RISQUES DE SAIGNER EN REGARDANT CE BOUT DE CODE
            CONTINUEZ A VOS RISQUES ET PERILES
            LA FACTORISATION DE CODE EST UN VESTIGE D'UNE CIVILISATION PASSEE
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
                responsableMat[row[0]] = row[1]

        fusionRessourcesDivisees(responsableMat, "responsable")
        responsableMat = dict(sorted(responsableMat.items()))

        ##################################################################

        # Comparaison des ressources et identification des erreurs
        erreursPlanningBible = {}

        totalErreursPlanningBible = 0

        for ressource in ressourcesAComparer.keys():
            if ressourcesAComparer[ressource][0] > ressourcesComparateur[ressource][0]:
                erreursPlanningBible[ressource + " total CM : "] = (
                    ressourcesAComparer[ressource][0], ressourcesComparateur[ressource][0])
                totalErreursPlanningBible += 1

            if ressourcesAComparer[ressource][1] > ressourcesComparateur[ressource][1]:
                erreursPlanningBible[ressource + " total TD : "] = (
                    ressourcesAComparer[ressource][1], ressourcesComparateur[ressource][1])
                totalErreursPlanningBible += 1

            if ressourcesAComparer[ressource][2] > ressourcesComparateur[ressource][2]:
                erreursPlanningBible[ressource + " total TP : "] = (
                    ressourcesAComparer[ressource][2], ressourcesComparateur[ressource][2])
                totalErreursPlanningBible += 1

        # Comparaison des ressources ainsi que des responsables et identification des warnings
        warningsPlanningPrevisions = {}
        erreursPlanningPrevisions = {}

        totalWarningsPlanningPrevisions = 0
        totalErreursPlanningPrevisions = 0

        for ressource in planningTotal:

            if planningTotal[ressource][0] != ressourcesComparateur[ressource][0]:
                if planningTotal[ressource][0] > ressourcesComparateur[ressource][0]:
                    erreursPlanningPrevisions[ressource + " total CM : "] = (
                        planningTotal[ressource][0], ressourcesComparateur[ressource][0])
                    totalErreursPlanningPrevisions += 1

                else:
                    warningsPlanningPrevisions[ressource + " total CM : "] = (
                        planningTotal[ressource][0], ressourcesComparateur[ressource][0])
                    totalWarningsPlanningPrevisions += 1

            if planningTotal[ressource][1] != ressourcesComparateur[ressource][1]:
                if planningTotal[ressource][1] > ressourcesComparateur[ressource][1]:
                    erreursPlanningPrevisions[ressource + " total TD : "] = (
                        planningTotal[ressource][1], ressourcesComparateur[ressource][1])
                    totalErreursPlanningPrevisions += 1

                else:
                    warningsPlanningPrevisions[ressource + " total TD : "] = (
                        planningTotal[ressource][1], ressourcesComparateur[ressource][1])
                    totalWarningsPlanningPrevisions += 1

            if planningTotal[ressource][2] != ressourcesComparateur[ressource][2]:
                if planningTotal[ressource][2] > ressourcesComparateur[ressource][2]:
                    erreursPlanningPrevisions[ressource + " total TP : "] = (
                        planningTotal[ressource][2], ressourcesComparateur[ressource][2])
                    totalErreursPlanningPrevisions += 1

                else:
                    warningsPlanningPrevisions[ressource + " total TP : "] = (
                        planningTotal[ressource][2], ressourcesComparateur[ressource][2])
                    totalWarningsPlanningPrevisions += 1

        warningsRespMat = {}

        totalWarningRespMat = 0

        for ressource in responsableMat:
            if responsableMat[ressource] != ressourcesAComparer[ressource][3]:
                warningsRespMat[ressource + " responsable ressource : "] = (
                    ressourcesAComparer[ressource][3], responsableMat[ressource])
                totalWarningRespMat += 1

        sqlConnection.close()

        out = {"totalErreursPlanningBible" : totalErreursPlanningBible,
               "totalErreursPlanningPrevisions" : totalErreursPlanningPrevisions,

               "totalWarningsPlanningPrevisions" : totalWarningsPlanningPrevisions,
               "totalWarningRespMat" : totalWarningRespMat,

               "erreursPlanningBible": erreursPlanningBible,
               "erreursPlanningPrevisions":erreursPlanningPrevisions,

               "warningsPlanningPrevisions": warningsPlanningPrevisions,
               "warningsRespMat": warningsRespMat}

        return out


    def writeRapport(self):
        rep = self.genRapport()

        totalErreursPlanningBible = rep["totalErreursPlanningBible"]
        totalErreursPlanningPrevisions = rep["totalErreursPlanningPrevisions"]

        totalWarningsPlanningPrevisions = rep["totalWarningsPlanningPrevisions"]
        totalWarningRespMat = rep["totalWarningRespMat"]

        erreursPlanningBible = rep["erreursPlanningBible"]
        erreursPlanningPrevisions = rep["erreursPlanningPrevisions"]

        warningsPlanningPrevisions = rep["warningsPlanningPrevisions"]
        warningsRespMat = rep["warningsRespMat"]

        # Ouverture du fichier de sortie pour le rapport d'erreurs
        fichierSortie = open("rapportErreurs.txt", 'w')
        fichierSortie.write("Début rapport d'erreurs.")

        sb = "\n\nErreurs : Tout dépassement d'heures trouvées sur le planning en comparaison :"
        sb += "\n   - aux heures définies par le programme de l'Etat,"
        sb += "\n   - aux heures prévues par le département"
        fichierSortie.write(sb)

        sb = "\n\nWarnings : Toute incohérence du planning n'étant pas une erreur :"
        sb += "\n   - incohérence entre les heures trouvées sur le planning et les heures prévues par le département,"
        sb += "\n   - incohérence entre le responsable de matière décrit par le planning et celui "
        sb += "\n     stocké en base de donnée"
        fichierSortie.write(sb)

        # Écriture du rapport d'erreurs dans le fichier de sortie
        totalErreurs = totalErreursPlanningBible + totalErreursPlanningPrevisions

        sb = "\n\n\nTotal erreurs trouvées : " + str(totalErreurs)
        fichierSortie.write(sb)

        sb = "\n\nErreur(s) Dépassements d'heures entre prévisions et actuelles : " + str(
            totalErreursPlanningBible) + "\n\n"
        fichierSortie.write(sb)

        if totalErreursPlanningBible == 0:
            fichierSortie.write("Rien à signaler.")
        else:
            for erreur in erreursPlanningBible.keys():
                sb = (erreur + "Attendu : " + str(erreursPlanningBible[erreur][1]) +
                      ", Trouvé :" + str(erreursPlanningBible[erreur][0]) + "\n")
                fichierSortie.write(sb)

        sb = "\nErreur(s) Dépassements d'heures planning / heures prévues : " + str(
            totalErreursPlanningPrevisions) + "\n\n"
        fichierSortie.write(sb)

        if totalErreursPlanningPrevisions == 0:
            fichierSortie.write("Rien à signaler.")
        else:
            for erreur in erreursPlanningPrevisions.keys():
                sb = (erreur + "Attendu : " + str(erreursPlanningPrevisions[erreur][1]) +
                      ", Trouvé :" + str(erreursPlanningPrevisions[erreur][0]) + "\n")
                fichierSortie.write(sb)

        # Écriture du rapport de warning dans le fichier de sortie
        totalWarnings = totalWarningsPlanningPrevisions + totalWarningRespMat

        sb = "\n\nTotal warnings trouvées : " + str(totalWarnings)
        fichierSortie.write(sb)

        sb = "\n\nWarning(s) Incohérence planning / heures prévues : " + str(
            totalWarningsPlanningPrevisions) + "\n\n"
        fichierSortie.write(sb)

        if totalWarningsPlanningPrevisions == 0:
            fichierSortie.write("Rien à signaler.")
        else:
            for warning in warningsPlanningPrevisions.keys():
                sb = warning + "Attendu : " + str(warningsPlanningPrevisions[warning][1]) + ", Trouvé : " + str(
                    warningsPlanningPrevisions[warning][0]) + "\n"
                fichierSortie.write(sb)

        sb = "\n\nWarning(s) Incohérence responsable matière : " + str(
            totalWarningRespMat) + "\n\n"
        fichierSortie.write(sb)

        if totalWarningRespMat == 0:
            fichierSortie.write("Rien à signaler.")
        else:
            for warning in warningsRespMat.keys():
                sb = warning + "Attendu : " + str(warningsRespMat[warning][1]) + ", Trouvé : " + str(
                    warningsRespMat[warning][0]) + "\n"
                fichierSortie.write(sb)

        # Fermeture du fichier de sortie, de la connexion à la base de données et du fichier Excel
        fichierSortie.write("\n\nFin rapport d'erreurs.\n")
        fichierSortie.close()

    def run(self):
        self.writeRapport()


def main():
    # Instanciation de la classe RapportErreurGenerator
    rapport_generator = RapportErreurGenerator()
    # Appel de la méthode run()
    rapport_generator.run()


if __name__ == "__main__":
    main()
