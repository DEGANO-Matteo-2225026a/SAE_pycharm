class GenerationRessources:
    def run(self):
        import openpyxl
        from openpyxl import load_workbook
        from pathlib import Path
        from openpyxl.styles import Font
        import sqlite3

        # connection à la base de donnée
        db_connection = sqlite3.connect("SAE.db")
        cursor = db_connection.cursor()

        # Creer un nouveau classeur Excel
        classeur = load_workbook("./Excels ressources/FichierRessources.xlsx")

        # Selectionner la feuille active
        ancienne_feuille = classeur.active

        def Remplissage(res, cmtot, tdtot, tptot, resp):
            # Créer une nouvelle feuille pour chaque matière
            feuille = classeur.copy_worksheet(ancienne_feuille)
            feuille.title = res


            # nom responsable de ressource
            cellule_selectionnee = feuille['H1']
            cellule_selectionnee.value = resp

            # nom ressource
            cellule_selectionnee = feuille['C1']
            cellule_selectionnee.value = res

            # volume horaire total
            cellule_selectionnee = feuille['B4']
            cellule_selectionnee.value = cmtot

            cellule_selectionnee = feuille['C4']
            cellule_selectionnee.value = tdtot

            cellule_selectionnee = feuille['D4']
            cellule_selectionnee.value = tptot

            # Récupération des donnees des intervenants
            requete = "SELECT Intervenant, Acronyme, CM, NombreGroupes FROM DONNEEPROF WHERE SUBSTR(MatiereActuelle, 1, INSTR(MatiereActuelle, ' ') - 1) = ?"
            cursor.execute(requete, (res,))
            liste_inter = cursor.fetchall()

            # insertion données dans la feuille
            for i in range(len(liste_inter)):
                # nom et acronyme
                inter = trouver_ligne(feuille,"Intervenants")

                ligne_insertion = 7 + i
                feuille.cell(row=ligne_insertion, column=1, value=liste_inter[i][0])
                feuille.cell(row=ligne_insertion, column=2, value=liste_inter[i][1])

                # pourcentage du total de cm effectué par le professeur
                cms = trouver_ligne(feuille, "CM ")


                ligne_insertion = cms + 1 + i
                feuille.cell(row=ligne_insertion, column=1, value=liste_inter[i][1])
                feuille.cell(row=ligne_insertion, column=2, value=liste_inter[i][2])

                tds = trouver_ligne(feuille, "TD ")


                # nombre de groupes en TD
                ligne_insertion = tds + 1 + i
                feuille.cell(row=ligne_insertion, column=1, value=liste_inter[i][1])
                feuille.cell(row=ligne_insertion, column=2, value=liste_inter[i][3])

                tpsD = trouver_ligne(feuille,"TP dedoubles")


                # nombre de groupes en TP dédoublés
                ligne_insertion = tpsD + 1 + i
                feuille.cell(row=ligne_insertion, column=1, value=liste_inter[i][1])
                feuille.cell(row=ligne_insertion, column=2, value=liste_inter[i][3])

                tps = trouver_ligne(feuille, "TP non  dedoubles")


                # nombre de groupes en TP non dédoublés
                ligne_insertion = tps + 1 + i
                feuille.cell(row=ligne_insertion, column=1, value=liste_inter[i][1])
                feuille.cell(row=ligne_insertion, column=2, value=liste_inter[i][3])

            # Récupération des données de cours par semaine
            cours = "SELECT Semaine, TypeCours, COUNT(*) AS NombreCours FROM PLANINFO WHERE Ressource = ? AND TypeCours IS NOT NULL GROUP BY Semaine, TypeCours "
            cursor.execute(cours, (res,))
            liste_cours = cursor.fetchall()

            dates = []
            for i in range(len(liste_cours)):

                hcm = trouver_ligne(feuille, "CM (h)")
                ligne_insertion = hcm + 1 + i
                feuille.insert_rows(ligne_insertion)
                feuille.cell(row=ligne_insertion, column=1, value=liste_cours[i][0])


                #feuille.insert_rows(ligne_insertion)
                #feuille.cell(row=ligne_insertion, column=1, value=liste_cours[i][0])

                ligne_date = trouver_ligne(feuille,liste_cours[i][0])
                if ligne_date is None:
                    continue

                if liste_cours[i][1] == "Cours":
                    feuille.cell(row=ligne_date, column=2, value=(liste_cours[i][2]) * 2)
                elif liste_cours[i][1] == "TD":
                    feuille.cell(row=ligne_date, column=3, value=(liste_cours[i][2] * 2))
                elif liste_cours[i][1] == "TP":
                    feuille.cell(row=ligne_date, column=4, value=(liste_cours[i][2] * 2))
                else:
                    feuille.cell(row=ligne_date, column=5, value=(liste_cours[i][2] * 2))


            # Récupération des données pour remplir tableau titulaire
            titu = "SELECT Intervenant, Feuille_title, SUBSTR(MatiereActuelle, 1, INSTR(MatiereActuelle, ' ') - 1) AS matiere,d.CM, d.TD, TPD, TDNonD, Test, p.CM FROM DONNEEPROF d JOIN PLANRESSOURCE p ON matiere = Ressource WHERE AlerteProf == 1 AND matiere = ?"
            cursor.execute(titu, (res,))
            liste_titu = cursor.fetchall()


            ligne_titu = trouver_ligne(feuille,"Service previsionnel titulaires")

            for i in range(len(liste_titu)):
                ligne_insertion = ligne_titu + i + 3
                feuille.cell(row=ligne_insertion, column=1, value=(liste_titu[i][0]))
                if "S1" or "S2" in liste_titu[i][1]:
                    feuille.cell(row=ligne_insertion, column=3, value="BUT1")
                elif "S3" or "S4" in liste_titu[i][1]:
                    feuille.cell(row=ligne_insertion, column=3, value="BUT2")
                else:
                    feuille.cell(row=ligne_insertion, column=3, value="BUT3")
                if "A" in liste_titu[i][1]:
                    feuille.cell(row=ligne_insertion, column=4, value="A")
                elif "B" in liste_titu[i][1]:
                    feuille.cell(row=ligne_insertion, column=4, value="B")
                if "S1" in liste_titu[i][1] or "S3" in liste_titu[i][1] or "S6" in liste_titu[i][1]:
                    feuille.cell(row=ligne_insertion, column=6, value=(liste_titu[i][3] * liste_titu[i][8]) * 2)
                    feuille.cell(row=ligne_insertion, column=7, value=liste_titu[i][4] * 2)
                    feuille.cell(row=ligne_insertion, column=8, value=liste_titu[i][5] * 2)
                    feuille.cell(row=ligne_insertion, column=9, value=liste_titu[i][6] * 2)
                else:
                    feuille.cell(row=ligne_insertion, column=12, value=(liste_titu[i][3] * liste_titu[i][8]) * 2)
                    feuille.cell(row=ligne_insertion, column=13, value=liste_titu[i][4] * 2)
                    feuille.cell(row=ligne_insertion, column=14, value=liste_titu[i][5] * 2)
                    feuille.cell(row=ligne_insertion, column=15, value=liste_titu[i][6] * 2)

                valeur_cellule_8 = feuille.cell(row=ligne_insertion, column=8).value
                valeur_cellule_9 = feuille.cell(row=ligne_insertion, column=9).value

                # Vérification si les cellules ne sont pas vides
                if valeur_cellule_8 is not None and valeur_cellule_9 is not None:
                    # Calcul de hetd_ares_1
                    hetd_ares_1 = valeur_cellule_8 + valeur_cellule_9

                    # Assignation de la valeur à la cellule 10
                    feuille.cell(row=ligne_insertion, column=10, value=hetd_ares_1)

                valeur_cellule_6 = feuille.cell(row=ligne_insertion, column=6).value
                valeur_cellule_7 = feuille.cell(row=ligne_insertion, column=7).value
                # Vérification si les cellules ne sont pas vides
                if valeur_cellule_8 is not None and valeur_cellule_9 is not None:
                    hetd_total_1 = valeur_cellule_6 * 1.5 + valeur_cellule_7 + valeur_cellule_8 + valeur_cellule_9
                    feuille.cell(row=ligne_insertion, column=11, value=hetd_total_1)

                valeur_cellule_11 = feuille.cell(row=ligne_insertion, column=11).value
                valeur_cellule_17 = feuille.cell(row=ligne_insertion, column=17).value

                if valeur_cellule_11 is not None and valeur_cellule_17 is not None:
                    hetd_final = valeur_cellule_11 + valeur_cellule_17
                else:
                    hetd_final = valeur_cellule_11 if valeur_cellule_11 is not None else valeur_cellule_17
                    feuille.cell(row=ligne_insertion, column=18, value=hetd_final)

            # Récupération des données pour remplir tableau vacataire
            vac = "SELECT Intervenant, Feuille_title, SUBSTR(MatiereActuelle, 1, INSTR(MatiereActuelle, ' ') - 1) AS matiere,d.CM, d.TD, TDNonD, TPD, Test, p.CM FROM DONNEEPROF d JOIN PLANRESSOURCE p ON matiere = Ressource WHERE AlerteProf == 0 AND matiere = ?"
            cursor.execute(vac, (res,))
            liste_vac = cursor.fetchall()

            ligne_vac = trouver_ligne(feuille, "Service previsionnel vacataires")

            for i in range(len(liste_vac)):
                ligne_insertion = ligne_vac + i + 3
                feuille.cell(row=ligne_insertion, column=1, value=(liste_vac[i][0]))
                if "S1" in liste_vac[i][1] or "S2" in liste_vac[i][1]:
                    feuille.cell(row=ligne_insertion, column=3, value="BUT1")
                elif "S3" in liste_vac[i][1] or "S4" in liste_vac[i][1]:
                    feuille.cell(row=ligne_insertion, column=3, value="BUT2")
                else:
                    feuille.cell(row=ligne_insertion, column=3, value="BUT3")
                if "A" in liste_vac[i][1]:
                    feuille.cell(row=ligne_insertion, column=4, value="A")
                elif "B" in liste_vac[i][1]:
                    feuille.cell(row=ligne_insertion, column=4, value="B")

                if "S1" in liste_vac[i][1] or "S3" in liste_vac[i][1] or "S6" in liste_vac[i][1]:
                    feuille.cell(row=ligne_insertion, column=6, value=(liste_vac[i][3] * liste_vac[i][8]) * 2)
                    feuille.cell(row=ligne_insertion, column=7, value=liste_vac[i][4] * 2)
                    feuille.cell(row=ligne_insertion, column=8, value=liste_vac[i][5] * 2)
                    feuille.cell(row=ligne_insertion, column=9, value=liste_vac[i][6] * 2)

                else:
                    feuille.cell(row=ligne_insertion, column=12, value=(liste_vac[i][3] * liste_vac[i][8]) * 2)
                    feuille.cell(row=ligne_insertion, column=13, value=liste_vac[i][4] * 2)
                    feuille.cell(row=ligne_insertion, column=14, value=liste_vac[i][5] * 2)
                    feuille.cell(row=ligne_insertion, column=15, value=liste_vac[i][6] * 2)

                valeur_cellule_8 = feuille.cell(row=ligne_insertion, column=8).value
                valeur_cellule_9 = feuille.cell(row=ligne_insertion, column=9).value

                # Vérification si les cellules ne sont pas vides
                if valeur_cellule_8 is not None and valeur_cellule_9 is not None:
                    # Calcul de hetd_ares_1
                    hetd_ares_1 = valeur_cellule_8 + valeur_cellule_9 * 2 / 3

                    # Assignation de la valeur à la cellule 10
                    feuille.cell(row=ligne_insertion, column=10, value=hetd_ares_1)

                valeur_cellule_6 = feuille.cell(row=ligne_insertion, column=6).value
                valeur_cellule_7 = feuille.cell(row=ligne_insertion, column=7).value
                # Vérification si les cellules ne sont pas vides
                if valeur_cellule_8 is not None and valeur_cellule_9 is not None:
                    hetd_total = valeur_cellule_6 * 1.5 + valeur_cellule_7 + valeur_cellule_8 * 2 / 3 + valeur_cellule_9
                    feuille.cell(row=ligne_insertion, column=11, value=hetd_total)

                valeur_cellule_11 = feuille.cell(row=ligne_insertion, column=11).value
                valeur_cellule_17 = feuille.cell(row=ligne_insertion, column=17).value

                if valeur_cellule_11 is not None and valeur_cellule_17 is not None:
                    hetd_final = valeur_cellule_11 + valeur_cellule_17
                else:
                    hetd_final = valeur_cellule_11 if valeur_cellule_11 is not None else valeur_cellule_17
                    feuille.cell(row=ligne_insertion, column=18, value=hetd_final)

        # fonction qui trouve la première ligne dans laquelle se trouve une cellule possédant la donnée recherchée
        def trouver_ligne(feuille, contenu_cible):
            for row_number, row in enumerate(feuille.iter_rows(values_only=True), start=1):
                if contenu_cible in row:
                    return row_number

            return None

        # recuperation des données pour chaque ressource existante
        requeteRemplissage = "SELECT Ressource, CM, TD, TP, Acronyme FROM PLANRESSOURCE WHERE Ressource NOT LIKE '%SAE%' GROUP BY Ressource ORDER BY Ressource "
        cursor.execute(requeteRemplissage)
        liste_res = cursor.fetchall()

        for i in range(len(liste_res)):
            Remplissage(liste_res[i][0], liste_res[i][1], liste_res[i][2], liste_res[i][3], liste_res[i][4])

        # Enregistrer le classeur modifié dans un nouveau fichier
        classeur.save('Ressources.xlsx')

def main():
    # Instanciation de la classe GenerationRessources
    ressource_generator = GenerationRessources()
    # Appel de la méthode run()
    ressource_generator.run()

if __name__ == "__main__":
    main()