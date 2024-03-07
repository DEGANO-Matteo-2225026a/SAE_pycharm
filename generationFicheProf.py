class GenerationFicheProf:
    def run(self):
        import openpyxl
        from openpyxl import load_workbook
        from pathlib import Path
        from openpyxl.styles import Font

        # Creer un nouveau classeur Excel
        classeur = load_workbook("./Excels ressources/FicheProf.xlsx")

        # Selectionner la feuille active
        ancienne_feuille = classeur['Exemple']

        recherche = load_workbook("Ressources.xlsx")
        f = recherche.active

        def ficheProf(prof):

            # Créer une nouvelle feuille pour chaque professeur
            feuille = classeur.copy_worksheet(ancienne_feuille)
            feuille.title = prof

            # nom professeur
            cellule_selectionnee = feuille['B1']
            cellule_selectionnee.value = prof

            
            for f in recherche:
                nomprof = trouver_ligne(f, prof, 1)
                if nomprof is None:
                    continue
                # On cherche si le professeur est responsable de la matière
                respounon = f.cell(row=nomprof - 1, column=1).value

                # Si oui on la place responsable
                if respounon == "Intervenants":
                    feuille.insert_rows(4)
                    feuille.cell(row=4, column=1, value=f.title)
                # Sinon on la place intervenant
                else:
                    rowInter = trouver_ligne(feuille, "Interventions", 1)
                    feuille.insert_rows(rowInter + 1)
                    feuille.cell(row=rowInter + 1, column=1, value=f.title)

                # On récupère les volumes horaires pour chaque ressource
                horaires = trouver_ligne(f, prof, 11)
                if horaires is None:
                    continue
                ligneH = trouver_ligne(feuille, "Matieres", 1)
                feuille.insert_rows(ligneH + 1)
                for i in range(6, 13):
                    cellule_selectionnee = f.cell(row=horaires, column=i).value
                    if cellule_selectionnee is None:
                        cellule_selectionnee = f.cell(row=horaires, column=i+6).value
                    feuille.cell(row=ligneH + 1, column=1, value=f.title)
                    feuille.cell(row=ligneH + 1, column=i-3, value=cellule_selectionnee)
                    # On insère l'année
                    cellule_selectionnee = f.cell(row=horaires, column=3).value
                    feuille.cell(row=ligneH + 1, column=2, value=cellule_selectionnee)





        def trouver_ligne(feuille, contenu_cible, debut):
            for row_number, row in enumerate(feuille.iter_rows(values_only=True), start=1):
                if row_number >= debut and contenu_cible in row:
                    return row_number

            return None

        liste_prof = []
        for p in recherche:
            for i in range(7, p.max_row):
                prof = p.cell(row=i, column=1).value
                if prof is None:
                    break
                if prof in liste_prof:
                    continue
                else:
                    liste_prof.append(prof)
                    ficheProf(prof)

        classeur.save("FicheProf.xlsx")

def main():
    # Instanciation de la classe GenerationFicheProf
    prof_generator = GenerationFicheProf()
    # Appel de la méthode run()
    prof_generator.run()

if __name__ == "__main__":
    main()

