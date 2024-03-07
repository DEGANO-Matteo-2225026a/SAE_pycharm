class GenerationFicheProf:
    def run(self):
        import openpyxl
        from openpyxl import load_workbook
        from pathlib import Path
        from openpyxl.styles import Font

        # Creer un nouveau classeur Excel
        classeur = load_workbook("./Excels ressources/FicheProf.xlsx")
        for r in classeur:
            print(r.title)

        # Selectionner la feuille active
        ancienne_feuille = classeur['Exemple']

        def ficheProf(prof):

            # Créer une nouvelle feuille pour chaque professeur
            feuille = classeur.copy_worksheet(ancienne_feuille)
            feuille.title = prof

            # nom professeur
            cellule_selectionnee = feuille['B1']
            cellule_selectionnee.value = prof

            """
            recherche = load_workbook("Ressources.xlsx")
            f = recherche.active
            
            for f in recherche:
                nomprof = trouver_ligne(f, prof)
                if nomprof is None:
                    continue
                print(nomprof)
                """

        def trouver_ligne(feuille, contenu_cible):
            for row_number, row in enumerate(feuille.iter_rows(values_only=True), start=1):
                if contenu_cible in row:
                    return row_number

            return None

        ficheProf("Alain Casali")
        classeur.save("FicheProf.xlsx")

def main():
    # Instanciation de la classe GenerationFicheProf
    prof_generator = GenerationFicheProf()
    # Appel de la méthode run()
    prof_generator.run()

if __name__ == "__main__":
    main()

