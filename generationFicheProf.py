class generationFicheProf:
    def run(self):
        import openpyxl
        from openpyxl import load_workbook

        # Creer un nouveau classeur Excel
        classeur = load_workbook("Ressources.xlsx")

        # Selectionner la feuille active
        ancienne_feuille = classeur.active