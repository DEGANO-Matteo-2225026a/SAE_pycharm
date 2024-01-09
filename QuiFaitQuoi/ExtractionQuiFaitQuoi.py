import pandas as pd
import openpyxl as op

"""
Le premier Intervenant renseigné dans chaque ressources est considéré par convention comme le responsable de celle-ci.
"""

# On charge le classeur Excel
QuiFaitQuoi = pd.ExcelFile('../Documents/QuiFaitQuoi_beta1.xlsx')
QuiFaitQuoiInfo = op.load_workbook('../Documents/QuiFaitQuoi_beta1.xlsx',data_only=True)

# On établie la liste des feuilles qui nous intéressent
ListeFeuillesQuiFaitQuoi = sorted(QuiFaitQuoi.sheet_names)

print(ListeFeuillesQuiFaitQuoi)

def RecuperationProfMatiere(FeuilleActuelle):

    # Active la page pour openPYXL
    Feuille = QuiFaitQuoiInfo[FeuilleActuelle]

    # Compteur maximum Lignes et Colonnes
    CompteurLigne = Feuille.max_row
    CompteurColonne = Feuille.max_column

    # Index pour lee "curseur"
    IndexLigne =  1
    IndexColonne = 2


    print(CompteurColonne,CompteurLigne)
    """ 
    print(Feuille.cell(1,1).internal_value)
    print(Feuille.cell(1,1).fill.start_color.index
    """

    while (Feuille.cell(IndexLigne,2).internal_value is not None) and (Feuille.cell(IndexLigne, 2).fill.start_color.index != 00000000):
        print("IndexLigne = ", IndexLigne)
        print(Feuille.cell(IndexLigne, IndexColonne).internal_value)
        print(Feuille.cell(IndexLigne, IndexColonne).fill.start_color.index)
        IndexLigne = IndexLigne + 1



# On parcours toutes les feuilles
def RecuperationParFeuille(ListeFeuilles):
    # On parcours toutes les feuilles fournis
    for Feuille in ListeFeuilles:
        RecuperationProfMatiere(Feuille)
        print("PAGE SUIVANTE")
    QuiFaitQuoi.close()

RecuperationParFeuille(ListeFeuillesQuiFaitQuoi)
