import pandas as pd
import openpyxl as op

"""
Le premier Intervenant renseigné dans chaque ressources est considéré par convention comme le responsable de celle-ci.
"""

# On charge le classeur Excel
QuiFaitQuoi = pd.ExcelFile('../Documents/QuiFaitQuoi_beta1.xlsx')

# On établie la liste des feuilles qui nous intéressent
ListeFeuillesQuiFaitQuoi = sorted(QuiFaitQuoi.sheet_names)

print(ListeFeuillesQuiFaitQuoi)

def RecuperationMatiere(FeuilleActuelle):
    Index = QuiFaitQuoi["Intervenants"]
    while couleurs != blanc and texte != null:
        if couleurs change mais != blanc


def RecuperationParFeuille(ListeFeuilles):
    for Feuille in ListeFeuilles:
        RecuperationMatiere(Feuille)
    QuiFaitQuoi.close()

RecuperationParFeuille(ListeFeuillesQuiFaitQuoi)
