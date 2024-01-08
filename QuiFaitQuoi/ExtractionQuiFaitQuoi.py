import pandas as pd
import openpyxl as op

"""
Le premier Intervenant renseigné dans chaque ressources est considéré par convention comme le responsable de celle-ci.
"""

# On charge le classeur Excel
QuiFaitQuoi = pd.ExcelFile('../Documents/QuiFaitQuoi_beta1.ods')

# On établie la liste des feuilles qui nous intéressent
ListeFeuillesQuiFaitQuoi = sorted(QuiFaitQuoi.sheet_names)

print(ListeFeuillesQuiFaitQuoi)
