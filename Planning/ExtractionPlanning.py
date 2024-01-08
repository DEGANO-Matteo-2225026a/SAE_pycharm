import pandas as pd
import openpyxl as op

"""
Les données du document planning nous apparaissent avec cette logique :
- Les données concernant les matières de chaque pages de semestre sont divisées en deux parties, la deuxième commençant 
après la colonne contenant "A placer dans EDT"
- Ces deux parties servent uniquement à faciliter le repèrage des matières, elles ont bien lieux dans la même semaine.
- On remarque un compteur d'heure en Amphi, Salle Machine (Y) et Test. Ce compteur nous permet de trouver (sauf erreurs)
le total d'heure en Salle de TD -> Heures Totales - ((Amphi + Y) * 2) 

On possède par les info suivantes :
- Nom du semestre
- Date Semaine
- Numéro de Semaine
- Cours
- TD
- TP
- Test
- Sae
- Heures Info
- Heures Non Info
- Heures Totales
"""

def LocateRessources():


# Automatise le changement de feuilles
def RecuperationParFeuille(ListeFeuilles):
    for Feuille in ListeFeuilles:
        print(i)
    Planning.close()


# On charge le classeur Excel
Planning = pd.ExcelFile('../Documents/Planning_2023-2024.xlsx')

# On établie la liste des feuilles qui nous intéressent
ListeFeuillesPlanning = sorted(Planning.sheet_names)
ListeFeuillesPlanning.pop(0)
ListeFeuillesPlanning.pop(0)

RecuperationParFeuille(ListeFeuillesPlanning)

print(ListeFeuillesPlanning)
