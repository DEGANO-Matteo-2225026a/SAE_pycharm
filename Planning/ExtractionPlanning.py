import pandas as pd

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

"""
Table de données planning :
-
"""

Planning = pd.read_excel('Planning 2023-2024.xlsx', sheet_name=None)
ListeFeuillesPlanning = xls.sheet_names
print(ListeFeuillesPlanning)
rn