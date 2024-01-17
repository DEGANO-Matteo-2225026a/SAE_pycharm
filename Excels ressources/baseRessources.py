import openpyxl
from openpyxl.styles import Font

# Créer un nouveau classeur Excel
classeur = openpyxl.Workbook()

# Sélectionner la feuille active
feuille = classeur.active

# Ajouter des données à la feuille
feuille['A1'] = 'Ressource'
feuille['A1'].font = Font(bold=True)
feuille['F1'] = 'Responsable'
feuille['F1'].font = Font(bold=True)
feuille['A3'] = 'Maquette'
feuille['B3'] = 'CM (h)'
feuille['C3'] = 'TD (h)'
feuille['D3'] = 'TP (h)'
feuille['A6'] = 'Intervenants'
feuille['A6'].font = Font(bold=True)
feuille['A8'] = 'Repartition'
feuille['A8'].font = Font(bold=True)
feuille['B8'] = 'Nombre de Groupes'
feuille['A9'] = 'CM'
feuille['A11'] = 'TD'
feuille['A13'] = 'TP dedoubles'
feuille['A15'] = 'TP non dedoubles'
feuille['A17'] = 'Organisation detaillee'
feuille['A17'].font = Font(bold=True)
feuille['B19'] = 'CM heures'
feuille['C19'] = 'TD heures'
feuille['D19'] = 'TP heures'
feuille['E19'] = 'Test'



# Enregistrer le classeur dans un fichier
classeur.save('Ressources.xlsx')
