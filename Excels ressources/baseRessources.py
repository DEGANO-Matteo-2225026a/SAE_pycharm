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

feuille['A21'] = 'Service prévisionnel vacataires'
feuille['A22'] = 'Nom'
feuille['C22'] = 'BUT1/BUT2/BUT3'
feuille['D22'] = 'ParcoursA \n ParcoursB'
feuille['E22'] = 'FI \n FA'
feuille['F22'] = 'Nombres d\'heures prévues\nde septembre à décembre'
feuille['L22'] = 'Nombres d\'heures prévues\nde janvier à août'
feuille['F23'] = 'CM'
feuille['G23'] = 'TD'
feuille['H23'] = 'TP\n(en 1/2 groupe)'
feuille['I23'] = 'TP\n(non\ndédoublé)'
feuille['J23'] = 'TP à\ndéclarer\nARES'
feuille['K23'] = 'Total en\nHETD'
feuille['L23'] = 'CM'
feuille['M23'] = 'TD'
feuille['N23'] = 'TP\n(en 1/2 groupe)'
feuille['O23'] = 'TP\n(non\ndédoublé)'
feuille['P23'] = 'TP à\ndéclarer\nARES'
feuille['Q23'] = 'Total en\nHETD'
feuille['R23'] = 'Total en\nHETD'

feuille.merge_cells('A22','A23','B22','B23')
feuille.merge_cells('C22','C23')
feuille.merge_cells('D22','D23')
feuille.merge_cells('E22','E23')
feuille.merge_cells('F22','G22','H22','I22')
feuille.merge_cells('L22','M22','N22','O22')


# Enregistrer le classeur dans un fichier
classeur.save('Ressources.xlsx')
