import pandas as pd
import openpyxl as op

"""
Le premier Intervenant renseigné dans chaque ressources est considéré par convention comme le responsable de celle-ci.
"""

# On charge le classeur Excel
QuiFaitQuoi = pd.ExcelFile('../Documents/QuiFaitQuoiPOURTEST.xlsx')
QuiFaitQuoiInfo = op.load_workbook('../Documents/QuiFaitQuoiPOURTEST.xlsx',data_only=True)

# On établie la liste des feuilles qui nous intéressent
ListeFeuillesQuiFaitQuoi = sorted(QuiFaitQuoi.sheet_names)

print(ListeFeuillesQuiFaitQuoi)

def RecuperationProfMatiere(FeuilleActuelle):

    # Active la page pour openPYXL
    Feuille = QuiFaitQuoiInfo[FeuilleActuelle]

    # Compteur maximum Lignes et Colonnes
    CompteurLigne = Feuille.max_row
    CompteurColonne = Feuille.max_column

    # Index pour le "curseur"
    IndexColonne = 2

    """
    print(CompteurColonne,CompteurLigne) # 11 Co 44 Li atm 
    print(Feuille.cell(1,1).internal_value)
    print(Feuille.cell(1,1).fill.start_color.index
    """

    for IndexLigne in range(2,CompteurLigne):
        if (Feuille.cell(IndexLigne, IndexColonne).fill.start_color.index) != '00000000':
            MatiereActuelle = Feuille.cell(IndexLigne, IndexColonne-1).internal_value
            AlerteProf = 1
            print("ALERTE NOUVELLE MATIERE")
            print("VOICI LE NOM DE LA MATIERE : ", MatiereActuelle)
            IndexLigne += 1

        Intervenant = Feuille.cell(IndexLigne, IndexColonne).internal_value
        Acronyme = Feuille.cell(IndexLigne, IndexColonne+1).internal_value
        Titulaire = Feuille.cell(IndexLigne, IndexColonne+2).internal_value
        NombreGroupes = Feuille.cell(IndexLigne, IndexColonne+3).internal_value
        CM = Feuille.cell(IndexLigne, IndexColonne+4).internal_value
        TDNonD = Feuille.cell(IndexLigne, IndexColonne+5).internal_value
        TPD = Feuille.cell(IndexLigne, IndexColonne+6).internal_value
        Test = Feuille.cell(IndexLigne, IndexColonne+7).internal_value

        print("Matière = ", MatiereActuelle,"," ,"Intervenant = ",Intervenant,"," ," Acronyme = ",Acronyme,"," ,
              "Titulaire = ",Titulaire,",","NombreGroupes = ",NombreGroupes,",","NombreGroupes = ",NombreGroupes,",",
              "TDNonD = ",TDNonD,",","TPD = ",TPD,",","Test = ",Test,)

        """
        print(Feuille.cell(IndexLigne, IndexColonne).internal_value)
        print(Feuille.cell(IndexLigne, IndexColonne).fill.start_color.index)
        """

        IndexLigne += 1

# On parcours toutes les feuilles
def RecuperationParFeuille(ListeFeuilles):
    # On parcours toutes les feuilles fournis
    for Feuille in ListeFeuilles:
        RecuperationProfMatiere(Feuille)
        print("")
        print("PAGE SUIVANTE")
        print("PAGE SUIVANTE")
        print("PAGE SUIVANTE")
        print("")

    QuiFaitQuoi.close()

RecuperationParFeuille(ListeFeuillesQuiFaitQuoi)
