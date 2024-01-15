
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

ROUGE = SEMAINE MORTE DONC OSEF
"""

# On charge le classeur Excel
# import pandas as pd
# Planning = pd.ExcelFile('../Documents/Planning_2023-2024.xlsx')
PlanningInfo = op.load_workbook('../Documents/Planning_2023-2024.xlsx',data_only=True)

def PurgeFeuille(Planning):
    FeuilleASupprimer = []

    for Feuille in Planning.sheetnames:
        if Feuille[0] != "S" or not Feuille[1].isdigit():
            FeuilleASupprimer.append(Feuille)

    for Feuille in FeuilleASupprimer:
        Planning.remove(Planning[Feuille])

def LocateRessources(Feuille,LimiteBoucle,LimiteDroite):

    IndiceLigne = LimiteBoucle + 2
    IndiceColonne = LimiteDroite

    for i in range(1,LimiteBoucle):
        for j in range(1,LimiteDroite):
            if Feuille.cell(IndiceLigne + i,j).fill.start_color.index != '00000000' and Feuille.cell(IndiceLigne + i,j).value == 'X':

                # Indice du nom de la matière d'après ce qui est écrit
                TrouveNom = 1
                while Feuille.cell(IndiceLigne + i, j + TrouveNom).value == None:
                    TrouveNom += 1

                # Nombre de CM de la matière d'après ce qui est écrit
                ValeurCM = 0
                if Feuille.cell(IndiceLigne + i, j + TrouveNom + 3).value != None:
                    ValeurCM = Feuille.cell(IndiceLigne + i, j + TrouveNom + 3).value

                # Nombre de TD de la matière d'après ce qui est écrit
                ValeurTD = 0
                if Feuille.cell(IndiceLigne + i, j + TrouveNom + 5).value != None:
                    ValeurTD = Feuille.cell(IndiceLigne + i, j + TrouveNom + 5).value

                # Nombre de TP de la matière d'après ce qui est écrit
                ValeurTP = 0
                if Feuille.cell(IndiceLigne + i, j + TrouveNom + 7).value != None:
                    ValeurTP = Feuille.cell(IndiceLigne + i, j + TrouveNom + 7).value

                # Acronyme du Responsable de la matière d'après ce qui est écrit
                Responsable = "Personne"
                if Feuille.cell(IndiceLigne + i, j + TrouveNom + 10).value != None:
                    Responsable = Feuille.cell(IndiceLigne + i, j + TrouveNom + 10).value

                print(Feuille.cell(IndiceLigne + i, j + TrouveNom).value,ValeurCM,ValeurTD,ValeurTP,Responsable)
    return

def LocateDate(Feuille):

    # On recherche l'Indice de la colonne des Dates
    IndiceDate = 1
    while Feuille.cell(1,IndiceDate).value != "Date":
        IndiceDate += 1

    # On cherche le nombre de semaine pour nos boucles.
    LongueurDate = 1
    BlancsDate = 0
    while BlancsDate != 3:
        if Feuille.cell(LongueurDate, IndiceDate).value != None:
            BlancsDate = 0
        else:
            BlancsDate += 1
        LongueurDate += 1

    # On oublie pas de retirer les lignes vides servant à confirmer la fin des infos de la colonne
    LongueurDate = LongueurDate - 4

    return IndiceDate, LongueurDate

def LocateLimite(Feuille,ColonneDate):

    LimiteGauche = 1
    LimiteDroite = ColonneDate

    # On cherche l'indice du début de la première colonne "Cours"
    while Feuille.cell(1,LimiteGauche).value != "Cours":
        LimiteGauche += 1

    # On cherche l'indice de fin de la deuxième colonne "Test"
    while Feuille.cell(1,LimiteDroite-1).value != "Test" or Feuille.cell(1,LimiteDroite).fill.start_color.index != '00000000':
        LimiteDroite += 1

    return LimiteDroite,LimiteGauche

def RecuperationDonneesFeuille(FeuilleActuelle):

    # Active la page pour openPYXL
    Feuille = PlanningInfo[FeuilleActuelle.title]

    # On récupère les informations de la colonne Date et de la longueur du tableau
    ColonneDate = LocateDate(Feuille)
    LimiteBoucle = ColonneDate[1]
    ColonneDate = ColonneDate[0]

    # On récupère les limites pour les informations du tableau
    LimiteDroite = LocateLimite(Feuille,ColonneDate)
    LimiteGauche = LimiteDroite[1]
    LimiteDroite = LimiteDroite[0]

    LocateRessources(Feuille, LimiteBoucle, LimiteDroite)

    print("VOICI LA LONGUEUR DE DATE :", LimiteBoucle, "ET SA COLONNE :", ColonneDate)
    print("VOICI LA GAUCHE DU TABLEAU :", LimiteGauche, "ET SA LIMITE DROITE :", LimiteDroite)

    return

# Automatise le changement de feuilles
def RecuperationParFeuille(ListeFeuilles):
    for Feuille in ListeFeuilles:
        RecuperationDonneesFeuille(Feuille)
        print(Feuille.title)
    return

# On établie la liste des feuilles qui nous intéressent
PurgeFeuille(PlanningInfo)

RecuperationParFeuille(PlanningInfo)
