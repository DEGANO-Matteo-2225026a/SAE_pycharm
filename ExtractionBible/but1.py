#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Oct 11 09:22:51 2023

@author: b22015136
"""
import pandas as pd

"""
Commencer par récuperer toutes les informations importantes qu'il faut
renseigner dans la base de données.

On essaye de faire le BUT 1 avant de passer aux autres.

En l'occurrence on récupère :
    - Code Apogée
    - Libellé complet Apogée
    - Volume Horaire Total (CM, TD, TP)
    - HETD et HETD PAcome
"""

# Indique le bon document et la bonne feuille du document
BUT1Info = pd.read_excel('Documents/BUT1_INFO_AIX.xlsx', sheet_name=None)
FeuilleInfoBut1 = 'BUT 1'

# Récupère la feuille
S1But1Info = BUT1Info[FeuilleInfoBut1]

#donnee importante 

# Donnée code Apogée
S1codeApogee = S1But1Info.iloc[8:33, 0]
S2codeApogee = S1But1Info.iloc[34:60, 0]

# Donnée Libellé complet Apogée
S1LibelleCompletApogee = S1But1Info.iloc[8:33, 2]
S2LibelleCompletApogee = S1But1Info.iloc[34:60, 2]

# Donnée Volume horaire total
#TODO: ajout TP
S1volumeHoraireTotal = S1But1Info.iloc[8:33, 17:19]
S2volumeHoraireTotal = S1But1Info.iloc[34:60, 17:19]


# Donnée nombre groupe td

S1nbGroupeTD = S1But1Info.iloc[24, 26]
S2nbGroupeTD = S1But1Info.iloc[49, 26]


# ----------------- LIGNE -----------------
# Ressources S1
#premier_element correspond au premier element du libelle pour nous facilité la tache lors de la recuperation des couleurs dans le fichier Planning
S1R101 = S1But1Info.iloc[10, [0,2,17,18,19,22,23]]
S1R102 = S1But1Info.iloc[11, [0,2,17,18,19,22,23]]
S1R103 = S1But1Info.iloc[12, [0,2,17,18,19,22,23]]
S1R104 = S1But1Info.iloc[13, [0,2,17,18,19,22,23]]
S1R105 = S1But1Info.iloc[14, [0,2,17,18,19,22,23]]
S1R106 = S1But1Info.iloc[15, [0,2,17,18,19,22,23]]
S1R107 = S1But1Info.iloc[16, [0,2,17,18,19,22,23]]
S1R108 = S1But1Info.iloc[17, [0,2,17,18,19,22,23]]
S1R109 = S1But1Info.iloc[18, [0,2,17,18,19,22,23]]
S1R110 = S1But1Info.iloc[19, [0,2,17,18,19,22,23]]
S1R111 = S1But1Info.iloc[20, [0,2,17,18,19,22,23]]
S1R112 = S1But1Info.iloc[21, [0,2,17,18,19,22,23]]
S1RL1 = S1But1Info.iloc[22, [0,2,17,18,19,22,23]]


# Sae S1
S1S101 = S1But1Info.iloc[23, [0,2,17,18,19,22,23]]
S1S102 = S1But1Info.iloc[24, [0,2,17,18,19,22,23]]
S1S103= S1But1Info.iloc[25, [0,2,17,18,19,22,23]]
S1S104 = S1But1Info.iloc[26, [0,2,17,18,19,22,23]]
S1S105 = S1But1Info.iloc[27, [0,2,17,18,19,22,23]]
S1S106 = S1But1Info.iloc[28, [0,2,17,18,19,22,23]]
S1ProjetTutoree = S1But1Info.iloc[29, [0,2,17,18,19,22,23]]
S1Portfolio = S1But1Info.iloc[23, [0,2,17,18,19,22,23]]


# Ressources S2

S2R201 = S1But1Info.iloc[36, [0,2,17,18,19,22,23]]
S2R202 = S1But1Info.iloc[37, [0,2,17,18,19,22,23]]
S2R203 = S1But1Info.iloc[38, [0,2,17,18,19,22,23]]
S2R204 = S1But1Info.iloc[39, [0,2,17,18,19,22,23]]
S2R205 = S1But1Info.iloc[40, [0,2,17,18,19,22,23]]
S2R206 = S1But1Info.iloc[41, [0,2,17,18,19,22,23]]
S2R207 = S1But1Info.iloc[42, [0,2,17,18,19,22,23]]
S2R208 = S1But1Info.iloc[43, [0,2,17,18,19,22,23]]
S2R209 = S1But1Info.iloc[44, [0,2,17,18,19,22,23]]
S2R210 = S1But1Info.iloc[45, [0,2,17,18,19,22,23]]
S2R211 = S1But1Info.iloc[46, [0,2,17,18,19,22,23]]
S2R212 = S1But1Info.iloc[47, [0,2,17,18,19,22,23]]
S2R213 = S1But1Info.iloc[48, [0,2,17,18,19,22,23]]
S2R214 = S1But1Info.iloc[49, [0,2,17,18,19,22,23]]

# Sae S2

S2S201 = S1But1Info.iloc[50, [0,2,17,18,19,22,23]]
S2S202 = S1But1Info.iloc[51, [0,2,17,18,19,22,23]]
S2S203 = S1But1Info.iloc[52, [0,2,17,18,19,22,23]]
S2S204 = S1But1Info.iloc[53, [0,2,17,18,19,22,23]]
S2S205 = S1But1Info.iloc[54, [0,2,17,18,19,22,23]]
S2S206 = S1But1Info.iloc[55, [0,2,17,18,19,22,23]]

S2ProjetTutoree = S1But1Info.iloc[56, [0,2,17,18,19,22,23]]
S2Portfolio = S1But1Info.iloc[57, [0,2,17,18,19,22,23]]


# ---- TEST ----

#print(S1R101)


"""
TO-DO :
    - Récupérer les infos et les préparer dans des listes pour ensuite les rangers
    - Refaire le même fichier type mais pour BUT 2 et BUT 3, tu verra c'est long
    car ya pleins de type de formation mais c'est tout le temps la même chose
    - Hésite pas à commenter ce que tu fais c'est toujours bon pour nous et les profs
    - Supprime pas les commentaires de manière générale on en a besoin pour faire des résumés aux profs
"""

