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
BUT2Info = pd.read_excel('../Documents/BUT2_INFO_AIX.xlsx', sheet_name=None)
FeuilleInfoBut2 = 'BUT 2 Parcours A FI'
FeuilleInfoBut2b = 'BUT 2 Parcours B FI'

# Récupère la feuille
S3But1Info = BUT2Info[FeuilleInfoBut2]
S3But1InfoB = BUT2Info[FeuilleInfoBut2b]



# ----------------- LIGNE -----------------
# Ressources S3 TRONC COMMUN
S3R301 = S3But1Info.iloc[11, [0,2,17,18,19,22,23]]
S3R302 = S3But1Info.iloc[12, [0,2,17,18,19,22,23]]
S3R303 = S3But1Info.iloc[13, [0,2,17,18,19,22,23]]
S3R304 = S3But1Info.iloc[14, [0,2,17,18,19,22,23]]
S3R305 = S3But1Info.iloc[15, [0,2,17,18,19,22,23]]
S3R306 = S3But1Info.iloc[16, [0,2,17,18,19,22,23]]
S3R307 = S3But1Info.iloc[17, [0,2,17,18,19,22,23]]
S3R308 = S3But1Info.iloc[18, [0,2,17,18,19,22,23]]
S3R309 = S3But1Info.iloc[19, [0,2,17,18,19,22,23]]
S3R310 = S3But1Info.iloc[20, [0,2,17,18,19,22,23]]
S3R311 = S3But1Info.iloc[21, [0,2,17,18,19,22,23]]
S3R312 = S3But1Info.iloc[22, [0,2,17,18,19,22,23]]
S3R313 = S3But1Info.iloc[23, [0,2,17,18,19,22,23]]
S3R314 = S3But1Info.iloc[24, [0,2,17,18,19,22,23]]
# SAE
SAE3 = S3But1Info.iloc[26, [0,2,17,18,19,22,23]]
# Projet tutoré           
SProjet = S3But1Info.iloc[27, [0,2,17,18,19,22,23]]
# Portfolio
SPortfolio = S3But1Info.iloc[28, [0,2,17,18,19,22,23]]
# Spécialité parcours A
S3AL1 = S3But1Info.iloc[25, [0,2,17,18,19,22,23]]
# Spécialité parcours B
S3R16X = S3But1InfoB.iloc[24, [0,2,17,18,19,22,23]]
S3P1BX = S3But1InfoB.iloc[25, [0,2,17,18,19,22,23]]



# Ressources S4

S4R401 = S3But1Info.iloc[34, [0,2,17,18,19,22,23]]
S4R402 = S3But1Info.iloc[35, [0,2,17,18,19,22,23]]
S4R403 = S3But1Info.iloc[36, [0,2,17,18,19,22,23]]
S4R404 = S3But1Info.iloc[37, [0,2,17,18,19,22,23]]
S4R405 = S3But1Info.iloc[38, [0,2,17,18,19,22,23]]
S4R406 = S3But1Info.iloc[39, [0,2,17,18,19,22,23]]
S4R407 = S3But1Info.iloc[40, [0,2,17,18,19,22,23]]


# Projets tutorés S4

SProjet = S3But1Info.iloc[48, [0,2,17,18,19,22,23]]
SPortfolio = S3But1Info.iloc[49, [0,2,17,18,19,22,23]]
Stage = S3But1Info.iloc[50, [0,2,17,18,19,22,23]]

# Ressources Parcours A 
S4R4A08 = S3But1Info.iloc[41, [0,2,17,18,19,22,23]]
S4R4A09 = S3But1Info.iloc[42, [0,2,17,18,19,22,23]]
S4R4A10 = S3But1Info.iloc[43, [0,2,17,18,19,22,23]]
S4R4A11 = S3But1Info.iloc[44, [0,2,17,18,19,22,23]]
S4R4A12 = S3But1Info.iloc[45, [0,2,17,18,19,22,23]]
S4AL1 = S3But1Info.iloc[46, [0,2,17,18,19,22,23]]
SAE4A01 = S3But1Info.iloc[47, [0,2,17,18,19,22,23]]

# Ressources Parcours B
S4R1BX = S3But1InfoB.iloc[40, [0,2,17,18,19,22,23]]
S4R2BX = S3But1InfoB.iloc[41, [0,2,17,18,19,22,23]]
S4R3BX = S3But1InfoB.iloc[42, [0,2,17,18,19,22,23]]
S4R4BX = S3But1InfoB.iloc[43, [0,2,17,18,19,22,23]]
S4R5BX = S3But1InfoB.iloc[44, [0,2,17,18,19,22,23]]
S4R9BX = S3But1InfoB.iloc[45, [0,2,17,18,19,22,23]]
S4R10BX = S3But1InfoB.iloc[46, [0,2,17,18,19,22,23]]




# ---- TEST ----



"""
TO-DO :
    - Récupérer les infos et les préparer dans des listes pour ensuite les rangers
    - Refaire le même fichier type mais pour BUT 2 et BUT 3, tu verra c'est long
    car ya pleins de type de formation mais c'est tout le temps la même chose
    - Hésite pas à commenter ce que tu fais c'est toujours bon pour nous et les profs
    - Supprime pas les commentaires de manière générale on en a besoin pour faire des résumés aux profs
"""

