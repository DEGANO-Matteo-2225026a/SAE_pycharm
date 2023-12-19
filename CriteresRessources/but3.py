#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Dec 19 14:43:58 2023

@author: b22015136
"""

import pandas as pd


BUT3Info = pd.read_excel('BUT3 _INFO_AIX.xlsx', sheet_name=None)
FeuilleInfoBut3a = 'BUT 3 Parcours A FI'
FeuilleInfoBut3b = 'BUT 3 Parcours B FI'

S5But1Info = BUT3Info[FeuilleInfoBut3a]
S5But1InfoB = BUT3Info[FeuilleInfoBut3b]

# ----------------- LIGNE -----------------
# Ressources S5 TRONC COMMUN
S5R501 = S5But1Info.iloc[9, [0,2,17,18,19,22,23]]
S5R502 = S5But1Info.iloc[10, [0,2,17,18,19,22,23]]
S5R503 = S5But1Info.iloc[11, [0,2,17,18,19,22,23]]

# Projet
SProjet = S5But1Info.iloc[25, [0,2,17,18,19,22,23]]
#Portfolio
SPortfolio = S5But1Info.iloc[26, [0,2,17,18,19,22,23]]

# Ressources Parcours A

S5R1AX= S5But1Info.iloc[12, [0,2,17,18,19,22,23]]
S5R2AX= S5But1Info.iloc[13, [0,2,17,18,19,22,23]]
S5R3AX= S5But1Info.iloc[14, [0,2,17,18,19,22,23]]
S5R4AX= S5But1Info.iloc[15, [0,2,17,18,19,22,23]]
S5R5AX= S5But1Info.iloc[16, [0,2,17,18,19,22,23]]
S5R6AX= S5But1Info.iloc[17, [0,2,17,18,19,22,23]]
S5R7AX= S5But1Info.iloc[18, [0,2,17,18,19,22,23]]
S5R8AX= S5But1Info.iloc[19, [0,2,17,18,19,22,23]]
S5R9AX= S5But1Info.iloc[20, [0,2,17,18,19,22,23]]
S5R10AX= S5But1Info.iloc[21, [0,2,17,18,19,22,23]]
S5R11AX= S5But1Info.iloc[22, [0,2,17,18,19,22,23]]
S5RC04X = S5But1Info.iloc[23, [0,2,17,18,19,22,23]]
S5RP1AX = S5But1Info.iloc[24, [0,2,17,18,19,22,23]]


# Ressources Parcours B

S5R1BX= S5But1InfoB.iloc[12, [0,2,17,18,19,22,23]]
S5R2BX= S5But1InfoB.iloc[13, [0,2,17,18,19,22,23]]
S5R3BX= S5But1InfoB.iloc[14, [0,2,17,18,19,22,23]]
S5R4BX= S5But1InfoB.iloc[15, [0,2,17,18,19,22,23]]
S5R5BX= S5But1InfoB.iloc[16, [0,2,17,18,19,22,23]]
S5R6BX= S5But1InfoB.iloc[17, [0,2,17,18,19,22,23]]
S5R7BX= S5But1InfoB.iloc[18, [0,2,17,18,19,22,23]]
S5R8BX= S5But1InfoB.iloc[19, [0,2,17,18,19,22,23]]
S5R9BX= S5But1InfoB.iloc[20, [0,2,17,18,19,22,23]]

S5R05X= S5But1InfoB.iloc[21, [0,2,17,18,19,22,23]]
S5P1BX = S5But1InfoB.iloc[22, [0,2,17,18,19,22,23]]


# Ressources S6 TRONC COMMUN

S6R601 = S5But1Info.iloc[32, [0,2,17,18,19,22,23]]
S6R602 = S5But1Info.iloc[33, [0,2,17,18,19,22,23]]
S6R603 = S5But1Info.iloc[34, [0,2,17,18,19,22,23]]
S6R604 = S5But1Info.iloc[35, [0,2,17,18,19,22,23]]

# Projets Tutorés 

SProjets = S5But1Info.iloc[36, [0,2,17,18,19,22,23]]
SPortfolio = S5But1Info.iloc[37, [0,2,17,18,19,22,23]]
Stage = S5But1Info.iloc[38, [0,2,17,18,19,22,23]]


# Ressources Parcours A 

S6R1AX = S5But1Info.iloc[36, [0,2,17,18,19,22,23]]
S6R2AX = S5But1Info.iloc[37, [0,2,17,18,19,22,23]]
S6R3AX = S5But1Info.iloc[38, [0,2,17,18,19,22,23]]

# Ressources Parcours B

S6R1BX = S5But1InfoB.iloc[34, [0,2,17,18,19,22,23]]
S6R2BX = S5But1InfoB.iloc[35, [0,2,17,18,19,22,23]]
S6R3BX = S5But1InfoB.iloc[36, [0,2,17,18,19,22,23]]




print(S6R3BX)
