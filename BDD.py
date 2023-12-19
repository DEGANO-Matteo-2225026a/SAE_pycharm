# -*- coding: utf-8 -*-
"""
Created on Thu Dec 14 10:36:22 2023

@author: degan
"""
import json
import sqlite3
import pandas as pd
import numpy as np

# Indique le bon document et la bonne feuille du document
BUT1Info = pd.read_excel('BUT1_INFO_AIX.xlsx', sheet_name=None)
FeuilleInfoBut1 = 'BUT 1'

# Récupère la feuille
S1But1Info = BUT1Info[FeuilleInfoBut1]



"""with sqlite3.connect("SAE.db") as connection:
    cursor = connection.cursor()

# cursor.execute(
#    "CREATE TABLE PROF (cle_prof INTEGER PRIMARY KEY,TITULAIRE BOOLEAN)")

# cursor.execute(
#    "CREATE TABLE RESSOURCE (code_apogee TEXT PRIMARY KEY,cle_prof INTEGER, libelle TEXT, total_cm INTEGER, total_td INTEGER, total_tp INTEGER, HETD INTEGER, FOREIGN KEY (cle_prof) REFERENCES PROF (cle_prof))")

# cursor.execute(
#    "CREATE TABLE ENSEIGNE (cle_prof INTEGER, code_apogee TEXT, nombre_cm INTEGER, nombre_td INTEGER, nombre_tp INTEGER, FOREIGN KEY (cle_prof) REFERENCES PROF(cle_prof), FOREIGN KEY (code_apogee) REFERENCES RESSOURCE(code_apogee))")

#
# Suppretion des tables
# cursor.execute(
#    "DROP TABLE ENSEIGNE;")
connection.close()"""

connexion = sqlite3.connect("SAE.db")

S1codeApogee = S1But1Info.iloc[8:33, 0]
S2codeApogee  = S1But1Info.iloc[34:60, 0]

S1codeApo = np.array(S1codeApogee)
data_to_insert = {'code_apogee': S2codeApogee}


# METTRE LES DONNEES COMME IL FAUT
data_to_insert.to_sql('RESSOURCE', connexion, if_exists='append', index=False)

connexion.commit()
connexion.close()
