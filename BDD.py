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



with sqlite3.connect("SAE.db") as connection:
    cursor = connection.cursor()

# suppretion des tables pour mettre a jour les données dedans si il y a des modifs
cursor.execute(
   "DROP TABLE IF EXISTS ENSEIGNE")
cursor.execute(
   "DROP TABLE IF EXISTS RESSOURCE")
cursor.execute(
   "DROP TABLE IF EXISTS PROF")

cursor.execute(
   "CREATE TABLE IF NOT EXISTS PROF ("
       "cle_prof INTEGER PRIMARY KEY,"
       "TITULAIRE BOOLEAN)")

cursor.execute(
   "CREATE TABLE IF NOT EXISTS RESSOURCE ("
       "code_apogee TEXT,"
       "cle_prof INTEGER,"
       "libelle TEXT,"
       "total_cm INTEGER,"
       "total_td INTEGER,"
       "total_tp INTEGER,"
       "HETD INTEGER,"
       "FOREIGN KEY (cle_prof) REFERENCES PROF (cle_prof))")

cursor.execute(
   "CREATE TABLE IF NOT EXISTS ENSEIGNE ("
       "cle_prof INTEGER,"
       "code_apogee TEXT,"
       "nombre_cm INTEGER,"
       "nombre_td INTEGER,"
       "nombre_tp INTEGER,"
       "FOREIGN KEY (cle_prof) REFERENCES PROF(cle_prof),"
       "FOREIGN KEY (code_apogee) REFERENCES RESSOURCE(code_apogee),"
       "PRIMARY KEY (cle_prof,code_apogee))")

cursor.execute(
   "CREATE TABLE IF NOT EXISTS BIBLE ("
       "code_apogee TEXT,"
       "cle_prof INTEGER,"
       "libelle TEXT,"
       "total_cm INTEGER,"
       "total_td INTEGER,"
       "total_tp INTEGER,"
       "HETD INTEGER,"
       "NOMBRE_GROUPE INTEGER)")

#supprime les données actuelle de la table pour ne pas faire de doublon lors de l'insertion des données ci-dessous
cursor.execute(f"DELETE FROM 'ENSEIGNE'")

cursor.execute(f"DELETE FROM 'PROF'")

cursor.execute(f"DELETE FROM 'RESSOURCE'")


# récuperation des données
S1codeApogee = S1But1Info.iloc[10:33, 0]
S2codeApogee  = S1But1Info.iloc[34:60, 0]

#fonction d'insertion des données
def inserer_donnees(data, table_name):
    for index, value in enumerate(data):
        if value != "NULL":
            cursor.execute(f"INSERT INTO {table_name} (code_apogee) VALUES (?)", (value,))

#liste des insertions
#TODO: mettre toutes les données dans la BDD grâce au code de récuperation des données fait par UGO le prolo
inserer_donnees(S1codeApogee, 'RESSOURCE')

connection.commit()
connection.close()




