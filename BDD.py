# -*- coding: utf-8 -*-
"""
Created on Thu Dec 14 10:36:22 2023

@author: degan
"""

import sqlite3

print("Gestion de la BDD avec Python et Sqlite3")

with sqlite3.connect("SAE.db") as connection:
    cursor = connection.cursor()

# cursor.execute(
#     "CREATE TABLE PROF (cle_prof INTEGER PRIMARY KEY,TITULAIRE BOOLEAN)")

# cursor.execute(
#     "CREATE TABLE RESSOURCE (code_apogee TEXT PRIMARY KEY,cle_prof INTEGER, libelle TEXT, total_cm INTEGER, total_td INTEGER, total_tp INTEGER, HETD INTEGER, FOREIGN KEY (cle_prof) REFERENCES PROF (cle_prof))")

# cursor.execute(
#     "CREATE TABLE ENSEIGNE (cle_prof INTEGER, code_apogee TEXT, nombre_cm INTEGER, nombre_td INTEGER, nombre_tp INTEGER, FOREIGN KEY (cle_prof) REFERENCES PROF(cle_prof), FOREIGN KEY (code_apogee) REFERENCES RESSOURCE(code_apogee))")

# METTRE LES DONNEES COMME IL FAUT
# cursor.execute(
#     "INSERT INTO contacte(nomComplet, email, telephone) VALUES('Ousseynou DIOP', 'ous@fakemail.com', '763887835')")
# connection.commit()



#
# Suppretion des tables
# cursor.execute(
#     "DROP TABLE RESSOURCE;")





connection.commit()