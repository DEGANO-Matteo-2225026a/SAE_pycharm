# -*- coding: utf-8 -*-

import sqlite3
import pandas as pd
import numpy as np

import sys
sys.path.insert(0, 'CriteresRessources')  # Ajoute le chemin relatif vers le dossier
from but1 import *
from but2 import *
from but3 import *

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
   "DROP TABLE IF EXISTS BIBLE")

cursor.execute(
   "CREATE TABLE IF NOT EXISTS PROF ("
       "cle_prof INTEGER PRIMARY KEY,"
       "TITULAIRE BOOLEAN)")

cursor.execute(
   "CREATE TABLE IF NOT EXISTS RESSOURCE ("
       "code_apogee TEXT PRIMARY KEY,"
       "cle_prof INTEGER,"
       "libelle TEXT,"
       "total_cm INTEGER,"
       "total_td INTEGER,"
       "total_tp INTEGER,"
       "HETD INTEGER,"
       "HETD_PACOME INTEGER,"
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
       "libelle TEXT,"
       "total_cm INTEGER,"
       "total_td INTEGER,"
       "total_tp INTEGER,"
       "HETD INTEGER,"
       "HETD_PACOME INTEGER,"
       "NOMBRE_GROUPE INTEGER,"
       "BUT TEXT)")

# récuperation des données depuis les import des fichiers disponible dans le dossier CriteresRessources

#fonction d'insertion des données
column_list_RESSOURCE = ['code_apogee', 'libelle', 'total_cm', 'total_td', 'total_tp', 'HETD', 'HETD_PACOME']

def ajouterDonneesToRESSOURCE(ligne_donnees,table_name):
    colonnes = ', '.join(column_list_RESSOURCE)
    valeurs = ', '.join(['?'] * len(column_list_RESSOURCE))
    cursor.execute(f"INSERT INTO {table_name} ({colonnes}) VALUES ({valeurs})", tuple(ligne_donnees))

#liste des insertions
#TODO: mettre toutes les données dans la BDD grâce au code de récuperation des données fait par UGO le prolo
ajouterDonneesToRESSOURCE(S1R101,'RESSOURCE')
ajouterDonneesToRESSOURCE(S1R102,'RESSOURCE')
ajouterDonneesToRESSOURCE(S1R103,'RESSOURCE')
ajouterDonneesToRESSOURCE(S1R104,'RESSOURCE')
ajouterDonneesToRESSOURCE(S1R105,'RESSOURCE')
ajouterDonneesToRESSOURCE(S1R106,'RESSOURCE')
ajouterDonneesToRESSOURCE(S1R107,'RESSOURCE')
ajouterDonneesToRESSOURCE(S1R108,'RESSOURCE')
ajouterDonneesToRESSOURCE(S1R109,'RESSOURCE')
ajouterDonneesToRESSOURCE(S1R110,'RESSOURCE')
ajouterDonneesToRESSOURCE(S1R111,'RESSOURCE')
ajouterDonneesToRESSOURCE(S1R112,'RESSOURCE')
ajouterDonneesToRESSOURCE(S1RL1,'RESSOURCE')


#TODO: ajouter le NOMBRE_GROUPE et BUT dans la fonction de recuperation de donnée pour pouvoir la recuperer ici plus simplement
column_list_BIBLE = ['code_apogee', 'libelle', 'total_cm', 'total_td', 'total_tp', 'HETD', 'HETD_PACOME', 'NOMBRE_GROUPE', 'BUT']

def ajouterDonneesToENSEIGNE(ligne_donnees, GroupeTD, BUT,table_name):
    colonnes = ', '.join(column_list_BIBLE)
    valeurs = ', '.join(['?'] * len(column_list_BIBLE))
    ligne_donnees_complet = ligne_donnees.tolist()
    nouvelles_donnees = [GroupeTD, BUT]
    ligne_donnees_complet.extend(nouvelles_donnees)
    cursor.execute(f"INSERT INTO {table_name} ({colonnes}) VALUES ({valeurs})", tuple(ligne_donnees_complet))


# Supposons que vous avez une liste de valeurs S1R1..
liste_S1R1 = ['S1R101', 'S1R102', 'S1R103', 'S1R104', 'S1R105', 'S1R106', 'S1R107', 'S1R108', 'S1R109', 'S1R110', 'S1R111', 'S1R112']
liste_S2R2 = ['S2R201', 'S2R202', 'S2R203', 'S2R204', 'S2R205', 'S2R206', 'S2R207', 'S2R208', 'S2R209', 'S2R210', 'S2R211', 'S2R212', 'S2R213', 'S2R214']
def remplissageInENSEIGNE(numeroListeSemestre,nombreGroupe,BUT):
    for S1R1 in numeroListeSemestre:
        # Appelez la fonction avec les paramètres appropriés
        ajouterDonneesToENSEIGNE(globals()[S1R1], nombreGroupe, BUT, 'BIBLE')

remplissageInENSEIGNE(liste_S1R1,S1nbGroupeTD, 'BUT1')
remplissageInENSEIGNE(liste_S2R2,S2nbGroupeTD, 'BUT1')

connection.commit()
connection.close()




