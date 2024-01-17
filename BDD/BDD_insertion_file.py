# -*- coding: utf-8 -*-

import sqlite3
import pandas as pd
import numpy as np

import sys
sys.path.insert(0, '../CriteresRessources')  # Ajoute le chemin relatif vers le dossier
from but1 import *
from but2 import *
from but3 import *

sys.path.insert(0, '../QuiFaitQuoi')  # Ajoute le chemin relatif vers le dossier
from ExtractionQuiFaitQuoi import *

sys.path.insert(0, '../Planning')  # Ajoute le chemin relatif vers le dossier
from ExtractionPlanning import *


with sqlite3.connect("../SAE.db") as connection:
    cursor = connection.cursor()
    

#creation des tables etc ....
exec(open('BDD_file.py').read())

# PARTIE d'insertion des données

# Liste des Semestre (relié au fichier de recuperation  de donnée)
liste_S1R1 = ['S1R101', 'S1R102', 'S1R103', 'S1R104', 'S1R105', 'S1R106', 'S1R107', 'S1R108', 'S1R109', 'S1R110', 'S1R111', 'S1R112','S1RL1']
liste_S2R2 = ['S2R201', 'S2R202', 'S2R203', 'S2R204', 'S2R205', 'S2R206', 'S2R207', 'S2R208', 'S2R209', 'S2R210', 'S2R211', 'S2R212', 'S2R213', 'S2R214']
liste_S3R3A = ['S3R301', 'S3R302', 'S3R303', 'S3R304', 'S3R305', 'S3R306', 'S3R307', 'S3R308', 'S3R309', 'S3R310', 'S3R311', 'S3R312', 'S3R313', 'S3R314', 'S3AL1']
liste_S4R4A = ['S4R401', 'S4R402', 'S4R403', 'S4R404', 'S4R405', 'S4R406', 'S4R407', 'S4R4A08', 'S4R4A09', 'S4R4A10', 'S4R4A11', 'S4R4A12', 'S4AL1']
liste_S3R3B = ['S3R301', 'S3R302', 'S3R303', 'S3R304', 'S3R305', 'S3R306', 'S3R307', 'S3R308', 'S3R309', 'S3R310', 'S3R311', 'S3R312', 'S3R313', 'S3R314', 'S3R16X', 'S3P1BX']
liste_S4R4B = ['S4R401', 'S4R402', 'S4R403', 'S4R404', 'S4R405', 'S4R406', 'S4R407', 'S4R1BX', 'S4R2BX', 'S4R3BX', 'S4R4BX', 'S4R5BX', 'S4R9BX']
liste_S5R5A = ['S5R501', 'S5R502', 'S5R503', 'S5R1AX', 'S5R2AX', 'S5R3AX', 'S5R4AX', 'S5R5AX', 'S5R6AX', 'S5R7AX', 'S5R8AX', 'S5R9AX', 'S5R10AX', 'S5R11AX', 'S5RC04X']
liste_S6R6A = ['S6R601', 'S6R602', 'S6R603', 'S6R604', 'S6R1AX', 'S6R2AX']
liste_S5R5B = ['S5R501', 'S5R502', 'S5R503', 'S5R1BX', 'S5R2BX', 'S5R3BX', 'S5R4BX', 'S5R5BX', 'S5R6BX', 'S5R7BX', 'S5R8BX', 'S5R9BX', 'S5R05X']
liste_S6R6B = ['S6R601', 'S6R602', 'S6R603', 'S6R604', 'S6R1BX', 'S6R2BX']

# TODO: ajouter le NOMBRE_GROUPE et BUT dans la fonction de recuperation de donnée pour pouvoir la recuperer ici plus simplement
# Code correspondant au TODO

# Liste des colonnes de la tablea BIBLE
column_list_BIBLE = ['code_apogee', 'libelle_simple', 'libelle_complet', 'total_cm', 'total_td', 'total_tp', 'HETD', 'HETD_PACOME', 'NOMBRE_GROUPE', 'BUT']

# Fonction d'insertion qui ajoute les tuples disponible dans ligne_donnees (contient les données sous forme de ligne) et les insert 1 par 1, colonne par colonne.
def ajouterDonneesToBIBLE(ligne_donnees, GroupeTD, BUT,table_name):
    colonnes = ', '.join(column_list_BIBLE)
    valeurs = ', '.join(['?'] * len(column_list_BIBLE))
    libelle_simple = str(ligne_donnees[1]).split(' ')[0]
    ligne_donnees_complet = [ligne_donnees[0]] + [libelle_simple] + ligne_donnees[1:].tolist()
    nouvelles_donnees = [GroupeTD, BUT]
    ligne_donnees_complet.extend(nouvelles_donnees)
    cursor.execute(f"INSERT INTO {table_name} ({colonnes}) VALUES ({valeurs})", tuple(ligne_donnees_complet))


# Fonction de remplissage pour éviter les répetitions de ligne

def remplissageInBIBLE(numeroListeSemestre, nombreGroupe,BUT):
    for S1R1 in numeroListeSemestre:
        ajouterDonneesToBIBLE(globals()[S1R1], nombreGroupe, BUT, 'BIBLE')

remplissageInBIBLE(liste_S1R1, S1nbGroupeTD, 'BUT1')
remplissageInBIBLE(liste_S2R2, S2nbGroupeTD, 'BUT1')
remplissageInBIBLE(liste_S3R3A, S3nbGroupeTDA, 'BUT2')
remplissageInBIBLE(liste_S3R3B, S3nbGroupeTDB, 'BUT2')
remplissageInBIBLE(liste_S4R4A, S4nbGroupeTDA, 'BUT2')
remplissageInBIBLE(liste_S4R4B, S4nbGroupeTDB, 'BUT2')
remplissageInBIBLE(liste_S5R5A, S5nbGroupeTDA, 'BUT3')
remplissageInBIBLE(liste_S5R5B, S5nbGroupeTDB, 'BUT3')
remplissageInBIBLE(liste_S6R6A, S6nbGroupeTDA, 'BUT3')
remplissageInBIBLE(liste_S6R6B, S6nbGroupeTDB, 'BUT3')

# PARTIE QuiFaitQuoi

column_list_DONNEEPROF = ['Feuille_title','MatiereActuelle','AlerteProf','Intervenant','Acronyme','Titulaire','NombreGroupes','CM','TD','TDNonD','TPD','Test']
def ajouterDonneesToDONNEEPROF(ligne_donnees,table_name):
    colonnes = ', '.join(column_list_DONNEEPROF)
    valeurs = ', '.join(['?'] * len(column_list_DONNEEPROF))
    for i in range(len(ligne_donnees)):
        cursor.execute(f"INSERT INTO {table_name} ({colonnes}) VALUES ({valeurs})", tuple(ligne_donnees[i]))

ajouterDonneesToDONNEEPROF(DonneesQuiFaitQuoi,'DONNEEPROF')

# PARTIE Planning

column_list_PLANRESSOURCE = ['Couleur','Ressource','CM','TD','TP','Acronyme']
def ajouterDonneesToPLANRESOURCE(tableau_donnees,table_name):
    colonnes = ', '.join(column_list_PLANRESSOURCE)
    valeurs = ', '.join(['?'] * len(column_list_PLANRESSOURCE))
    for ligne_donnees in tableau_donnees:
        for i in range(len(ligne_donnees)):
            cursor.execute(f"INSERT INTO {table_name} ({colonnes}) VALUES ({valeurs})", tuple(ligne_donnees[i]))

ajouterDonneesToPLANRESOURCE(TableauDonnees[0],'PLANRESSOURCE')

column_list_PLANINFO = ['Semaine','Ressource','TypeCours','TypeSalle']
def ajouterDonneesToPLANINFO(ligne_donnees,table_name):
    colonnes = ', '.join(column_list_PLANINFO)
    valeurs = ', '.join(['?'] * len(column_list_PLANINFO))
    for i in range(len(ligne_donnees)):
        cursor.execute(f"INSERT INTO {table_name} ({colonnes}) VALUES ({valeurs})", tuple(ligne_donnees[i]))

ajouterDonneesToPLANINFO(TableauDonnees[1],'PLANINFO')

connection.commit()
connection.close()




