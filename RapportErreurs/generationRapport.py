import sqlite3 as sql
import openpyxl as opp

sqlConnection = sql.connect("../SAE.db")
cursor =    sqlConnection.cursor()

def extractionRessources(annee) :
