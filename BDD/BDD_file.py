import sqlite3

#connection a la base de donnée SQLITE
with sqlite3.connect("../SAE.db") as connection:
    cursor = connection.cursor()
# Suppretion des tables pour mettre a jour les données dedans si il y a des modifs
cursor.execute(
   "DROP TABLE IF EXISTS ENSEIGNE")
cursor.execute(
   "DROP TABLE IF EXISTS DONNEEPROF")
cursor.execute(
   "DROP TABLE IF EXISTS RESSOURCE")
cursor.execute(
   "DROP TABLE IF EXISTS PROF")
cursor.execute(
   "DROP TABLE IF EXISTS BIBLE")

# Création des tables
cursor.execute(
   "CREATE TABLE IF NOT EXISTS PROF ("
       "cle_prof INTEGER PRIMARY KEY,"
       "acronyme TEXT,"
       "TITULAIRE BOOLEAN)")

cursor.execute(
   "CREATE TABLE IF NOT EXISTS DONNEEPROF ("
       "Id INTEGER PRIMARY KEY,"
       "Feuille_title TEXT,"
       "MatiereActuelle TEXT,"
       "AlerteProf BOOLEAN,"
       "Intervenant TEXT,"
       "Acronyme TEXT,"
       "Titulaire TEXT,"
       "NombreGroupes INTEGER,"
       "CM FLOAT,"
       "TDNonD INTEGER,"
       "TPD INTEGER,"
       "Test INTEGER)")

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
       "libelle_simple TEXT,"
       "libelle_complet TEXT,"
       "total_cm INTEGER,"
       "total_td INTEGER,"
       "total_tp INTEGER,"
       "HETD INTEGER,"
       "HETD_PACOME INTEGER,"
       "NOMBRE_GROUPE INTEGER,"
       "BUT TEXT)")

# Récuperation des données depuis les import des fichiers disponible dans le dossier CriteresRessources