import sqlite3

#connection a la base de donnée SQLITE
with sqlite3.connect("SAE.db") as connection:
    cursor = connection.cursor()
# Suppression des tables pour mettre a jour les données dedans si il y a des modifs
cursor.execute(
   "DROP TABLE IF EXISTS ENSEIGNE")
cursor.execute(
   "DROP TABLE IF EXISTS DONNEEPROF")
cursor.execute(
   "DROP TABLE IF EXISTS PLANRESSOURCE")
cursor.execute(
   "DROP TABLE IF EXISTS PLANINFO")
cursor.execute(
   "DROP TABLE IF EXISTS RESSOURCE")
cursor.execute(
   "DROP TABLE IF EXISTS PROF")
cursor.execute(
   "DROP TABLE IF EXISTS BIBLE")

# Création des tables :

#table DonneeProf
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
       "CM INTEGER,"
       "TD FLOAT,"
       "TDNonD INTEGER,"
       "TPD INTEGER,"
       "Test INTEGER,"
       "FOREIGN KEY (MatiereActuelle) REFERENCES BIBLE (libelle_complet))")

#table PlanningRessource
cursor.execute(
   "CREATE TABLE IF NOT EXISTS PLANRESSOURCE ("
       "Id INTEGER PRIMARY KEY,"
       "Couleur TEXT,"
       "Ressource TEXT,"
       "CM INTEGER,"
       "TD INTEGER,"
       "TP INTEGER,"
       "Acronyme TEXT,"
       "FOREIGN KEY (Ressource) REFERENCES BIBLE (libelle_simple),"
       "FOREIGN KEY (Acronyme) REFERENCES DONNEEPROF (Acronyme))")

#table PlanningInformation
cursor.execute(
   "CREATE TABLE IF NOT EXISTS PLANINFO ("
       "Id INTEGER PRIMARY KEY,"
       "Semaine DATE,"
       "Ressource TEXT,"
       "TypeCours TEXT,"
       "TypeSalle TEXT,"
       "FOREIGN KEY (Ressource) REFERENCES PLANRESSOURCE (Ressource))")

# table Bible
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
       "BUT TEXT,"
       "groupe TEXT,"
       "PRIMARY KEY (libelle_simple, groupe))")

# Récuperation des données depuis les import des fichiers disponible dans le dossier CriteresRessources
