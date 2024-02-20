import sqlite3


with sqlite3.connect("SAE.db") as connection:
    cursor = connection.cursor()

exec(open('BDD_file.py').read())
exec(open('bddInsertion.py').read())
print("\nResultat requete : \n")
cursor.execute("""
    SELECT dp.Intervenant, b.code_apogee
    FROM DONNEEPROF dp
    INNER JOIN BIBLE b ON dp.MatiereActuelle = b.libelle_complet
    WHERE dp.MatiereActuelle = "R1.01 Initiation au développement"
""")
for result in cursor.fetchall():
    print(result)

