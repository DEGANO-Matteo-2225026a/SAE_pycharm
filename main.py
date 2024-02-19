# Importez la classe RapportErreurGenerator du fichier generationRapport.py
import sys

from generationRapport import RapportErreurGenerator
from generationRessources import GenerationRessources
from bddInsertion import BddInsertion
from ExtractionPlanning import *
def main():
        # Instanciez la classe RapportErreurGenerator
    rapport_generator = RapportErreurGenerator()
    ressource_generator = GenerationRessources()
    insertion_bdd = BddInsertion()

    # Appelez la méthode run_entire_file() pour exécuter tout le code du fichier generationRapport.py
    rapport_generator.run()
    ressource_generator.run()
    insertion_bdd.run()

# Vérifiez si le script est exécuté en tant que programme principal
if __name__ == "__main__":
    main()