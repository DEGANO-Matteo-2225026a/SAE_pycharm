# Importez la classe RapportErreurGenerator du fichier generationRapport.py
import sys

from generationRapport import RapportErreurGenerator
from generationRessources import GenerationRessources
from generationFicheProf import GenerationFicheProf
from bddInsertion import BddInsertion
from ExtractionPlanning import *
def main():
        # Instanciez la classe RapportErreurGenerator
    rapport_generator = RapportErreurGenerator()
    ressource_generator = GenerationRessources()
    prof_generator = GenerationFicheProf()
    insertion_bdd = BddInsertion()

    # Appelez la méthode run_entire_file() pour exécuter tout le code du fichier generationRapport.py
    insertion_bdd.run()
    rapport_generator.run()
    ressource_generator.run()
    prof_generator.run()


# Vérifiez si le script est exécuté en tant que programme principal
if __name__ == "__main__":
    main()