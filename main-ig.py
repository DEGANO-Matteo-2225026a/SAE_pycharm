def executer_fichiers_selectionnes():
    for widget in page_principale.winfo_children():
        widget.destroy()

    sortie_texte = tk.Text(page_principale, bg="white", fg="black", font=("Helvetica", 10))
    sortie_texte.pack(fill=tk.BOTH, expand=True)

    # Redirection de la sortie standard vers un objet de capture
    capture_sortie = StringIO()
    sys.stdout = capture_sortie

    for fichier, coche in zip(fichiers_selectionnes, fichiers_coche):
        if coche.get():
            if fichier == "bddInsertion.py":
                print("-----------------")
                print("Lancement de bddInsertion.py")
                print("-----------------")
                try:
                    exec(open(fichier).read())
                    print(f"Succès : {fichier}")
                except Exception as e:
                    print(f"Erreur lors de l'exécution du fichier {fichier} : {e}")

            elif fichier == "generationRapport.py":
                print("-----------------")
                print("Lancement de generationRapport.py")
                print("-----------------")
                try:
                    exec(open(fichier).read())
                    print(f"Succès : {fichier}")
                    print("-----------------")
                except Exception as e:
                    print(f"Erreur lors de l'exécution du fichier {fichier} : {e}")
                print("Contenu du fichier rapportErreurs.txt : ")
                print("-----------------")
                try:
                    with open("rapportErreurs.txt", "r") as rapport_file:
                        contenu_rapport = rapport_file.read()
                        print(contenu_rapport)
                except Exception as e:
                    print(f"Erreur lors de la lecture du fichier rapportErreurs.txt : {e}")

            elif fichier == "generationRessources.py":
                print("-----------------")
                print("Lancement de generationRessources.py")
                print("-----------------")
                try:
                    exec(open(fichier).read())
                    print(f"Succès : {fichier}")
                    print("-----------------")
                except Exception as e:
                    print(f"Erreur lors de l'exécution du fichier {fichier} : {e}")

    # Restaurer la sortie standard
    sys.stdout = sys.__stdout__

    # Récupérer le contenu de la sortie capturée
    contenu_resultat = capture_sortie.getvalue()

    # Afficher le contenu dans le widget texte en utilisant la fonction afficher_resultat
    afficher_resultat(sortie_texte, contenu_resultat)

def afficher_resultat(sortie_texte, contenu_resultat):
    sortie_texte.insert("1.0", contenu_resultat)
    fenetre.update()
