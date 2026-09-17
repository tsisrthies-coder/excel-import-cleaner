# EXCEL IMPORT CLEANER

Application Windows permettant de nettoyer et de préparer les fichiers Excel avant leur importation dans des systèmes comme KoboToolbox ou E-collect, en résolvant spécifiquement les erreurs du type `e.trim is not a function`.

## Installation (Mode développeur)
1. Assurez-vous d'avoir Python 3.11+ installé.
2. Ouvrez un terminal dans le dossier du projet.
3. Installez les dépendances : `pip install -r requirements.txt`

## Lancement
Exécutez la commande : `python app.py`

## Génération de l'EXE Windows
Double-cliquez simplement sur le fichier `build.bat`. 
Une fois terminé, votre application autonome se trouvera dans le dossier `dist/Excel_Import_Cleaner.exe`. Vous pourrez partager ce fichier à n'importe quel ordinateur Windows sans avoir besoin d'installer Python.

## Utilisation
1. Ouvrez l'application.
2. Cliquez sur **Parcourir** et sélectionnez votre fichier `.xlsx` ou `.xlsm`.
3. Cochez les **feuilles** que vous souhaitez nettoyer.
4. Laissez l'option sur **Automatique (recommandé)**.
5. Cliquez sur **PRÉPARER LE FICHIER**.
6. Un fichier `VOTRE_FICHIER_PRET_A_IMPORTER.xlsx` sera créé dans le même dossier, sans toucher au fichier original.