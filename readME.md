
# Automatisation des imports fournisseurs

Ce script Python permet de convertir et fusionner des fichiers d'entrée (input, complément d'infos de TecDoc) et d'écrire les données collectées dans un fichier de sortie au format MIP eBay.

  

## Prérequis

Avant d'exécuter le script, assurez-vous d'avoir les éléments suivants :

  

* Python 3 installé sur votre machine.

  

* Les packages openpyxl et csv installés. Vous pouvez les installer en exécutant la commande suivante :

  
>     pip install openpyxl csv

## Utilisation

Placez les fichiers d'entrée (input, complément d'infos de TecDoc) dans les répertoires appropriés :

  

Le fichier d'entrée en format .xlsx (envoi par le fournisseur sur le serveur) doit être placé dans le répertoire "Input".

Le fichier d'entrée "TecDoc.xlsx" doit être placé dans le répertoire "ExempleTD".

Exécutez le script Python "stocks-v3.py" pour convertir et fusionner les fichiers d'entrée, ainsi que pour écrire les données collectées dans le fichier de sortie "MIP.csv" :



> py stocks-v3.py

Le fichier de sortie "MIP.csv" contiendra les données fusionnées et mappées, prêtes à être utilisées pour l'import dans MIP eBay.

  

## Configuration

Vous pouvez personnaliser le script en modifiant le mapping des champs entre les fichiers d'entrée et le fichier de sortie. Le mapping actuel est défini dans la variable mapping du script.

  

Pour ajouter de nouvelles fonctionnalités au script, vous pouvez suivre le modèle existant en ajoutant de nouvelles clés au dictionnaire mapping et en implémentant la logique correspondante dans la boucle de création des données mappées.

  

### Auteur

Ce script a été développé par Charlotte Redier.