# Script de Conversion et Fusion de Fichiers

Ce script est utilisé pour convertir des fichiers Excel en fichiers CSV et fusionner les données des fichiers convertis avec d'autres fichiers.

## Prérequis

Avant d'exécuter ce script, assurez-vous d'installer les dépendances suivantes :

- openpyxl : `pip install openpyxl`
q
## Configuration

Le script utilise différentes variables de configuration pour spécifier les dossiers et les fichiers d'entrée/sortie. Vous pouvez les ajuster selon vos besoins en modifiant les valeurs des variables suivantes :

- `input_folder` : Le dossier contenant les fichiers d'entrée.
- `input_converted_folder` : Le dossier de sortie pour les fichiers convertis en CSV.
- `tecdoc_converted_folder` : Le dossier de sortie pour le fichier TecDoc converti en CSV.
- `tecdoc_folder` : Le dossier contenant le fichier TecDoc d'origine au format Excel.
- `tecdoc_file` : Le nom du fichier TecDoc d'origine au format Excel.
- `mip_folder` : Le dossier contenant le fichier MIP d'origine au format Excel.
- `mip_converted_folder` : Le dossier de sortie pour le fichier MIP converti en CSV.
- `mip_file` : Le nom du fichier MIP d'origine au format Excel.
- `fusion_folder` : Le dossier de sortie pour les fichiers de fusion.
- `mip_done_folder` : Le dossier de sortie pour les fichiers MIP finaux.
- `mip_converted_file` : Le nom du fichier MIP converti en CSV.

## Conversion des fichiers d'entrée en CSV

Le script commence par convertir les fichiers d'entrée du dossier spécifié en fichiers CSV. Il utilise la bibliothèque openpyxl pour charger les fichiers Excel, lire les données et les écrire dans des fichiers CSV. Les fichiers convertis sont enregistrés dans le dossier `input_converted_folder`. Le script affiche également un message pour indiquer le nom du fichier converti.

## Conversion du fichier TecDoc en CSV

Ensuite, le script convertit le fichier TecDoc d'origine au format Excel en un fichier CSV. Il utilise la même approche que pour la conversion des fichiers d'entrée. Le fichier CSV converti est enregistré dans le dossier `tecdoc_converted_folder`. Le script affiche un message pour indiquer la réussite de la conversion.

## Conversion du fichier MIP en CSV

De même, le script convertit le fichier MIP d'origine au format Excel en un fichier CSV. Le fichier CSV converti est enregistré dans le dossier `mip_converted_folder`. Le script affiche un message pour indiquer la réussite de la conversion.

## Fusion des fichiers d'entrée avec TecDoc.csv

Le script effectue ensuite la fusion des fichiers d'entrée convertis avec le fichier TecDoc.csv. La clé de jointure pour chaque fichier d'entrée est spécifiée dans le dictionnaire `input_keys`. Le script recherche les correspondances entre les clés de jointure des fichiers d'entrée et celle du fichier TecDoc, puis fusionne les lignes correspondantes. Les fichiers fusionnés sont enregistrés dans le dossier `fusion_folder`. Le script affiche un message pour indiquer la réussite de chaque fusion.

## Écriture des fichiers MIP fusionnés

Le script écrit ensuite les données fusionnées dans les fichiers MIP finaux. Il utilise un dictionnaire de mappage des champs entre les fichiers de fusion et les fichiers MIP. Les fichiers de fusion du dossier `fusion_folder` sont parcourus, et les données sont extraites et écrites dans les fichiers MIP correspondants du dossier `mip_done_folder`. Les champs complexes (SKU, OE, Ktype) sont traités spécifiquement pour obtenir les valeurs correctes. Le script affiche un message pour indiquer la réussite de l'écriture des fichiers MIP.

## Notes supplémentaires

- Assurez-vous d'installer les dépendances requises avant d'exécuter le script.
- Vous pouvez ajouter d'autres fichiers d'entrée avec leurs clés de jointure correspondantes dans le dictionnaire `input_keys`.
- Vous pouvez ajouter d'autres mappages de champs pour les autres fichiers de fusion dans le dictionnaire `field_mapping`.

C'est tout ! Vous pouvez maintenant exécuter le script pour convertir vos fichiers et effectuer la fusion des données. N'hésitez pas à personnaliser le script en fonction de vos besoins spécifiques.


  

### Auteur

Ce script a été développé par Charlotte Redier.