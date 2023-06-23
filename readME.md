# Script de Conversion et Fusion de Fichiers

Ce script Python effectue la conversion de fichiers Excel (.xlsx) en fichiers CSV (.csv) et fusionne ces fichiers avec un fichier de référence.

## Prérequis

Avant d'exécuter ce script, assurez-vous d'avoir les packages Python suivants installés :

- `os`
- `csv`
- `openpyxl`
- `re`

## Configuration

Le script utilise les dossiers et fichiers suivants :

- `input_folder` : dossier contenant les fichiers d'entrée au format Excel
- `input_converted_folder` : dossier de sortie pour les fichiers convertis en CSV
- `tecdoc_converted_folder` : dossier de sortie pour le fichier TecDoc converti en CSV
- `tecdoc_folder` : dossier contenant le fichier TecDoc au format Excel
- `tecdoc_file` : nom du fichier TecDoc au format Excel
- `mip_folder` : dossier contenant le fichier MIP au format Excel
- `mip_converted_folder` : dossier de sortie pour le fichier MIP converti en CSV
- `mip_file` : nom du fichier MIP au format Excel
- `fusion_folder` : dossier de sortie pour les fichiers fusionnés
- `mip_done_folder` : dossier de sortie pour les fichiers MIP fusionnés
- `mip_converted_file` : nom du fichier MIP converti en CSV

Assurez-vous d'adapter ces paramètres selon votre configuration.

## Conversion des fichiers d'entrée en CSV

Le script parcourt le dossier `input_folder` et convertit tous les fichiers Excel en fichiers CSV. Les fichiers convertis sont enregistrés dans le dossier `input_converted_folder`.

## Conversion du fichier TecDoc en CSV

Le fichier TecDoc spécifié dans la configuration est converti en un fichier CSV. Le fichier CSV résultant est enregistré dans le dossier `tecdoc_converted_folder`.

## Conversion du fichier MIP en CSV

Le fichier MIP spécifié dans la configuration est converti en un fichier CSV. Le fichier CSV résultant est enregistré dans le dossier `mip_converted_folder`.

## Fusion des fichiers d'entrée avec TecDoc.csv

Le script fusionne chaque fichier d'entrée avec le fichier TecDoc.csv en utilisant une clé de jointure spécifiée dans le dictionnaire `input_keys`. Les fichiers fusionnés sont enregistrés dans le dossier `fusion_folder`.

## Écriture des fichiers de sortie

Le script utilise un mapping de champs (`field_mapping`) pour écrire les fichiers de sortie MIP. Chaque fichier de fusion est associé à un mapping spécifique dans le dictionnaire `field_mapping`.

Les fichiers MIP fusionnés sont enregistrés dans le dossier `mip_done_folder` avec le préfixe "MIP-".

## Améliorations suggérées

Voici quelques améliorations suggérées pour le script :

- Aligner toutes les paires Attribute Name/Value à la suite.
- Séparer les numéros de Ktype dans le champ Compatible Product 1. Créer un nouveau Compatible Product {i} pour chaque numéro de Ktype trouvé, en utilisant le format "Ktype={Ktype}".
- Séparer les numéros du champ Attribute Value correspondant à "OE" par des '|' comme séparateurs. Limiter.


### Auteur

Ce script a été développé par Charlotte Redier.