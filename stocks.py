import pandas as pd

table1_file = 'Input/valeo.xlsx'

# Vérifier le type de fichier de la table 1
if table1_file.endswith('.csv'):
    df1 = pd.read_csv(table1_file)
elif table1_file.endswith('.xlsx'):
    df1 = pd.read_excel(table1_file)
else:
    raise ValueError("Format de fichier non pris en charge.")

# Charger la table 2 depuis un fichier Excel
df2 = pd.read_excel('ExempleTD/TecDoc.xlsx')

# Spécifier les colonnes de jointure dans chaque DataFrame
colonne_table1 = 'MPN (or OE)'
colonne_table2 = 'MPN'

# Joindre les deux tables en utilisant les colonnes de jointure spécifiées
df_merged = pd.merge(df1, df2, left_on=colonne_table1, right_on=colonne_table2, how='inner')

# Écrire les données complètes dans un nouveau fichier Excel
df_merged.to_excel('fusion-input-tecdoc.xlsx', index=False)

# Ecrire les données dans le tableau excel MIP en remplissant les bons champs
df_destination = pd.DataFrame()
mapping = {'nom_colonne_source': 'nom_colonne_destination', 'nom_colonne_source2': 'nom_colonne_destination2', ...}
for colonne_source, colonne_destination in mapping.items():
    df_destination[colonne_destination] = df_source[colonne_source]

df_destination.to_excel('chemin/vers/le/deuxieme/fichier.xlsx', index=False)












# Target columns -> "SKU","Localized For","Title","Subtitle","Product Description","Additional Info","Group Picture URL","Picture URL 1","Picture URL 2","Picture URL 3","Picture URL 4","UPC","ISBN","EAN","Attribute Name 1","Attribute Value 1","Attribute Name 2","Attribute Value 2","Attribute Name 3","Attribute Value 3","Condition","Condition Description","Total Ship To Home Quantity","Channel ID","Category","Shipping policy","Payment Policy","Return Policy","List Price","VAT Percent","Max Quantity Per Buyer","Strikethrough Price","Store Category Name 1","Store Category Name 2","Include eBay Product Details","Template Name","Compatible Product 1","Compatible Product 2","Compatible Product 3"