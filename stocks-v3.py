import csv
from openpyxl import load_workbook

# convertir le fichier input depuis excel vers csv
valeo_file = load_workbook(filename = 'Input/valeo.xlsx')
sheet_valeo = valeo_file.active
csv_data_valeo = []
for value in sheet_valeo.iter_rows(values_only=True):
    csv_data_valeo.append(list(value))
with open('Valeo.csv', 'w') as valeo_csv:
    writer = csv.writer(valeo_csv, delimiter=',')
    for line in csv_data_valeo:
        writer.writerow(line)

# convertir le fichier tecdoc depuis excel vers csv
tecdoc_file = load_workbook(filename = 'ExempleTD/TecDoc.xlsx')
sheet_tecdoc = tecdoc_file.active
csv_data_tecdoc = []
for value in sheet_tecdoc.iter_rows(values_only=True):
    csv_data_tecdoc.append(list(value))
with open('Tecdoc.csv', 'w') as tecdoc_csv:
    writer = csv.writer(tecdoc_csv, delimiter=',')
    for line in csv_data_tecdoc:
        writer.writerow(line)

# lire les données dans les csv
valeo_data = []
with open('Valeo.csv', 'r') as valeo_csv:
    reader = csv.DictReader(valeo_csv)
    for row in reader:
        valeo_data.append(row)

tecdoc_data = []
with open('Tecdoc.csv', 'r') as tecdoc_csv:
    reader = csv.DictReader(tecdoc_csv)
    for row in reader:
        tecdoc_data.append(row)

# écrire les données du csv tecdoc dans un dictionnaire
tecdoc_dict = {}
for row in tecdoc_data:
    mpn = row['MPN']
    tecdoc_dict[mpn] = row

# parcourir les données dans le csv valeo pour trouver le champ qui sert de clé de jointure
merged_data = []
for row in valeo_data:
    mpn = row['MPN (or OE)']
    if mpn in tecdoc_dict:
        merged_row = {**row, **tecdoc_dict[mpn]}
        merged_data.append(merged_row)

# fusionner les 2 csv sur le champ "MPN" / "MPN (or OE)"
fieldnames = list(merged_data[0].keys())

with open('fusion-input-tecdoc.csv', 'w', newline='') as fusion_csv:
    writer = csv.DictWriter(fusion_csv, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(merged_data)