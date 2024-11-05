import pandas as pd

# Chemin vers le fichier Excel
file_path = '/mnt/data/your_excel_file.xlsx'  # Remplacez par le chemin de votre fichier

# Charger le fichier Excel dans un DataFrame pandas
df = pd.read_excel(file_path)

# Afficher les premières lignes du DataFrame pour vérifier le chargement
print(df.head())

# Créer une structure hiérarchique
hierarchy = {}

# Parcourir chaque ligne du DataFrame et organiser les données
for index, row in df.iterrows():
    province = row['Province']
    region = row['Region']
    distrika = row['Distrika']
    kaomina = row['Kaomina']
    fokontany = row['fokontany']

    if province not in hierarchy:
        hierarchy[province] = {}
    if region not in hierarchy[province]:
        hierarchy[province][region] = {}
    if distrika not in hierarchy[province][region]:
        hierarchy[province][region][distrika] = {}
    if kaomina not in hierarchy[province][region][distrika]:
        hierarchy[province][region][distrika][kaomina] = []
    
    hierarchy[province][region][distrika][kaomina].append(fokontany)

# Afficher la structure hiérarchique
for province, regions in hierarchy.items():
    print(f"Province: {province}")
    for region, distrikas in regions.items():
        print(f"  Region: {region}")
        for distrika, kaominas in distrikas.items():
            print(f"    Distrika: {distrika}")
            for kaomina, fokontanies in kaominas.items():
                print(f"      Kaomina: {kaomina}")
                for fokontany in fokontanies:
                    print(f"        Fokontany: {fokontany}")

# Sauvegarder la structure hiérarchique dans un fichier JSON pour une utilisation future
import json
with open('hierarchy.json', 'w', encoding='utf-8') as f:
    json.dump(hierarchy, f, ensure_ascii=False, indent=4)

print("La hiérarchie des données a été sauvegardée dans 'hierarchy.json'.")
