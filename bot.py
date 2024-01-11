import openpyxl

# Charger le fichier Excel
wb = openpyxl.load_workbook(r'C:\Users\ibrah\OneDrive\Documents\Ibrahima\ListeEntreprises.xlsx')
sheet = wb.active

# Ensemble pour stocker les adresses e-mail uniques
unique_email_addresses = set()

# Parcourir les lignes du fichier Excel et extraire les adresses e-mail uniques
for row in sheet.iter_rows(values_only=True):
    for cell_value in row:
        if isinstance(cell_value, str) and '@' in cell_value:
            unique_email_addresses.add(cell_value.strip())

# Convertir l'ensemble en liste pour conserver l'ordre
email_addresses = list(unique_email_addresses)

# Afficher les adresses e-mail uniques extraites
print('Adresses e-mail uniques extraites du fichier Excel :')
for email in email_addresses:
    print(email)

# Afficher le nombre d'adresses e-mail uniques extraites
print('Nombre d\'adresses e-mail uniques extraites :', len(email_addresses))