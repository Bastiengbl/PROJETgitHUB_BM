import openpyxl

l=["Initiation_aux_reseaux_informatiques.xlsx"]

for fichier in l:
    wb = openpyxl.load_workbook(fichier)

    # Sélectionnez la feuille de calcul à utiliser
    sheet = wb['Feuille']

    # Définissez le texte à rechercher
    texte_a_rechercher = "Tokechup"

    # Parcourez chaque ligne de la feuille de calcul
    for row in sheet.rows:
        # Récupérez la cellule de la colonne A
        cell = row[2]
        # Si le texte est trouvé dans la cellule, imprimez la valeur de la cellule adjacente
        if cell.value == texte_a_rechercher:
            cellule_adjacente = sheet.cell(row=cell.row, column=cell.col_idx + 1)
            print(cellule_adjacente.value)
