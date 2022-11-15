import openpyxl #openpyxl permet de lire, écrire et modifier des fichiers .xlsx
import random #Permet de generer des nombres aléatoire (notes élèves)
def matiere():
    workbook = openpyxl.load_workbook('Nom_Prenom.xlsx') #Ouvre le fichier "Nom_Prenom.xlsx", il est donc possible de modifier les noms des élèves a tout moment.
    sheet_ranges = workbook['Feuille']
    nb_eleve = (sheet_ranges['A2'].value)
    
    
    
    rdm_note = round(random.uniform(0.0, 20.0), 1)
    
 
    

    sheet = workbook.active
    sheet.cell(1, 4).value = 'Note'
    sheet.cell(2, 4).value = rdm_note
    for row in sheet.iter_cols(min_col = 1, max_col = 3, min_raw = 1, max_raw = 5):
        
        
        
        
        
    workbook.save('Mathematiques.xlsx')
    
    
print(matiere())