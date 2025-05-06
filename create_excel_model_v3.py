import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell

def create_excel_model():
    # Création du fichier Excel
    writer = pd.ExcelWriter('modele_bilan_social_v3.xlsx', engine='xlsxwriter')
    workbook = writer.book

    # Définition des styles
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#4472C4',
        'font_color': 'white',
        'border': 1
    })

    instruction_format = workbook.add_format({
        'text_wrap': True,
        'valign': 'top',
        'border': 1
    })

    percent_format = workbook.add_format({
        'num_format': '0.00%',
        'border': 1
    })

    number_format = workbook.add_format({
        'num_format': '#,##0',
        'border': 1
    })

    currency_format = workbook.add_format({
        'num_format': '#,##0.00 €',
        'border': 1
    })

    # Feuille Instructions
    instructions = [
        ['Instructions pour le remplissage du modèle'],
        [''],
        ['1. Informations générales'],
        ['   - Remplissez le nom de l\'entreprise et l\'année d\'évaluation'],
        ['   - Saisissez le nombre total d\'effectifs et la répartition par genre'],
        ['   - Indiquez le nombre de personnes en situation de handicap'],
        [''],
        ['2. Répartition par âge'],
        ['   - Saisissez le nombre d\'employés par tranche d\'âge'],
        ['   - Les totaux et pourcentages sont calculés automatiquement'],
        [''],
        ['3. Rémunérations'],
        ['   - Indiquez les rémunérations moyennes par genre et catégorie'],
        ['   - Les écarts salariaux sont calculés automatiquement'],
        [''],
        ['4. Formation et Recrutement'],
        ['   - Saisissez les données sur la formation et le recrutement'],
        ['   - Les taux sont calculés automatiquement'],
        [''],
        ['5. Calculs automatiques'],
        ['   - Cette feuille est protégée et contient les calculs automatiques'],
        ['   - Ne modifiez pas les formules'],
        [''],
        ['Notes importantes :'],
        ['- Tous les champs numériques doivent être positifs'],
        ['- Les pourcentages sont exprimés en décimal (ex: 0.4 pour 40%)'],
        ['- Les montants sont en euros'],
        ['- Les dates sont au format JJ/MM/AAAA']
    ]

    df_instructions = pd.DataFrame(instructions)
    df_instructions.to_excel(writer, sheet_name='Instructions', index=False, header=False)
    worksheet = writer.sheets['Instructions']
    worksheet.set_column('A:A', 80)
    for row_num, row in enumerate(instructions):
        worksheet.write(row_num, 0, row[0], instruction_format)

    # Feuille Données générales
    general_data = {
        'Informations': ['Nom de l\'entreprise', 'Année d\'évaluation', 'Effectif total', 'Nombre de femmes', 'Nombre d\'hommes', 'Nombre de personnes en situation de handicap'],
        'Valeur': ['', '', 0, 0, 0, 0]
    }
    df_general = pd.DataFrame(general_data)
    df_general.to_excel(writer, sheet_name='Données générales', index=False)
    worksheet = writer.sheets['Données générales']
    worksheet.set_column('A:A', 30)
    worksheet.set_column('B:B', 20)
    
    # Validation des données numériques
    for row in range(3, 6):
        worksheet.data_validation(f'B{row+1}', {'validate': 'decimal', 'criteria': '>=', 'value': 0})

    # Feuille Répartition par âge
    age_data = {
        'Tranche d\'âge': ['< 25 ans', '25-34 ans', '35-44 ans', '45-54 ans', '55-64 ans', '> 64 ans', 'Total'],
        'Effectif': [0, 0, 0, 0, 0, 0, '=SUM(B2:B7)'],
        'Pourcentage': ['=B2/$B$8', '=B3/$B$8', '=B4/$B$8', '=B5/$B$8', '=B6/$B$8', '=B7/$B$8', '=SUM(C2:C7)']
    }
    df_age = pd.DataFrame(age_data)
    df_age.to_excel(writer, sheet_name='Répartition par âge', index=False)
    worksheet = writer.sheets['Répartition par âge']
    worksheet.set_column('A:A', 15)
    worksheet.set_column('B:C', 15)
    
    # Format des pourcentages
    for row in range(1, 8):
        worksheet.write(f'C{row+1}', None, percent_format)

    # Feuille Rémunérations
    remuneration_data = {
        'Catégorie': ['Cadres', 'Ingénieurs', 'Techniciens', 'Administratifs', 'Ouvriers'],
        'Rémunération moyenne hommes': [0, 0, 0, 0, 0],
        'Rémunération moyenne femmes': [0, 0, 0, 0, 0],
        'Écart salarial': ['=(B2-C2)/B2', '=(B3-C3)/B3', '=(B4-C4)/B4', '=(B5-C5)/B5', '=(B6-C6)/B6']
    }
    df_remuneration = pd.DataFrame(remuneration_data)
    df_remuneration.to_excel(writer, sheet_name='Rémunérations', index=False)
    worksheet = writer.sheets['Rémunérations']
    worksheet.set_column('A:A', 15)
    worksheet.set_column('B:E', 20)
    
    # Format des montants et pourcentages
    for row in range(1, 6):
        worksheet.write(f'B{row+1}', None, currency_format)
        worksheet.write(f'C{row+1}', None, currency_format)
        worksheet.write(f'D{row+1}', None, percent_format)

    # Feuille Formation et Recrutement
    formation_data = {
        'Indicateur': [
            'Nombre de stagiaires',
            'Nombre d\'apprentis',
            'Nombre de contrats de professionnalisation',
            'Nombre de CDI',
            'Nombre de CDD',
            'Nombre d\'heures de formation',
            'Nombre de recrutements internes',
            'Nombre de recrutements externes',
            'Nombre d\'employés en temps partiel',
            'Nombre d\'employés en télétravail',
            'Nombre de promotions femmes',
            'Nombre total de promotions'
        ],
        'Valeur': [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        'Unité': ['Personnes', 'Personnes', 'Personnes', 'Personnes', 'Personnes', 'Heures', 'Personnes', 'Personnes', 'Personnes', 'Personnes', 'Personnes', 'Personnes']
    }
    df_formation = pd.DataFrame(formation_data)
    df_formation.to_excel(writer, sheet_name='Formation et Recrutement', index=False)
    worksheet = writer.sheets['Formation et Recrutement']
    worksheet.set_column('A:A', 40)
    worksheet.set_column('B:C', 15)
    
    # Validation des données numériques
    for row in range(1, 13):
        worksheet.data_validation(f'B{row+1}', {'validate': 'decimal', 'criteria': '>=', 'value': 0})

    # Feuille Calculs automatiques
    calculs_data = {
        'Indicateur': [
            'Taux de féminisation',
            'Taux de femmes cadres',
            'Taux de handicap',
            'Écart salarial moyen',
            'Score diversité des âges',
            'Taux de CDI',
            'Taux de formation',
            'Taux de recrutement interne',
            'Taux de temps partiel',
            'Taux de télétravail',
            'Taux de promotion des femmes'
        ],
        'Formule': [
            '=Données générales!B4/Données générales!B3',
            '=Rémunérations!C2/Rémunérations!B2',
            '=Données générales!B6/Données générales!B3',
            '=AVERAGE(Rémunérations!D2:D6)',
            '=1-MAX(ABS(Répartition par âge!C2:C7-0.167))',
            '=Formation et Recrutement!B4/(Formation et Recrutement!B4+Formation et Recrutement!B5)',
            '=Formation et Recrutement!B6/(Données générales!B3*1600)',
            '=Formation et Recrutement!B7/(Formation et Recrutement!B7+Formation et Recrutement!B8)',
            '=Formation et Recrutement!B9/Données générales!B3',
            '=Formation et Recrutement!B10/Données générales!B3',
            '=Formation et Recrutement!B11/Formation et Recrutement!B12'
        ],
        'Valeur': ['', '', '', '', '', '', '', '', '', '', ''],
        'Objectif': ['40%', '40%', '6%', '<5%', '>70%', '>80%', '>5%', '>30%', '<20%', '>20%', '>40%'],
        'Seuil légal': ['-', '-', '6%', '-', '-', '-', '-', '-', '-', '-', '-'],
        'Explication': [
            'Nombre de femmes / Effectif total',
            'Nombre de femmes cadres / Nombre total de cadres',
            'Nombre de personnes en situation de handicap / Effectif total',
            'Moyenne des écarts salariaux par catégorie',
            'Mesure de la diversité des âges (1 - écart max à la répartition idéale)',
            'Nombre de CDI / (Nombre de CDI + Nombre de CDD)',
            'Nombre d\'heures de formation / (Effectif total * 1600)',
            'Nombre de recrutements internes / Nombre total de recrutements',
            'Nombre d\'employés en temps partiel / Effectif total',
            'Nombre d\'employés en télétravail / Effectif total',
            'Nombre de promotions femmes / Nombre total de promotions'
        ]
    }
    df_calculs = pd.DataFrame(calculs_data)
    df_calculs.to_excel(writer, sheet_name='Calculs automatiques', index=False)
    worksheet = writer.sheets['Calculs automatiques']
    worksheet.set_column('A:A', 25)
    worksheet.set_column('B:B', 50)
    worksheet.set_column('C:E', 15)
    worksheet.set_column('F:F', 60)
    
    # Format des pourcentages
    for row in range(1, 12):
        worksheet.write(f'C{row+1}', None, percent_format)

    # Protection des feuilles
    for sheet_name in ['Calculs automatiques']:
        worksheet = writer.sheets[sheet_name]
        worksheet.protect()

    # Sauvegarde du fichier
    writer.close()

if __name__ == '__main__':
    create_excel_model() 