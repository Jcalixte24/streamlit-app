import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell

def create_excel_model():
    # Création des feuilles de calcul
    with pd.ExcelWriter('modele_bilan_social_v2.xlsx', engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Styles
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        instruction_format = workbook.add_format({
            'italic': True,
            'font_color': '#666666',
            'text_wrap': True
        })
        
        percent_format = workbook.add_format({'num_format': '0.00%'})
        number_format = workbook.add_format({'num_format': '#,##0'})
        currency_format = workbook.add_format({'num_format': '#,##0 €'})
        
        # Feuille 1: Instructions
        worksheet_instructions = workbook.add_worksheet('Instructions')
        worksheet_instructions.set_column('A:A', 100)
        
        instructions = [
            "Bienvenue dans le modèle d'évaluation de la Diversité et Inclusion V2",
            "",
            "Ce fichier vous permettra de collecter les données nécessaires pour l'évaluation de votre entreprise.",
            "Veuillez suivre ces étapes :",
            "",
            "1. Remplissez d'abord la feuille 'Données générales' avec les informations de base",
            "2. Complétez la 'Répartition par âge' en respectant les totaux",
            "3. Indiquez les données de rémunération dans la feuille correspondante",
            "4. Complétez les informations sur la formation et le recrutement",
            "5. Les calculs seront effectués automatiquement",
            "",
            "Important :",
            "- Tous les champs numériques doivent être remplis avec des nombres positifs",
            "- Les pourcentages sont calculés automatiquement",
            "- Vérifiez que les totaux correspondent bien à votre effectif total",
            "",
            "Une fois le fichier complété, importez-le dans l'application d'évaluation pour obtenir votre rapport détaillé.",
            "",
            "Pour toute question ou assistance :",
            "Email : support@diversite-inclusion.fr",
            "Téléphone : 01 23 45 67 89"
        ]
        
        for i, instruction in enumerate(instructions):
            worksheet_instructions.write(i, 0, instruction, instruction_format)
        
        # Feuille 2: Données générales
        df_general = pd.DataFrame({
            'Information': [
                'Nom de l\'entreprise',
                'Année',
                'Effectif total',
                'Nombre de femmes',
                'Nombre d\'hommes',
                'Nombre de cadres',
                'Nombre de femmes cadres',
                'Nombre de salariés en situation de handicap',
                'Nombre de jours travaillés',
                'Nombre de jours d\'absence',
                'Nombre de contrats CDI',
                'Nombre de contrats CDD',
                'Nombre de contrats d\'apprentissage',
                'Nombre de contrats de professionnalisation'
            ],
            'Valeur': ['', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
            'Unité': ['', '', 'personnes', 'personnes', 'personnes', 'personnes', 'personnes', 'personnes', 'jours', 'jours', 'contrats', 'contrats', 'contrats', 'contrats']
        })
        
        worksheet_general = workbook.add_worksheet('Données générales')
        worksheet_general.set_column('A:A', 40)
        worksheet_general.set_column('B:B', 20)
        worksheet_general.set_column('C:C', 20)
        
        # Écrire les en-têtes et données
        for col, header in enumerate(df_general.columns):
            worksheet_general.write(0, col, header, header_format)
        
        for row, (info, val, unit) in enumerate(zip(df_general['Information'], df_general['Valeur'], df_general['Unité']), start=1):
            worksheet_general.write(row, 0, info)
            if isinstance(val, (int, float)):
                worksheet_general.write(row, 1, val, number_format)
            else:
                worksheet_general.write(row, 1, val)
            worksheet_general.write(row, 2, unit)
        
        # Validations et formules
        worksheet_general.data_validation('B3:B14', {
            'validate': 'decimal',
            'criteria': '>=',
            'value': 0,
            'input_message': 'Veuillez entrer un nombre positif',
            'error_message': 'La valeur doit être positive'
        })
        
        worksheet_general.write_formula('B5', '=B3-B4', number_format)
        worksheet_general.write_formula('B7', '=B4*0.3', number_format)
        
        # Feuille 3: Répartition par âge
        df_age = pd.DataFrame({
            'Tranche d\'âge': ['< 30 ans', '30-50 ans', '> 50 ans'],
            'Nombre d\'hommes': [0, 0, 0],
            'Nombre de femmes': [0, 0, 0],
            'Total': [0, 0, 0],
            '% du total': ['0%', '0%', '0%']
        })
        
        worksheet_age = workbook.add_worksheet('Répartition par âge')
        worksheet_age.set_column('A:A', 20)
        worksheet_age.set_column('B:E', 20)
        
        # Écrire les données
        for col, header in enumerate(df_age.columns):
            worksheet_age.write(0, col, header, header_format)
        
        for row, (age, hommes, femmes, total, pct) in enumerate(zip(df_age['Tranche d\'âge'], df_age['Nombre d\'hommes'], 
                                                                  df_age['Nombre de femmes'], df_age['Total'], 
                                                                  df_age['% du total']), start=1):
            worksheet_age.write(row, 0, age)
            worksheet_age.write(row, 1, hommes, number_format)
            worksheet_age.write(row, 2, femmes, number_format)
            worksheet_age.write(row, 3, total, number_format)
            worksheet_age.write(row, 4, pct, percent_format)
        
        # Formules
        for row in range(1, 4):
            worksheet_age.write_formula(row, 3, f'=SUM(B{row+1}:C{row+1})')
            worksheet_age.write_formula(row, 4, f'=D{row+1}/Données générales!B3')
        
        # Feuille 4: Rémunérations
        df_remuneration = pd.DataFrame({
            'Catégorie': ['Cadres', 'Non-cadres'],
            'Salaire moyen hommes': [0, 0],
            'Salaire moyen femmes': [0, 0],
            'Écart (%)': [0, 0],
            'Écart en euros': [0, 0]
        })
        
        worksheet_rem = workbook.add_worksheet('Rémunérations')
        worksheet_rem.set_column('A:A', 20)
        worksheet_rem.set_column('B:E', 20)
        
        # Écrire les données
        for col, header in enumerate(df_remuneration.columns):
            worksheet_rem.write(0, col, header, header_format)
        
        for row, (cat, sal_h, sal_f, ecart, ecart_eur) in enumerate(zip(df_remuneration['Catégorie'], 
                                                                       df_remuneration['Salaire moyen hommes'],
                                                                       df_remuneration['Salaire moyen femmes'], 
                                                                       df_remuneration['Écart (%)'],
                                                                       df_remuneration['Écart en euros']), start=1):
            worksheet_rem.write(row, 0, cat)
            worksheet_rem.write(row, 1, sal_h, currency_format)
            worksheet_rem.write(row, 2, sal_f, currency_format)
            worksheet_rem.write(row, 3, ecart, percent_format)
            worksheet_rem.write(row, 4, ecart_eur, currency_format)
        
        # Formules
        for row in range(1, 3):
            worksheet_rem.write_formula(row, 3, f'=(B{row+1}-C{row+1})/B{row+1}')
            worksheet_rem.write_formula(row, 4, f'=B{row+1}-C{row+1}')
        
        # Feuille 5: Formation et Recrutement
        df_formation = pd.DataFrame({
            'Indicateur': [
                'Nombre de formations suivies',
                'Nombre de formations obligatoires',
                'Nombre de formations à l\'initiative du salarié',
                'Budget formation',
                'Nombre de recrutements',
                'Nombre de recrutements internes',
                'Nombre de recrutements externes'
            ],
            'Valeur': [0, 0, 0, 0, 0, 0, 0],
            'Unité': ['formations', 'formations', 'formations', '€', 'recrutements', 'recrutements', 'recrutements']
        })
        
        worksheet_formation = workbook.add_worksheet('Formation et Recrutement')
        worksheet_formation.set_column('A:A', 40)
        worksheet_formation.set_column('B:C', 20)
        
        # Écrire les données
        for col, header in enumerate(df_formation.columns):
            worksheet_formation.write(0, col, header, header_format)
        
        for row, (ind, val, unit) in enumerate(zip(df_formation['Indicateur'], df_formation['Valeur'], df_formation['Unité']), start=1):
            worksheet_formation.write(row, 0, ind)
            worksheet_formation.write(row, 1, val, number_format)
            worksheet_formation.write(row, 2, unit)
        
        # Feuille 6: Calculs automatiques
        df_calculs = pd.DataFrame({
            'Indicateur': [
                'Taux de féminisation global',
                'Taux de femmes cadres',
                'Taux d\'emploi des personnes en situation de handicap',
                'Écart de salaire hommes/femmes (moyenne)',
                'Score d\'équilibre des âges',
                'Taux d\'absentéisme',
                'Taux de CDI',
                'Taux de formation',
                'Taux de recrutement interne'
            ],
            'Formule': [
                '=Nombre de femmes / Effectif total * 100',
                '=Nombre de femmes cadres / Nombre de cadres * 100',
                '=Nombre de salariés en situation de handicap / Effectif total * 100',
                '=MOYENNE(Écart %)',
                'Calcul basé sur la répartition par âge',
                '=Nombre de jours d\'absence / Nombre de jours travaillés * 100',
                '=Nombre de CDI / Effectif total * 100',
                '=Nombre de formations / Effectif total * 100',
                '=Recrutements internes / Total recrutements * 100'
            ],
            'Valeur calculée': [0, 0, 0, 0, 0, 0, 0, 0, 0],
            'Unité': ['%', '%', '%', '%', '%', '%', '%', '%', '%'],
            'Seuil légal': ['-', '-', '6%', '-', '-', '-', '-', '-', '-'],
            'Objectif recommandé': ['40%', '40%', '6%', '<5%', '>70%', '<4%', '>80%', '>5%', '>30%'],
            'Explication': [
                'Pourcentage de femmes dans l\'effectif total',
                'Pourcentage de femmes parmi les postes cadres',
                'Pourcentage de salariés en situation de handicap',
                'Écart moyen de rémunération entre hommes et femmes',
                'Mesure de la diversité des âges',
                'Taux d\'absence par rapport aux jours travaillés',
                'Pourcentage de contrats CDI dans l\'effectif',
                'Taux de participation aux formations',
                'Pourcentage de promotions internes'
            ]
        })
        
        worksheet_calc = workbook.add_worksheet('Calculs automatiques')
        worksheet_calc.set_column('A:A', 40)
        worksheet_calc.set_column('B:B', 40)
        worksheet_calc.set_column('C:F', 20)
        worksheet_calc.set_column('G:G', 60)
        
        # Écrire les données
        for col, header in enumerate(df_calculs.columns):
            worksheet_calc.write(0, col, header, header_format)
        
        for row, (ind, form, val, unit, seuil, obj, expl) in enumerate(zip(df_calculs['Indicateur'], df_calculs['Formule'],
                                                                           df_calculs['Valeur calculée'], df_calculs['Unité'],
                                                                           df_calculs['Seuil légal'], df_calculs['Objectif recommandé'],
                                                                           df_calculs['Explication']), start=1):
            worksheet_calc.write(row, 0, ind)
            worksheet_calc.write(row, 1, form)
            worksheet_calc.write(row, 2, val, percent_format)
            worksheet_calc.write(row, 3, unit)
            worksheet_calc.write(row, 4, seuil)
            worksheet_calc.write(row, 5, obj)
            worksheet_calc.write(row, 6, expl)
        
        # Formules
        worksheet_calc.write_formula('C2', '=Données générales!B4/Données générales!B3')
        worksheet_calc.write_formula('C3', '=Données générales!B7/Données générales!B6')
        worksheet_calc.write_formula('C4', '=Données générales!B8/Données générales!B3')
        worksheet_calc.write_formula('C5', '=AVERAGE(Rémunérations!D2:D3)')
        worksheet_calc.write_formula('C6', '=Données générales!B10/Données générales!B9')
        worksheet_calc.write_formula('C7', '=Données générales!B11/Données générales!B3')
        worksheet_calc.write_formula('C8', '=Formation et Recrutement!B2/Données générales!B3')
        worksheet_calc.write_formula('C9', '=Formation et Recrutement!B6/Formation et Recrutement!B5')
        
        # Protection des feuilles
        for sheet in [worksheet_calc]:
            sheet.protect()

if __name__ == '__main__':
    create_excel_model() 