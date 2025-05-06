import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io

# Configuration de la page
st.set_page_config(
    page_title="Convertisseur Bilan Social",
    page_icon="🔄",
    layout="wide"
)

# Titre et introduction
st.title("🔄 Convertisseur de Bilan Social")
st.markdown("""
Cette application vous permet de convertir les données de votre bilan social au format requis pour l'évaluation de la diversité et inclusion.
""")

# Fonction pour créer le modèle Excel
def create_excel_template():
    # Création d'un DataFrame avec les champs nécessaires
    df = pd.DataFrame({
        'Information': [
            'Nom de l\'entreprise',
            'Année',
            'Effectif total',
            'Nombre de femmes',
            'Nombre de cadres',
            'Nombre de femmes cadres',
            'Nombre de salariés en situation de handicap',
            'Nombre de jours travaillés',
            'Nombre de jours d\'absence',
            'Répartition par âge - Moins de 30 ans',
            'Répartition par âge - 30-50 ans',
            'Répartition par âge - Plus de 50 ans',
            'Salaire moyen hommes (€)',
            'Salaire moyen femmes (€)'
        ],
        'Valeur': ['', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        'Unité': ['', '', 'personnes', 'personnes', 'personnes', 'personnes', 'personnes', 'jours', 'jours', 'personnes', 'personnes', 'personnes', '€', '€']
    })
    
    # Création du fichier Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Données', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Données']
        
        # Format pour les en-têtes
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1
        })
        
        # Format pour les nombres
        number_format = workbook.add_format({'num_format': '#,##0'})
        
        # Appliquer les formats
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Ajuster les largeurs de colonnes
        worksheet.set_column('A:A', 40)
        worksheet.set_column('B:B', 20)
        worksheet.set_column('C:C', 20)
        
        # Ajouter des validations
        for row in range(1, len(df) + 1):
            worksheet.write(row, 1, df.iloc[row-1]['Valeur'], number_format)
    
    return output.getvalue()

# Fonction pour convertir les données
def convert_data(df):
    try:
        # Vérification des données requises
        required_fields = ['Information', 'Valeur']
        if not all(field in df.columns for field in required_fields):
            raise ValueError("Le fichier Excel doit contenir les colonnes 'Information' et 'Valeur'")
        
        # Extraction des données
        nom_entreprise = str(df.iloc[0]['Valeur'])
        if not nom_entreprise:
            raise ValueError("Le nom de l'entreprise est requis")
            
        annee = int(df.iloc[1]['Valeur'])
        if annee < 2000 or annee > datetime.now().year:
            raise ValueError(f"L'année doit être comprise entre 2000 et {datetime.now().year}")
        
        # Calcul des indicateurs
        effectif_total = float(df.iloc[2]['Valeur'])
        if effectif_total <= 0:
            raise ValueError("L'effectif total doit être supérieur à 0")
            
        nb_femmes = float(df.iloc[3]['Valeur'])
        if nb_femmes < 0 or nb_femmes > effectif_total:
            raise ValueError("Le nombre de femmes doit être compris entre 0 et l'effectif total")
            
        nb_cadres = float(df.iloc[4]['Valeur'])
        if nb_cadres < 0 or nb_cadres > effectif_total:
            raise ValueError("Le nombre de cadres doit être compris entre 0 et l'effectif total")
            
        nb_femmes_cadres = float(df.iloc[5]['Valeur'])
        if nb_femmes_cadres < 0 or nb_femmes_cadres > nb_cadres:
            raise ValueError("Le nombre de femmes cadres doit être compris entre 0 et le nombre total de cadres")
            
        nb_handicap = float(df.iloc[6]['Valeur'])
        if nb_handicap < 0 or nb_handicap > effectif_total:
            raise ValueError("Le nombre de salariés en situation de handicap doit être compris entre 0 et l'effectif total")
            
        jours_travailles = float(df.iloc[7]['Valeur'])
        if jours_travailles <= 0:
            raise ValueError("Le nombre de jours travaillés doit être supérieur à 0")
            
        jours_absence = float(df.iloc[8]['Valeur'])
        if jours_absence < 0:
            raise ValueError("Le nombre de jours d'absence ne peut pas être négatif")
            
        # Récupération des effectifs par âge
        moins_30 = float(df.iloc[9]['Valeur'])
        entre_30_50 = float(df.iloc[10]['Valeur'])
        plus_50 = float(df.iloc[11]['Valeur'])
        
        # Vérification de la cohérence des effectifs par âge
        total_age = moins_30 + entre_30_50 + plus_50
        if abs(total_age - effectif_total) > 1:  # Tolérance de 1 personne pour les arrondis
            raise ValueError("La somme des effectifs par âge doit correspondre à l'effectif total")
        
        # Calcul des pourcentages
        taux_feminisation = (nb_femmes / effectif_total) * 100
        taux_femmes_cadres = (nb_femmes_cadres / nb_cadres) * 100 if nb_cadres > 0 else 0
        taux_handicap = (nb_handicap / effectif_total) * 100
        
        # Calcul des pourcentages par âge
        moins_30_pct = (moins_30 / effectif_total) * 100
        entre_30_50_pct = (entre_30_50 / effectif_total) * 100
        plus_50_pct = (plus_50 / effectif_total) * 100
        
        # Calcul de l'écart salarial
        salaire_hommes = float(df.iloc[12]['Valeur'])
        salaire_femmes = float(df.iloc[13]['Valeur'])
        if salaire_hommes <= 0 or salaire_femmes <= 0:
            raise ValueError("Les salaires moyens doivent être supérieurs à 0")
        ecart_salaire = ((salaire_hommes - salaire_femmes) / salaire_hommes) * 100
        
        # Calcul du taux d'absentéisme
        taux_absenteisme = (jours_absence / jours_travailles) * 100
        
        # Création du DataFrame de sortie
        output_data = {
            'Indicateur': [
                'nom_entreprise', 'annee', 'taux_feminisation', 'taux_femmes_cadres',
                'taux_handicap', 'ecart_salaire', 'moins_30_ans', 'entre_30_50_ans',
                'plus_50_ans', 'taux_absenteisme'
            ],
            'Valeur': [
                nom_entreprise, annee, taux_feminisation, taux_femmes_cadres,
                taux_handicap, ecart_salaire, moins_30_pct, entre_30_50_pct,
                plus_50_pct, taux_absenteisme
            ]
        }
        
        return pd.DataFrame(output_data)
    
    except ValueError as ve:
        st.error(f"Erreur de validation des données : {str(ve)}")
        return None
    except Exception as e:
        st.error(f"Erreur lors de la conversion des données : {str(e)}")
        return None

# Interface principale
st.markdown("## 📝 Instructions")
st.markdown("""
1. Téléchargez d'abord le modèle Excel
2. Remplissez-le avec les données de votre bilan social
3. Téléchargez le fichier rempli
4. L'application convertira automatiquement les données au format requis
""")

# Bouton pour télécharger le modèle
st.markdown("### 1. Télécharger le modèle")
excel_data = create_excel_template()
st.download_button(
    label="📥 Télécharger le modèle Excel",
    data=excel_data,
    file_name="modele_bilan_social.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Section pour télécharger le fichier rempli
st.markdown("### 2. Télécharger votre fichier rempli")
uploaded_file = st.file_uploader("Choisissez votre fichier Excel rempli", type=['xlsx'])

if uploaded_file is not None:
    try:
        # Lecture du fichier
        df = pd.read_excel(uploaded_file)
        
        # Conversion des données
        df_converted = convert_data(df)
        
        if df_converted is not None:
            st.success("✅ Conversion réussie !")
            
            # Affichage des données converties
            st.markdown("### 3. Données converties")
            st.dataframe(df_converted)
            
            # Bouton pour télécharger le fichier CSV
            csv = df_converted.to_csv(index=False)
            st.download_button(
                label="📥 Télécharger le fichier CSV",
                data=csv,
                file_name=f"donnees_di_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
            
            st.markdown("""
            ### 4. Prochaines étapes
            1. Téléchargez le fichier CSV généré
            2. Importez-le dans l'application d'évaluation D&I
            3. Consultez votre rapport détaillé
            """)
    
    except Exception as e:
        st.error(f"Erreur lors du traitement du fichier : {str(e)}")
        st.error("Veuillez vérifier que le fichier est correctement formaté.")

# Pied de page
st.markdown("---")
st.markdown("""
**Convertisseur Bilan Social** | Développé par Japhet Calixte N'DRI | Version 1.0
""") 