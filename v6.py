import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import altair as alt
import io
import plotly.graph_objects as go
import plotly.express as px
from io import StringIO
import base64
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
import tempfile
import os
from datetime import datetime
import kaleido  # Pour la génération d'images Plotly
import pdfkit
import jinja2

# Configuration de la page Streamlit
st.set_page_config(
    page_title="Évaluateur D&I",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Titre et introduction de l'application
st.title("📊 Évaluateur de Diversité et Inclusion en Entreprise")
st.markdown("""
Cette application analyse les indicateurs sociaux d'une entreprise en matière de diversité et inclusion,
et attribue des notes de A à E sur 6 dimensions clés, basées sur des seuils adaptés au secteur énergie/industrie.
""")

# Fonction pour attribuer une note (A-E) selon les seuils définis
def attribuer_note(valeur, seuils, ordre_croissant=True):
    """
    Attribue une note de A à E selon les seuils définis.
    
    Args:
        valeur: La valeur à évaluer
        seuils: Liste de 4 seuils [seuil_A, seuil_B, seuil_C, seuil_D]
        ordre_croissant: Si True, une valeur plus élevée donne une meilleure note
                        Si False, une valeur plus basse donne une meilleure note
    
    Returns:
        Une lettre entre A et E correspondant à la note
    """
    if ordre_croissant:
        if valeur >= seuils[0]:
            return "A"
        elif valeur >= seuils[1]:
            return "B"
        elif valeur >= seuils[2]:
            return "C"
        elif valeur >= seuils[3]:
            return "D"
        else:
            return "E"
    else:  # Ordre décroissant (plus petit = meilleur)
        if valeur <= seuils[0]:
            return "A"
        elif valeur <= seuils[1]:
            return "B"
        elif valeur <= seuils[2]:
            return "C"
        elif valeur <= seuils[3]:
            return "D"
        else:
            return "E"

# Fonction pour convertir les notes en valeurs numériques
def note_vers_chiffre(note):
    """
    Convertit une note de A à E en valeur numérique.
    
    Args:
        note: Une lettre entre A et E
    
    Returns:
        Un entier entre 1 et 5
    """
    conversion = {"A": 5, "B": 4, "C": 3, "D": 2, "E": 1}
    return conversion.get(note, 0)

# Fonction pour convertir un score numérique en note de A à E
def chiffre_vers_note(score):
    """
    Convertit un score numérique en note de A à E.
    
    Args:
        score: Un nombre entre 1 et 5
    
    Returns:
        Une lettre entre A et E
    """
    if score >= 4.5:
        return "A"
    elif score >= 3.5:
        return "B"
    elif score >= 2.5:
        return "C"
    elif score >= 1.5:
        return "D"
    else:
        return "E"

# Fonction pour calculer la répartition équilibrée des âges
def calculer_equilibre_age(moins_30, entre_30_50, plus_50):
    """
    Calcule un score d'équilibre des âges entre 0 et 1.
    Un score de 1 représente une distribution parfaitement équilibrée (33.33% dans chaque catégorie).
    
    Args:
        moins_30, entre_30_50, plus_50: Pourcentages dans chaque tranche d'âge
    
    Returns:
        Un score entre 0 et 1
    """
    # Distribution idéale: 33.33% dans chaque catégorie
    distribution_ideale = 33.33
    
    # Calculer l'écart pour chaque catégorie
    ecart_moins_30 = abs(moins_30 - distribution_ideale)
    ecart_entre_30_50 = abs(entre_30_50 - distribution_ideale)
    ecart_plus_50 = abs(plus_50 - distribution_ideale)
    
    # Calculer l'écart moyen
    ecart_moyen = (ecart_moins_30 + ecart_entre_30_50 + ecart_plus_50) / 3
    
    # Convertir l'écart en score (0 = écart max possible de 66.67, 1 = écart de 0)
    score = 1 - (ecart_moyen / 66.67)
    return score * 100  # Transformer en pourcentage

# Correction des seuils pour le secteur énergie/industrie
seuils = {
    "taux_feminisation": [40, 35, 30, 25],  # % (augmenté pour refléter les objectifs du secteur)
    "taux_femmes_cadres": [35, 30, 25, 20],  # % (ajusté selon les objectifs 2025)
    "taux_handicap": [6, 5, 4, 3],  # % (maintenu avec le seuil légal de 6%)
    "ecart_salaire": [2, 4, 8, 12],  # % (réduit pour plus d'ambition)
    "equilibre_age": [85, 75, 65, 55],  # % (augmenté pour favoriser la diversité des âges)
    "taux_absenteisme": [2.5, 3.5, 4.5, 5.5]  # % (ajusté selon les standards du secteur)
}

# Explication des indicateurs
st.markdown("## 📌 Explication des indicateurs")
st.write("""
1. **Taux de féminisation global** : Pourcentage de femmes dans l'effectif total
2. **Taux de femmes cadres** : Pourcentage de femmes parmi les postes de cadres
3. **Taux d'emploi des personnes en situation de handicap** : Pourcentage de salariés en situation de handicap (seuil légal = 6%)
4. **Écart de salaire hommes/femmes** : Écart moyen en % à poste équivalent (0% = parfaite égalité)
5. **Répartition des effectifs par âge** : Équilibre entre les tranches d'âge (<30 ans, 30-50 ans, >50 ans)
6. **Taux d'absentéisme** : Pourcentage de jours d'absence par rapport au nombre total de jours travaillés
""")

# Affichage des seuils de notation pour chaque indicateur
st.markdown("## 📏 Grilles de notation")

# Créer un dictionnaire pour stocker les grilles de notation
grilles_notation = {}

# Taux de féminisation global
grilles_notation["Taux de féminisation global"] = {
    "A": f"≥ {seuils['taux_feminisation'][0]}%",
    "B": f"{seuils['taux_feminisation'][1]}% à {seuils['taux_feminisation'][0]-0.1}%",
    "C": f"{seuils['taux_feminisation'][2]}% à {seuils['taux_feminisation'][1]-0.1}%",
    "D": f"{seuils['taux_feminisation'][3]}% à {seuils['taux_feminisation'][2]-0.1}%",
    "E": f"< {seuils['taux_feminisation'][3]}%"
}

# Taux de femmes cadres
grilles_notation["Taux de femmes cadres"] = {
    "A": f"≥ {seuils['taux_femmes_cadres'][0]}%",
    "B": f"{seuils['taux_femmes_cadres'][1]}% à {seuils['taux_femmes_cadres'][0]-0.1}%",
    "C": f"{seuils['taux_femmes_cadres'][2]}% à {seuils['taux_femmes_cadres'][1]-0.1}%",
    "D": f"{seuils['taux_femmes_cadres'][3]}% à {seuils['taux_femmes_cadres'][2]-0.1}%",
    "E": f"< {seuils['taux_femmes_cadres'][3]}%"
}

# Taux d'emploi des personnes en situation de handicap
grilles_notation["Taux d'emploi des personnes en situation de handicap"] = {
    "A": f"≥ {seuils['taux_handicap'][0]}%",
    "B": f"{seuils['taux_handicap'][1]}% à {seuils['taux_handicap'][0]-0.1}%",
    "C": f"{seuils['taux_handicap'][2]}% à {seuils['taux_handicap'][1]-0.1}%",
    "D": f"{seuils['taux_handicap'][3]}% à {seuils['taux_handicap'][2]-0.1}%",
    "E": f"< {seuils['taux_handicap'][3]}%"
}

# Écart de salaire hommes/femmes
grilles_notation["Écart de salaire hommes/femmes"] = {
    "A": f"≤ {seuils['ecart_salaire'][0]}%",
    "B": f"{seuils['ecart_salaire'][0]+0.1}% à {seuils['ecart_salaire'][1]}%",
    "C": f"{seuils['ecart_salaire'][1]+0.1}% à {seuils['ecart_salaire'][2]}%",
    "D": f"{seuils['ecart_salaire'][2]+0.1}% à {seuils['ecart_salaire'][3]}%",
    "E": f"> {seuils['ecart_salaire'][3]}%"
}

# Équilibre des âges
grilles_notation["Répartition des effectifs par âge"] = {
    "A": f"Score d'équilibre ≥ {seuils['equilibre_age'][0]}%",
    "B": f"{seuils['equilibre_age'][1]}% à {seuils['equilibre_age'][0]-0.1}%",
    "C": f"{seuils['equilibre_age'][2]}% à {seuils['equilibre_age'][1]-0.1}%",
    "D": f"{seuils['equilibre_age'][3]}% à {seuils['equilibre_age'][2]-0.1}%",
    "E": f"< {seuils['equilibre_age'][3]}%"
}

# Taux d'absentéisme
grilles_notation["Taux d'absentéisme"] = {
    "A": f"≤ {seuils['taux_absenteisme'][0]}%",
    "B": f"{seuils['taux_absenteisme'][0]+0.1}% à {seuils['taux_absenteisme'][1]}%",
    "C": f"{seuils['taux_absenteisme'][1]+0.1}% à {seuils['taux_absenteisme'][2]}%",
    "D": f"{seuils['taux_absenteisme'][2]+0.1}% à {seuils['taux_absenteisme'][3]}%",
    "E": f"> {seuils['taux_absenteisme'][3]}%"
}

# Afficher les grilles de notation dans un format organisé
col1, col2 = st.columns(2)

with col1:
    indicators = list(grilles_notation.keys())[:3]
    for indicator in indicators:
        st.markdown(f"### {indicator}")
        df = pd.DataFrame(list(grilles_notation[indicator].items()), columns=['Note', 'Critère'])
        st.table(df)

with col2:
    indicators = list(grilles_notation.keys())[3:]
    for indicator in indicators:
        st.markdown(f"### {indicator}")
        df = pd.DataFrame(list(grilles_notation[indicator].items()), columns=['Note', 'Critère'])
        st.table(df)

# Méthode d'entrée des données
st.markdown("## 📝 Entrée des données")
methode = st.radio("Choisissez la méthode d'entrée des données:", 
                  ["Saisie manuelle", "Téléchargement de fichier CSV/Excel"])

# Variables pour stocker les données saisies
indicateurs = {}

if methode == "Saisie manuelle":
    # Créer deux colonnes pour l'entrée des données
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Données générales")
        nom_entreprise = st.text_input("Nom de l'entreprise", "EDF SA")
        annee = st.number_input("Année", min_value=2000, max_value=2030, value=2022)
        
        st.subheader("Mixité et égalité professionnelle")
        indicateurs["taux_feminisation"] = st.number_input("Taux de féminisation global (%)", 
                                                         min_value=0.0, max_value=100.0, value=30.0)
        indicateurs["taux_femmes_cadres"] = st.number_input("Taux de femmes cadres (%)", 
                                                          min_value=0.0, max_value=100.0, value=28.0)
        indicateurs["ecart_salaire"] = st.number_input("Écart de salaire hommes/femmes (%)", 
                                                      min_value=0.0, max_value=50.0, value=5.0)
    
    with col2:
        st.subheader("Inclusion et diversité")
        indicateurs["taux_handicap"] = st.number_input("Taux d'emploi des personnes en situation de handicap (%)", 
                                                    min_value=0.0, max_value=20.0, value=5.5)
        
        st.subheader("Répartition par âge")
        moins_30 = st.number_input("Effectifs < 30 ans (%)", 
                                  min_value=0.0, max_value=100.0, value=15.0)
        entre_30_50 = st.number_input("Effectifs 30-50 ans (%)", 
                                     min_value=0.0, max_value=100.0, value=45.0)
        plus_50 = st.number_input("Effectifs > 50 ans (%)", 
                                min_value=0.0, max_value=100.0, value=40.0)
        
        # Vérifier que la somme des pourcentages est égale à 100%
        sum_age = moins_30 + entre_30_50 + plus_50
        if abs(sum_age - 100) > 0.01:  # Tolérance de 0.01%
            st.warning(f"La somme des pourcentages par âge doit être égale à 100% (actuellement {sum_age:.2f}%)")
        
        # Calculer l'équilibre des âges
        indicateurs["equilibre_age"] = calculer_equilibre_age(moins_30, entre_30_50, plus_50)
        
        st.subheader("Climat social")
        indicateurs["taux_absenteisme"] = st.number_input("Taux d'absentéisme (%)", 
                                                        min_value=0.0, max_value=20.0, value=4.2)

elif methode == "Téléchargement de fichier CSV/Excel":
    # Template de fichier à télécharger
    st.subheader("Téléchargez un modèle de fichier")
    
    # Créer un modèle de fichier CSV
    template_data = {
        "Indicateur": [
            "nom_entreprise", "annee", "taux_feminisation", "taux_femmes_cadres", 
            "ecart_salaire", "taux_handicap", "moins_30_ans", "entre_30_50_ans", 
            "plus_50_ans", "taux_absenteisme"
        ],
        "Valeur": ["EDF SA", 2022, 30.0, 28.0, 5.0, 5.5, 15.0, 45.0, 40.0, 4.2]
    }
    template_df = pd.DataFrame(template_data)
    
    # Convertir le DataFrame en CSV et créer un lien de téléchargement
    csv = template_df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="modele_indicateurs_di.csv">Télécharger le modèle CSV</a>'
    st.markdown(href, unsafe_allow_html=True)
    
    # Option pour télécharger un fichier Excel
    excel_buffer = io.BytesIO()
    template_df.to_excel(excel_buffer, index=False)
    excel_data = excel_buffer.getvalue()
    b64_excel = base64.b64encode(excel_data).decode()
    href_excel = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}" download="modele_indicateurs_di.xlsx">Télécharger le modèle Excel</a>'
    st.markdown(href_excel, unsafe_allow_html=True)
    
    # Téléchargement du fichier par l'utilisateur
    st.subheader("Importez votre fichier CSV ou Excel")
    uploaded_file = st.file_uploader("Choisir un fichier", type=['csv', 'xlsx', 'xls'])
    
    if uploaded_file is not None:
        # Déterminer le type de fichier et le lire
        try:
            if uploaded_file.name.endswith('.csv'):
                data = pd.read_csv(uploaded_file)
            else:  # Excel
                data = pd.read_excel(uploaded_file)
            
            # Vérifier le format du fichier
            expected_columns = set(["Indicateur", "Valeur"])
            if not expected_columns.issubset(set(data.columns)):
                st.error("Le fichier doit contenir les colonnes 'Indicateur' et 'Valeur'")
            else:
                # Convertir en dictionnaire
                data_dict = dict(zip(data['Indicateur'], data['Valeur']))
                
                # Extraire les informations
                nom_entreprise = data_dict.get('nom_entreprise', "Entreprise inconnue")
                annee = int(data_dict.get('annee', 2022))
                
                # Remplir les indicateurs
                indicateurs["taux_feminisation"] = float(data_dict.get('taux_feminisation', 0))
                indicateurs["taux_femmes_cadres"] = float(data_dict.get('taux_femmes_cadres', 0))
                indicateurs["ecart_salaire"] = float(data_dict.get('ecart_salaire', 0))
                indicateurs["taux_handicap"] = float(data_dict.get('taux_handicap', 0))
                
                # Calculer l'équilibre des âges
                moins_30 = float(data_dict.get('moins_30_ans', 0))
                entre_30_50 = float(data_dict.get('entre_30_50_ans', 0))
                plus_50 = float(data_dict.get('plus_50_ans', 0))
                
                # Vérifier que la somme des pourcentages est égale à 100%
                sum_age = moins_30 + entre_30_50 + plus_50
                if abs(sum_age - 100) > 0.01:  # Tolérance de 0.01%
                    st.warning(f"La somme des pourcentages par âge doit être égale à 100% (actuellement {sum_age:.2f}%)")
                
                indicateurs["equilibre_age"] = calculer_equilibre_age(moins_30, entre_30_50, plus_50)
                indicateurs["taux_absenteisme"] = float(data_dict.get('taux_absenteisme', 0))
                
                st.success("Données importées avec succès!")
                
                # Afficher les données importées
                st.subheader("Données importées")
                st.write(f"**Entreprise:** {nom_entreprise}")
                st.write(f"**Année:** {annee}")
                
                for key, value in indicateurs.items():
                    st.write(f"**{key}:** {value}")
                
        except Exception as e:
            st.error(f"Erreur lors de la lecture du fichier: {e}")

# Bouton pour lancer l'évaluation
if st.button("Évaluer", type="primary") and indicateurs:
    st.markdown("## 📊 Résultats de l'évaluation")
    
    # Calculer les notes pour chaque indicateur
    resultats = {
        "Taux de féminisation global": attribuer_note(indicateurs["taux_feminisation"], seuils["taux_feminisation"]),
        "Taux de femmes cadres": attribuer_note(indicateurs["taux_femmes_cadres"], seuils["taux_femmes_cadres"]),
        "Taux d'emploi handicap": attribuer_note(indicateurs["taux_handicap"], seuils["taux_handicap"]),
        "Écart de salaire H/F": attribuer_note(indicateurs["ecart_salaire"], seuils["ecart_salaire"], False),
        "Équilibre des âges": attribuer_note(indicateurs["equilibre_age"], seuils["equilibre_age"]),
        "Taux d'absentéisme": attribuer_note(indicateurs["taux_absenteisme"], seuils["taux_absenteisme"], False)
    }
    
    # Convertir les notes en valeurs numériques
    notes_numeriques = {cle: note_vers_chiffre(note) for cle, note in resultats.items()}
    
    # Calculer le score global
    score_global = sum(notes_numeriques.values()) / len(notes_numeriques)
    note_globale = chiffre_vers_note(score_global)
    
    # Préparation des données pour l'affichage
    df_resultats = pd.DataFrame({
        "Indicateur": list(resultats.keys()),
        "Note": list(resultats.values()),
        "Score": list(notes_numeriques.values()),
        "Valeur": [
            f"{indicateurs['taux_feminisation']:.1f}%",
            f"{indicateurs['taux_femmes_cadres']:.1f}%",
            f"{indicateurs['taux_handicap']:.1f}%",
            f"{indicateurs['ecart_salaire']:.1f}%",
            f"{indicateurs['equilibre_age']:.1f}%",
            f"{indicateurs['taux_absenteisme']:.1f}%"
        ]
    })
    
    # Définir des couleurs pour chaque note
    couleurs_notes = {
        "A": "#4CAF50",  # Vert
        "B": "#8BC34A",  # Vert clair
        "C": "#FFC107",  # Jaune
        "D": "#FF9800",  # Orange
        "E": "#F44336"   # Rouge
    }
    
    # Ajouter une colonne de couleurs
    df_resultats["Couleur"] = df_resultats["Note"].map(couleurs_notes)
    
    # Afficher le score global
    col1, col2 = st.columns([1, 3])
    
    with col1:
        # Créer un indicateur visuel pour la note globale
        fig = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=score_global,
            domain={'x': [0, 1], 'y': [0, 1]},
            title={'text': f"Score Global: {note_globale}", 'font': {'size': 24}},
            gauge={
                'axis': {'range': [0, 5], 'tickwidth': 1, 'tickcolor': "darkblue"},
                'bar': {'color': couleurs_notes.get(note_globale, "#888888")},
                'steps': [
                    {'range': [0, 1.5], 'color': "#F44336"},
                    {'range': [1.5, 2.5], 'color': "#FF9800"},
                    {'range': [2.5, 3.5], 'color': "#FFC107"},
                    {'range': [3.5, 4.5], 'color': "#8BC34A"},
                    {'range': [4.5, 5], 'color': "#4CAF50"}
                ],
                'threshold': {
                    'line': {'color': "red", 'width': 4},
                    'thickness': 0.75,
                    'value': score_global
                }
            }
        ))
        
        fig.update_layout(
            height=300,
            margin=dict(l=20, r=20, t=50, b=20),
        )
        
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        # Créer un graphique radar pour visualiser les scores par dimension
        categories = df_resultats["Indicateur"].tolist()
        fig = go.Figure()
        
        fig.add_trace(go.Scatterpolar(
            r=df_resultats["Score"].tolist(),
            theta=categories,
            fill='toself',
            name='Scores par dimension',
            line_color='rgba(32, 128, 255, 0.8)',
            fillcolor='rgba(32, 128, 255, 0.3)'
        ))
        
        fig.update_layout(
            polar=dict(
                radialaxis=dict(
                    visible=True,
                    range=[0, 5]
                )
            ),
            showlegend=False,
            height=300,
            margin=dict(l=70, r=70, t=20, b=20),
        )
        
        st.plotly_chart(fig, use_container_width=True)
    
    # Afficher le tableau des résultats détaillés
    st.subheader("Détail des notes par indicateur")
    
    # Créer une fonction de mise en forme pour le tableau
    def highlight_note(s):
        return [f'background-color: {couleurs_notes.get(s["Note"], "#888888")}; color: white; font-weight: bold' 
                if col == "Note" else '' for col in s.index]
    
    # Afficher le DataFrame avec mise en forme
    st.dataframe(
        df_resultats[["Indicateur", "Valeur", "Note", "Score"]].style.apply(highlight_note, axis=1),
        use_container_width=True,
        hide_index=True
    )
    
    # Graphique à barres des scores par indicateur
    st.subheader("Comparaison des scores par indicateur")
    
    # Trier le DataFrame par score, du plus élevé au plus bas
    df_sorted = df_resultats.sort_values("Score", ascending=False)
    
    # Créer un graphique à barres avec Plotly
    fig = px.bar(
        df_sorted,
        x="Indicateur",
        y="Score",
        color="Note",
        color_discrete_map=couleurs_notes,
        text="Note",
        labels={"Score": "Score (1-5)", "Indicateur": ""},
        height=400
    )
    
    fig.update_layout(
        xaxis_tickangle=-45,
        yaxis=dict(range=[0, 5.5]),
        margin=dict(l=20, r=20, t=20, b=80),
    )
    
    fig.update_traces(textposition='outside')
    
    st.plotly_chart(fig, use_container_width=True)
    
    # Résumé et recommandations
    st.markdown("## 📝 Analyse et recommandations")
    
    # Identifier les points forts (notes A et B)
    points_forts = df_resultats[df_resultats["Note"].isin(["A", "B"])]["Indicateur"].tolist()
    
    # Identifier les points à améliorer (notes D et E)
    points_amelioration = df_resultats[df_resultats["Note"].isin(["D", "E"])]["Indicateur"].tolist()
    
    # Générer les recommandations
    if len(points_forts) > 0:
        st.markdown("### Points forts")
        for point in points_forts:
            st.markdown(f"✅ **{point}**: Performance solide, à maintenir")
    
    if len(points_amelioration) > 0:
        st.markdown("### Points à améliorer en priorité")
        for point in points_amelioration:
            if point == "Taux de féminisation global":
                st.markdown(f"🔍 **{point}**: Mettre en place des actions pour augmenter le recrutement de femmes, notamment dans les métiers techniques")
            elif point == "Taux de femmes cadres":
                st.markdown(f"🔍 **{point}**: Développer des programmes de mentorat et de promotion des femmes vers les postes de cadres")
            elif point == "Taux d'emploi handicap":
                st.markdown(f"🔍 **{point}**: Renforcer la politique de recrutement et d'aménagement des postes pour atteindre le seuil légal de 6%")
            elif point == "Écart de salaire H/F":
                st.markdown(f"🔍 **{point}**: Mettre en place une revue systématique des rémunérations et un plan de rattrapage salarial")
            elif point == "Équilibre des âges":
                st.markdown(f"🔍 **{point}**: Diversifier les recrutements pour équilibrer la pyramide des âges et favoriser le transfert de compétences")
            elif point == "Taux d'absentéisme":
                st.markdown(f"🔍 **{point}**: Analyser les causes profondes et mettre en place des actions d'amélioration de la qualité de vie au travail")
    
    # Points intermédiaires (note C)
    points_intermediaires = df_resultats[df_resultats["Note"] == "C"]["Indicateur"].tolist()
    if len(points_intermediaires) > 0:
        st.markdown("### Points à consolider")
        for point in points_intermediaires:
            st.markdown(f"🔄 **{point}**: Performance moyenne, des progrès sont encore possibles")
    
    # Conclusion générale
    st.markdown("### Conclusion générale")
    if note_globale in ["A", "B"]:
        st.markdown(f"""
        **Avec une note globale de {note_globale} (score {score_global:.2f}/5)**, l'entreprise démontre un engagement solide en matière de diversité et d'inclusion.
        Les bonnes pratiques en place méritent d'être valorisées et partagées.
        """)
    elif note_globale == "C":
        st.markdown(f"""
        **Avec une note globale de {note_globale} (score {score_global:.2f}/5)**, l'entreprise présente des résultats mitigés en matière de diversité et d'inclusion.
        Des progrès significatifs sont encore nécessaires pour atteindre l'excellence dans ce domaine.
        """)
    else:
        st.markdown(f"""
        **Avec une note globale de {note_globale} (score {score_global:.2f}/5)**, l'entreprise présente des performances insuffisantes en matière de diversité et d'inclusion.
        Un plan d'action ambitieux et global est nécessaire pour améliorer ces résultats.
        """)
    
    # Référence aux objectifs de développement durable
    st.markdown("### Lien avec les Objectifs de Développement Durable (ODD)")
    st.markdown("""
    Cette évaluation s'inscrit dans le cadre des Objectifs de Développement Durable des Nations Unies, en particulier :
    - **ODD 5** : Égalité entre les sexes
    - **ODD 8** : Travail décent et croissance économique
    - **ODD 10** : Réduction des inégalités
    """)
    
    # Possibilité de télécharger le rapport
    st.markdown("## 📥 Téléchargement du rapport")
    
    # Préparation du rapport au format CSV
    rapport_csv = df_resultats.to_csv(index=False)
    st.download_button(
        label="Télécharger le rapport (CSV)",
        data=rapport_csv,
        file_name=f"rapport_di_{nom_entreprise}_{annee}.csv",
        mime="text/csv",
    )
    
    # Préparation du rapport au format Excel
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Feuille des résultats
        df_resultats.to_excel(writer, sheet_name='Résultats', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Résultats']
        
        # Formats pour les cellules
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#007BFF',
            'color': 'white',
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        
        # Appliquer le format d'en-tête
        for col_num, value in enumerate(df_resultats.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Ajuster la largeur des colonnes
        for i, col in enumerate(df_resultats.columns):
            column_width = max(df_resultats[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, column_width)
        
        # Ajouter une feuille pour les informations générales
        info_data = {
            'Information': [
                'Entreprise',
                'Année',
                'Note globale',
                'Score global',
                'Date d\'évaluation'
            ],
            'Valeur': [
                nom_entreprise,
                annee,
                note_globale,
                f"{score_global:.2f}/5",
                pd.Timestamp.now().strftime("%d/%m/%Y")
            ]
        }
        pd.DataFrame(info_data).to_excel(writer, sheet_name='Informations', index=False)
    
    excel_data = buffer.getvalue()
    st.download_button(
        label="Télécharger le rapport (Excel)",
        data=excel_data,
        file_name=f"rapport_di_{nom_entreprise}_{annee}.xlsx",
        mime="application/vnd.ms-excel",
    )
    
    # Ajouter un exemple de données EDF basé sur le document fourni
    st.markdown("## 🔍 Exemple : Données EDF 2022")
    st.markdown("""
    Selon le bilan social d'EDF SA pour 2022, les indicateurs clés sont les suivants :
    - **Taux de féminisation** : 30%
    - **Taux de femmes cadres** : 28%
    - **Taux d'emploi des personnes en situation de handicap** : en progression, mais encore sous le seuil légal
    - **Écart de salaire hommes/femmes** : des progrès notables mais des écarts persistent
    - **Répartition par âge** : âge moyen de 42,5 ans, avec une concentration importante dans les tranches supérieures
    - **Taux d'absentéisme** : légèrement au-dessus de la moyenne du secteur
    
    EDF s'est fixé des objectifs ambitieux pour 2025 (33% de femmes à tous les niveaux) et 2030 (36 à 40% de femmes),
    avec un plan d'action pour féminiser notamment les métiers techniques et SI.
    """)

# Pied de page
st.markdown("---")
st.markdown("""
**Diversité & Inclusion Analytics** | Développé avec Streamlit | Basé sur les référentiels du secteur énergie/industrie
""")

# Ajout d'informations sur la méthodologie dans la barre latérale
with st.sidebar:
    st.header("À propos")
    st.markdown("""
    Cet outil d'évaluation est basé sur les pratiques et standards du secteur énergie/industrie.
    Les seuils utilisés pour l'évaluation sont définis à partir de benchmarks sectoriels
    et des exigences légales en vigueur.
    
    **Sources**:
    - Bilans sociaux des entreprises du CAC40
    - Rapports de performance extra-financière
    - Études sectorielles
    - Cadre légal (notamment la loi Copé-Zimmermann, loi Avenir professionnel)
    
    **Développé par**: Japhet Calixte N'DRI
    """)
    
    # Affichage des références
    st.header("Références")
    st.markdown("""
    - Bruna, M. G. (2011). Diversité dans l'entreprise : d'impératif éthique à levier de créativité.
    - Arreola, F., & Sachet Milliat, A. (2022). Question(s) de diversité et inclusion dans l'emploi : nouvelles perspectives.
    - Thomas, D. A., & Ely, R. J. (1996). Making differences matter: A new paradigm for managing diversity.
    """)
    
    # Ajouter le logo EDF
    st.markdown("### Étude de cas: D&I en entreprise ")
   # Modifier l'URL de l'image EDF (l'ancienne URL Wikimedia était instable)
st.image("https://www.bing.com/images/search?view=detailV2&ccid=vFt7rua0&id=E2C879114DF30A4FEE64FA724ECECCCCF70CA783&thid=OIP.vFt7rua0kv5VXUJykDV4TQHaE8&mediaurl=https%3a%2f%2fgroupemenway.com%2fwp-content%2fuploads%2f2023%2f02%2fmodern-companies-encourage-cultural-diversity-in-t-2022-02-22-14-06-52-utc-scaled.jpg&cdnurl=https%3a%2f%2fth.bing.com%2fth%2fid%2fR.bc5b7baee6b492fe555d42729035784d%3frik%3dg6cM98zMzk5y%252bg%26pid%3dImgRaw%26r%3d0&exph=1707&expw=2560&q=divers%c3%a9+et+inclusion&simid=608006326097945224&FORM=IRPRST&ck=872819FAE7187BE2C28F9395FD7298CC&selectedIndex=1&itb=0.", width=150)

# Styles PDF
styles = getSampleStyleSheet()
title_style = ParagraphStyle(
    'CustomTitle',
    parent=styles['Heading1'],
    fontSize=24,
    spaceAfter=30,
    textColor=colors.HexColor('#4472C4')
)
heading_style = ParagraphStyle(
    'CustomHeading',
    parent=styles['Heading2'],
    fontSize=16,
    spaceAfter=12,
    textColor=colors.HexColor('#4472C4')
)

def prepare_data_for_pdf(resultats, points_forts, axes_amelioration, note_globale, score_global, indicateurs):
    """
    Prépare les données pour la génération du PDF.
    """
    try:
        # Mapping des noms d'indicateurs
        mapping_indicateurs = {
            "Taux de féminisation global": "taux_feminisation",
            "Taux de femmes cadres": "taux_femmes_cadres",
            "Taux d'emploi handicap": "taux_handicap",
            "Écart de salaire H/F": "ecart_salaire",
            "Équilibre des âges": "equilibre_age",
            "Taux d'absentéisme": "taux_absenteisme"
        }

        return {
            "resultats": {
                k: {
                    "note": v,
                    "valeur": float(indicateurs.get(mapping_indicateurs[k], 0)),
                    "seuils": seuils.get(mapping_indicateurs[k], [0, 0, 0, 0])
                } for k, v in resultats.items()
            },
            "points_forts": points_forts if points_forts else [],
            "axes_amelioration": axes_amelioration if axes_amelioration else [],
            "note_globale": note_globale,
            "score_global": float(score_global)
        }
    except Exception as e:
        raise Exception(f"Erreur lors de la préparation des données : {str(e)}")

def generate_pdf(data, company_name, year):
    """
    Génère un rapport PDF avec les résultats de l'évaluation en utilisant pdfkit.
    """
    try:
        # Template HTML avec styles améliorés
        template = """
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>Rapport D&I - {{nom_entreprise}}</title>
            <style>
                body { 
                    font-family: Arial, sans-serif; 
                    margin: 40px;
                    color: #333;
                    line-height: 1.6;
                }
                .header { 
                    text-align: center; 
                    margin-bottom: 30px;
                    padding: 20px;
                    background-color: #f8f9fa;
                    border-radius: 8px;
                    border: 2px solid #1E3A8A;
                }
                .section { 
                    margin: 20px 0; 
                    padding: 20px; 
                    border: 1px solid #ddd;
                    border-radius: 8px;
                    background-color: white;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                }
                table { 
                    width: 100%; 
                    border-collapse: collapse; 
                    margin: 20px 0;
                    background-color: white;
                }
                th, td { 
                    padding: 12px; 
                    border: 1px solid #ddd; 
                    text-align: left;
                }
                th { 
                    background-color: #1E3A8A; 
                    color: white;
                    font-weight: bold;
                }
                tr:nth-child(even) {
                    background-color: #f8f9fa;
                }
                .note-A { 
                    color: #4CAF50;
                    font-weight: bold;
                }
                .note-B { 
                    color: #8BC34A;
                    font-weight: bold;
                }
                .note-C { 
                    color: #FFC107;
                    font-weight: bold;
                }
                .note-D { 
                    color: #FF9800;
                    font-weight: bold;
                }
                .note-E { 
                    color: #F44336;
                    font-weight: bold;
                }
                ul {
                    list-style-type: none;
                    padding-left: 0;
                }
                li {
                    margin: 10px 0;
                    padding-left: 20px;
                    position: relative;
                }
                li:before {
                    content: "•";
                    position: absolute;
                    left: 0;
                    color: #1E3A8A;
                }
                h1, h2, h3 {
                    color: #1E3A8A;
                }
                .footer {
                    text-align: center;
                    font-size: 0.9em;
                    color: #666;
                    margin-top: 30px;
                    padding-top: 20px;
                    border-top: 1px solid #ddd;
                }
                .nutriscore {
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    margin: 20px 0;
                }
                .nutriscore-letter {
                    width: 80px;
                    height: 80px;
                    border-radius: 50%;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    font-size: 32px;
                    font-weight: bold;
                    color: white;
                    margin: 0 5px;
                    box-shadow: 0 4px 8px rgba(0,0,0,0.2);
                }
                .nutriscore-A { background-color: #4CAF50; }
                .nutriscore-B { background-color: #8BC34A; }
                .nutriscore-C { background-color: #FFC107; }
                .nutriscore-D { background-color: #FF9800; }
                .nutriscore-E { background-color: #F44336; }
                .progress-bar {
                    width: 100%;
                    height: 20px;
                    background-color: #f0f0f0;
                    border-radius: 10px;
                    margin: 10px 0;
                    overflow: hidden;
                }
                .progress-fill {
                    height: 100%;
                    background-color: #1E3A8A;
                    border-radius: 10px;
                    transition: width 0.3s ease;
                }
                .chart-container {
                    width: 100%;
                    height: 300px;
                    margin: 20px 0;
                }
                .highlight-box {
                    background-color: #f8f9fa;
                    border-left: 4px solid #1E3A8A;
                    padding: 15px;
                    margin: 15px 0;
                }
                .recommendation-box {
                    background-color: #fff3cd;
                    border-left: 4px solid #ffc107;
                    padding: 15px;
                    margin: 15px 0;
                }
                .resultat-table {
                    width: 100%;
                    border-collapse: collapse;
                    margin: 20px 0;
                    background-color: white;
                }
                .resultat-table th {
                    background-color: #1E3A8A;
                    color: white;
                    padding: 15px;
                    text-align: left;
                    font-weight: bold;
                    border: 1px solid #ddd;
                }
                .resultat-table td {
                    padding: 15px;
                    border: 1px solid #ddd;
                    vertical-align: top;
                }
                .resultat-table tr:nth-child(even) {
                    background-color: #f8f9fa;
                }
                .valeur-cell {
                    font-weight: bold;
                    text-align: center;
                }
                .note-cell {
                    text-align: center;
                    font-weight: bold;
                }
                .analyse-cell {
                    font-style: italic;
                }
            </style>
        </head>
        <body>
            <div class="header">
                <h1>Rapport d'Évaluation Diversité & Inclusion</h1>
                <h2>{{nom_entreprise}} - {{annee}}</h2>
            </div>

            <div class="section">
                <h3>Score Global</h3>
                <div class="highlight-box">
                    <p>Score : {{score_global}}/5</p>
                    <p>Note : <span class="note-{{note_globale}}">{{note_globale}}</span></p>
                </div>
                
                <div class="nutriscore">
                    <div class="nutriscore-letter nutriscore-{{note_globale}}">{{note_globale}}</div>
                </div>

                <div class="progress-bar">
                    <div class="progress-fill" style="width: {{(score_global/5)*100}}%"></div>
                </div>
            </div>

            <div class="section">
                <h3>Résultats Détaillés</h3>
                <table class="resultat-table">
                    <tr>
                        <th>Indicateur</th>
                        <th>Valeur Réelle</th>
                        <th>Note</th>
                        <th>Analyse</th>
                    </tr>
                    {% for resultat in resultats %}
                    <tr>
                        <td>{{resultat.indicateur}}</td>
                        <td class="valeur-cell">{{resultat.valeur}}</td>
                        <td class="note-cell note-{{resultat.note}}">{{resultat.note}}</td>
                        <td class="analyse-cell">{{resultat.analyse}}</td>
                    </tr>
                    {% endfor %}
                </table>
            </div>

            {% if points_forts %}
            <div class="section">
                <h3>Points Forts</h3>
                <div class="highlight-box">
                    <ul>
                    {% for point in points_forts %}
                        <li>{{point}}</li>
                    {% endfor %}
                    </ul>
                </div>
            </div>
            {% endif %}

            {% if axes_amelioration %}
            <div class="section">
                <h3>Axes d'Amélioration</h3>
                <div class="recommendation-box">
                    <ul>
                    {% for axe in axes_amelioration %}
                        <li>{{axe}}</li>
                    {% endfor %}
                    </ul>
                </div>
            </div>
            {% endif %}

            {% if recommandations %}
            <div class="section">
                <h3>Recommandations</h3>
                <div class="recommendation-box">
                    <ul>
                    {% for reco in recommandations %}
                        <li>{{reco}}</li>
                    {% endfor %}
                    </ul>
                </div>
            </div>
            {% endif %}

            <div class="section">
                <h3>Conclusion</h3>
                <div class="highlight-box">
                    <p>{{conclusion}}</p>
                </div>
            </div>

            <div class="footer">
                <p>Rapport généré le {{date_generation}}</p>
                <p>© 2024 Diversité & Inclusion Analytics</p>
            </div>
        </body>
        </html>
        """

        # Préparation des données pour le template avec les valeurs réelles
        template_data = {
            'nom_entreprise': company_name,
            'annee': year,
            'score_global': round(float(data['score_global']), 2),
            'note_globale': data['note_globale'],
            'resultats': [
                {
                    'indicateur': k,
                    'valeur': f"{round(float(v['valeur']), 1)}%",
                    'note': v['note'],
                    'analyse': get_analyse_indicateur(k, v['valeur'], v['note'])
                }
                for k, v in data['resultats'].items()
            ],
            'points_forts': data['points_forts'],
            'axes_amelioration': data['axes_amelioration'],
            'recommandations': [
                reco for indicateur, note in data['resultats'].items()
                if note['note'] in ['D', 'E']
                for reco in get_recommandations(indicateur, note['valeur'], note['note']).split('\n')
                if reco.strip()
            ],
            'conclusion': get_conclusion_phrase(data['note_globale']),
            'date_generation': datetime.now().strftime('%d/%m/%Y à %H:%M')
        }

        # Génération du HTML
        html = jinja2.Template(template).render(**template_data)

        # Configuration de pdfkit avec options optimisées
        options = {
            'page-size': 'A4',
            'margin-top': '20mm',
            'margin-right': '20mm',
            'margin-bottom': '20mm',
            'margin-left': '20mm',
            'encoding': 'UTF-8',
            'enable-local-file-access': None,
            'dpi': 300,
            'image-quality': 100,
            'quiet': '',
            'no-outline': None,
            'print-media-type': None
        }

        # Vérification du chemin de wkhtmltopdf avec plusieurs chemins possibles
        wkhtmltopdf_paths = [
            'C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe',
            'C:\\Program Files (x86)\\wkhtmltopdf\\bin\\wkhtmltopdf.exe',
            '/usr/local/bin/wkhtmltopdf',
            '/usr/bin/wkhtmltopdf'
        ]
        
        wkhtmltopdf_path = None
        for path in wkhtmltopdf_paths:
            if os.path.exists(path):
                wkhtmltopdf_path = path
                break

        if not wkhtmltopdf_path:
            st.error("wkhtmltopdf n'est pas installé. Veuillez l'installer depuis https://wkhtmltopdf.org/downloads.html")
            return None

        config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)
        
        # Création d'un fichier temporaire pour le PDF
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
            try:
                # Génération du PDF
                pdfkit.from_string(html, tmp.name, configuration=config, options=options)
                
                # Lecture du PDF généré
                with open(tmp.name, 'rb') as f:
                    pdf_data = f.read()
                
                return pdf_data
            finally:
                # Suppression du fichier temporaire même en cas d'erreur
                try:
                    os.unlink(tmp.name)
                except:
                    pass

    except Exception as e:
        st.error(f"Erreur lors de la génération du PDF : {str(e)}")
        return None

def get_analyse_indicateur(indicateur, valeur, note):
    """
    Génère une analyse détaillée pour chaque indicateur.
    """
    if valeur == 0:
        return "Données non disponibles pour cet indicateur."
        
    analyses = {
        "Taux de féminisation global": {
            "A": f"Avec {valeur}% de femmes, l'entreprise montre une excellente parité.",
            "B": f"Avec {valeur}% de femmes, l'entreprise est proche de la parité.",
            "C": f"Avec {valeur}% de femmes, l'entreprise a une mixité moyenne.",
            "D": f"Avec {valeur}% de femmes, l'entreprise doit améliorer sa mixité.",
            "E": f"Avec {valeur}% de femmes, l'entreprise présente un déséquilibre important."
        },
        "Taux de femmes cadres": {
            "A": f"Avec {valeur}% de femmes cadres, l'entreprise montre une excellente représentation des femmes aux postes de direction.",
            "B": f"Avec {valeur}% de femmes cadres, l'entreprise a une bonne représentation des femmes aux postes de direction.",
            "C": f"Avec {valeur}% de femmes cadres, l'entreprise a une représentation moyenne des femmes aux postes de direction.",
            "D": f"Avec {valeur}% de femmes cadres, l'entreprise doit améliorer la représentation des femmes aux postes de direction.",
            "E": f"Avec {valeur}% de femmes cadres, l'entreprise présente un déséquilibre important dans les postes de direction."
        },
        "Taux d'emploi handicap": {
            "A": f"Avec {valeur}% de personnes en situation de handicap, l'entreprise dépasse largement le seuil légal de 6%.",
            "B": f"Avec {valeur}% de personnes en situation de handicap, l'entreprise respecte bien le seuil légal de 6%.",
            "C": f"Avec {valeur}% de personnes en situation de handicap, l'entreprise est proche du seuil légal de 6%.",
            "D": f"Avec {valeur}% de personnes en situation de handicap, l'entreprise est en dessous du seuil légal de 6%.",
            "E": f"Avec {valeur}% de personnes en situation de handicap, l'entreprise est très en dessous du seuil légal de 6%."
        },
        "Écart de salaire H/F": {
            "A": f"Avec un écart de {valeur}%, l'entreprise montre une excellente équité salariale.",
            "B": f"Avec un écart de {valeur}%, l'entreprise montre une bonne équité salariale.",
            "C": f"Avec un écart de {valeur}%, l'entreprise a une équité salariale moyenne.",
            "D": f"Avec un écart de {valeur}%, l'entreprise doit améliorer son équité salariale.",
            "E": f"Avec un écart de {valeur}%, l'entreprise présente un écart salarial important."
        },
        "Équilibre des âges": {
            "A": f"Avec un score d'équilibre de {valeur}%, l'entreprise montre une excellente diversité des âges.",
            "B": f"Avec un score d'équilibre de {valeur}%, l'entreprise montre une bonne diversité des âges.",
            "C": f"Avec un score d'équilibre de {valeur}%, l'entreprise a une diversité des âges moyenne.",
            "D": f"Avec un score d'équilibre de {valeur}%, l'entreprise doit améliorer sa diversité des âges.",
            "E": f"Avec un score d'équilibre de {valeur}%, l'entreprise présente un déséquilibre important des âges."
        },
        "Taux d'absentéisme": {
            "A": f"Avec un taux d'absentéisme de {valeur}%, l'entreprise montre une excellente gestion de la santé au travail.",
            "B": f"Avec un taux d'absentéisme de {valeur}%, l'entreprise montre une bonne gestion de la santé au travail.",
            "C": f"Avec un taux d'absentéisme de {valeur}%, l'entreprise a une gestion moyenne de la santé au travail.",
            "D": f"Avec un taux d'absentéisme de {valeur}%, l'entreprise doit améliorer sa gestion de la santé au travail.",
            "E": f"Avec un taux d'absentéisme de {valeur}%, l'entreprise présente des problèmes importants de santé au travail."
        }
    }
    return analyses.get(indicateur, {}).get(note, "Analyse non disponible.")

def get_recommandations(indicateur, valeur, note):
    recommandations = {
        "Taux de féminisation global": {
            "D": """• Mettre en place un plan de recrutement ciblé pour les femmes
• Développer des partenariats avec des écoles/universités pour attirer les talents féminins
• Créer un programme de mentorat pour les femmes
• Communiquer sur les opportunités de carrière pour les femmes""",
            "E": """• Établir un plan d'action urgent pour la féminisation
• Fixer des objectifs chiffrés de recrutement de femmes
• Former les recruteurs à la lutte contre les biais
• Mettre en place un réseau de femmes dans l'entreprise"""
        },
        "Taux de femmes cadres": {
            "D": """• Identifier les femmes à fort potentiel
• Créer un programme de développement de carrière
• Mettre en place un système de parrainage
• Former les managers à la détection des talents""",
            "E": """• Réviser les processus de promotion
• Créer un programme accéléré de développement des talents féminins
• Mettre en place un comité de suivi de la parité
• Établir des objectifs de progression annuels"""
        },
        "Taux d'emploi handicap": {
            "D": """• Renforcer les partenariats avec les organismes spécialisés
• Former les managers à l'accueil des personnes en situation de handicap
• Adapter les postes de travail
• Sensibiliser les équipes""",
            "E": """• Élaborer un plan d'action urgent pour atteindre le seuil légal
• Créer un poste dédié à l'inclusion des personnes en situation de handicap
• Mettre en place un réseau d'ambassadeurs
• Réviser les processus de recrutement"""
        },
        "Écart de salaire H/F": {
            "D": """• Réaliser un audit complet des rémunérations
• Mettre en place un plan de rattrapage progressif
• Former les managers à l'équité salariale
• Établir des grilles de salaire transparentes""",
            "E": """• Corriger immédiatement les écarts injustifiés
• Mettre en place un système de contrôle régulier
• Créer un comité de suivi des rémunérations
• Publier les indicateurs d'écart de rémunération"""
        },
        "Équilibre des âges": {
            "D": """• Développer des programmes de transfert de compétences
• Mettre en place un système de tutorat intergénérationnel
• Adapter les conditions de travail pour tous les âges
• Promouvoir la diversité des âges dans la communication""",
            "E": """• Élaborer un plan de renouvellement des effectifs
• Créer des programmes de reconversion
• Mettre en place un système de préparation à la retraite
• Développer des parcours de carrière adaptés"""
        },
        "Taux d'absentéisme": {
            "D": """• Analyser les causes de l'absentéisme
• Mettre en place des actions de prévention
• Améliorer les conditions de travail
• Développer le télétravail""",
            "E": """• Réaliser un audit complet des conditions de travail
• Mettre en place un plan d'action immédiat
• Renforcer le suivi médical
• Créer un groupe de travail dédié"""
        }
    }
    return recommandations.get(indicateur, {}).get(note, "Aucune recommandation spécifique disponible.")

def get_conclusion_phrase(note):
    conclusions = {
        "A": "démontre une excellence en matière de diversité et d'inclusion.",
        "B": "présente de bonnes pratiques en matière de diversité et d'inclusion.",
        "C": "a des résultats moyens en matière de diversité et d'inclusion.",
        "D": "nécessite des améliorations significatives en matière de diversité et d'inclusion.",
        "E": "doit mettre en place un plan d'action urgent pour améliorer la diversité et l'inclusion."
    }
    return conclusions.get(note, "présente des résultats à analyser en matière de diversité et d'inclusion.")

# Section principale de génération du rapport
st.markdown("## 📄 Génération du rapport")

if 'indicateurs' in locals() and indicateurs:
    try:
        # Vérification des indicateurs requis
        indicateurs_requis = [
            "taux_feminisation", "taux_femmes_cadres", "taux_handicap",
            "ecart_salaire", "equilibre_age", "taux_absenteisme"
        ]
        
        for ind in indicateurs_requis:
            if ind not in indicateurs:
                raise ValueError(f"L'indicateur {ind} est manquant")
            if not isinstance(indicateurs[ind], (int, float)):
                raise ValueError(f"L'indicateur {ind} doit être un nombre")
        
        # Calcul des notes
        resultats = {
            "Taux de féminisation global": attribuer_note(indicateurs["taux_feminisation"], seuils["taux_feminisation"]),
            "Taux de femmes cadres": attribuer_note(indicateurs["taux_femmes_cadres"], seuils["taux_femmes_cadres"]),
            "Taux d'emploi handicap": attribuer_note(indicateurs["taux_handicap"], seuils["taux_handicap"]),
            "Écart de salaire H/F": attribuer_note(indicateurs["ecart_salaire"], seuils["ecart_salaire"], False),
            "Équilibre des âges": attribuer_note(indicateurs["equilibre_age"], seuils["equilibre_age"]),
            "Taux d'absentéisme": attribuer_note(indicateurs["taux_absenteisme"], seuils["taux_absenteisme"], False)
        }
        
        # Calcul des points forts et axes d'amélioration
        points_forts = []
        axes_amelioration = []
        
        for indicateur, note in resultats.items():
            if note in ["A", "B"]:
                points_forts.append(f"{indicateur}: Performance solide (note {note})")
            elif note in ["D", "E"]:
                axes_amelioration.append(f"{indicateur}: Nécessite des améliorations (note {note})")
        
        # Calcul de la note globale
        notes_numeriques = {cle: note_vers_chiffre(note) for cle, note in resultats.items()}
        score_global = sum(notes_numeriques.values()) / len(notes_numeriques)
        note_globale = chiffre_vers_note(score_global)
        
        # Préparation des données pour le PDF
        data = prepare_data_for_pdf(resultats, points_forts, axes_amelioration, note_globale, score_global, indicateurs)
        
        # Création du bouton pour générer le PDF
        if st.button("Générer le rapport PDF", type="primary"):
            try:
                # Génération du PDF
                pdf_data = generate_pdf(data, nom_entreprise, annee)
                
                if pdf_data:
                    # Création du bouton de téléchargement
                    st.download_button(
                        label="📥 Télécharger le rapport PDF",
                        data=pdf_data,
                        file_name=f"rapport_diversite_inclusion_{nom_entreprise}_{annee}.pdf",
                        mime="application/pdf"
                    )
                    st.success("Le rapport PDF a été généré avec succès !")
                else:
                    st.error("La génération du PDF a échoué.")
                
            except Exception as e:
                st.error(f"Erreur lors de la génération du PDF : {str(e)}")
        
    except ValueError as ve:
        st.error(f"Erreur de validation des données : {str(ve)}")
        st.error("Veuillez vérifier que toutes les données sont correctement saisies.")
    except Exception as e:
        st.error(f"Erreur lors du traitement des données : {str(e)}")
        st.error("Veuillez vérifier que toutes les données sont correctement saisies.")
else:
    st.warning("Veuillez d'abord saisir les données nécessaires pour générer le rapport.")