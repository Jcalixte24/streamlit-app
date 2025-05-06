import streamlit as st
import os
import sys
import subprocess

# Configuration de la page
st.set_page_config(
    page_title="Diversité & Inclusion - Menu Principal",
    page_icon="🏢",
    layout="wide"
)

# Style CSS personnalisé
st.markdown("""
<style>
    .main {
        background-color: #f5f5f5;
    }
    .stButton>button {
        width: 100%;
        height: 100px;
        font-size: 24px;
        margin: 10px 0;
        background-color: #1E3A8A;
        color: white;
    }
    .stButton>button:hover {
        background-color: #2563EB;
    }
    .title {
        text-align: center;
        color: #1E3A8A;
        font-size: 48px;
        margin-bottom: 30px;
    }
    .subtitle {
        text-align: center;
        color: #4B5563;
        font-size: 24px;
        margin-bottom: 50px;
    }
    .footer {
        text-align: center;
        color: #6B7280;
        margin-top: 50px;
    }
    .card {
        background-color: white;
        border-radius: 10px;
        padding: 20px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# Titre et introduction
st.markdown('<h1 class="title">🏢 Diversité & Inclusion</h1>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Plateforme d\'évaluation et d\'analyse de la diversité et inclusion en entreprise</p>', unsafe_allow_html=True)

# Création de deux colonnes pour les boutons
col1, col2 = st.columns(2)

# Fonction pour lancer une application
def launch_app(app_name):
    try:
        if app_name == "converter":
            script_path = os.path.join(os.path.dirname(__file__), "converter.py")
            if not os.path.exists(script_path):
                st.error(f"Le fichier {script_path} n'existe pas.")
                return
            subprocess.Popen([sys.executable, "-m", "streamlit", "run", script_path])
        elif app_name == "evaluation":
            script_path = os.path.join(os.path.dirname(__file__), "v6.py")
            if not os.path.exists(script_path):
                st.error(f"Le fichier {script_path} n'existe pas.")
                return
            subprocess.Popen([sys.executable, "-m", "streamlit", "run", script_path])
    except Exception as e:
        st.error(f"Erreur lors du lancement de l'application : {str(e)}")
        st.error("Veuillez vérifier que tous les fichiers nécessaires sont présents et que les dépendances sont installées.")

# Bouton pour le Convertisseur
with col1:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("""
    ### 📊 Convertisseur de Bilan Social
    Transformez vos données de bilan social au format requis pour l'évaluation.
    """)
    if st.button("Lancer le Convertisseur", key="converter"):
        launch_app("converter")
    st.markdown('</div>', unsafe_allow_html=True)

# Bouton pour l'Évaluation
with col2:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("""
    ### 📈 Évaluation D&I
    Analysez et évaluez la diversité et l'inclusion dans votre entreprise.
    """)
    if st.button("Lancer l'Évaluation", key="evaluation"):
        launch_app("evaluation")
    st.markdown('</div>', unsafe_allow_html=True)

# Section d'aide
st.markdown("---")
st.markdown("""
### 📚 Guide d'utilisation

1. **Convertisseur de Bilan Social**
   - Téléchargez le modèle Excel
   - Remplissez-le avec vos données
   - Obtenez un fichier CSV au format requis

2. **Évaluation D&I**
   - Importez votre fichier CSV
   - Consultez les analyses détaillées
   - Téléchargez le rapport PDF

### 🔧 Prérequis
- Python 3.7 ou supérieur
- Packages requis : streamlit, pandas, numpy, matplotlib, altair, plotly, reportlab, kaleido
- Pour installer les dépendances : `pip install -r requirements.txt`
""")

# Pied de page
st.markdown("""
<div class="footer">
    <p>Développé par Japhet Calixte N'DRI | Version 1.0</p>
    <p>© 2024 Tous droits réservés</p>
</div>
""", unsafe_allow_html=True) 