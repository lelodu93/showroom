import sys
import subprocess
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import plotly.express as px
from wordcloud import WordCloud
import nltk
from nltk.corpus import stopwords

# Fonction pour installer les packages si nécessaire
def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Installe les bibliothèques nécessaires
install("streamlit")
install("pandas")
install("plotly")
install("matplotlib")
install("wordcloud")

# Titre du tableau de bord
st.title("Analyse des avis clients - Showroomprive")
st.write("Ce tableau de bord permet d'analyser les avis clients collectés sur le Google Play Store.")

# Chargement des données
@st.cache_data
def load_data():
    return pd.read_excel("hetic-showroom-5mars.xlsx")

data = load_data()

# Vérification et conversion des dates
if 'date_only' in data.columns:
    data['date_only'] = pd.to_datetime(data['date_only'], errors='coerce')

# Filtres interactifs
st.sidebar.header("Filtres")

# Filtre par sentiment
sentiment_filter = st.sidebar.multiselect("Filtrer par sentiment", options=data['Sentiment'].dropna().unique(), default=data['Sentiment'].dropna().unique())

# Filtre par année
if 'date_only' in data.columns:
    min_year, max_year = int(data['date_only'].dt.year.min()), int(data['date_only'].dt.year.max())
    selected_year_range = st.sidebar.slider("Filtrer par année", min_year, max_year, (min_year, max_year))

# Filtre par subjectivité
if 'Subjectivité' in data.columns:
    min_subj, max_subj = float(data['Subjectivité'].min()), float(data['Subjectivité'].max())
    selected_subj_range = st.sidebar.slider("Filtrer par subjectivité", min_subj, max_subj, (min_subj, max_subj))

# Filtre par score
if 'score' in data.columns:
    min_score, max_score = int(data['score'].min()), int(data['score'].max())
    selected_score_range = st.sidebar.slider("Filtrer par score", min_score, max_score, (min_score, max_score))

# Application des filtres sur les données
filtered_data = data[data['Sentiment'].isin(sentiment_filter)]

if 'date_only' in data.columns:
    filtered_data = filtered_data[
        (filtered_data['date_only'].dt.year >= selected_year_range[0]) &
        (filtered_data['date_only'].dt.year <= selected_year_range[1])
    ]

if 'Subjectivité' in data.columns:
    filtered_data = filtered_data[
        (filtered_data['Subjectivité'] >= selected_subj_range[0]) &
        (filtered_data['Subjectivité'] <= selected_subj_range[1])
    ]

if 'score' in data.columns:
    filtered_data = filtered_data[
        (filtered_data['score'] >= selected_score_range[0]) &
        (filtered_data['score'] <= selected_score_range[1])
    ]

# Affichage des KPI
st.header("Indicateurs clés de performance (KPI)")
col1, col2 = st.columns(2)
with col1:
    st.metric("Nombre total d'avis", len(filtered_data))
with col2:
    st.metric("Pourcentage d'avis positifs", f"{len(filtered_data[filtered_data['Sentiment'] == 'positif']) / len(filtered_data) * 100:.2f}%" if len(filtered_data) > 0 else "0%")

# Graphique des sentiments
st.header("Répartition des sentiments")
sentiment_counts = filtered_data['Sentiment'].value_counts()
fig, ax = plt.subplots()
ax.pie(sentiment_counts, labels=sentiment_counts.index, autopct='%1.1f%%', startangle=90)
ax.axis('equal')
st.pyplot(fig)

# Vérifier les types de données et nettoyer les scores
if 'score' in filtered_data.columns:
    if not pd.api.types.is_numeric_dtype(filtered_data['score']):
        filtered_data['score'] = pd.to_numeric(filtered_data['score'], errors='coerce')

    # Supprimer les valeurs nulles ou aberrantes
    filtered_data = filtered_data.dropna(subset=['score'])
    filtered_data = filtered_data[filtered_data['score'].between(1, 5)]  # On garde uniquement les scores entre 1 et 5

# Tracer le graphique corrigé avec clé unique
st.header("Distribution des notes (scores) - ")
if not filtered_data.empty and 'score' in filtered_data.columns:
    score_counts = filtered_data['score'].value_counts().sort_index()
    fig_score = px.bar(score_counts, x=score_counts.index, y=score_counts.values, labels={'x': 'Score', 'y': "Nombre d'avis"})
    st.plotly_chart(fig_score, key="score_chart")

# Nuage de mots amélioré
st.header("Nuage de mots des avis")
try:
    stopwords_fr = set(stopwords.words('french'))
    mots_a_exclure = {"le", "la", "les", "de", "des", "un", "une", "en", "au", "aux", "avec", "pour", "dans", "sur", "est", "du", "c'est", "n'est", "pas", "vous", "je", "il", "elle", "à", "et", "ou", "son", "sa", "vos", "www", "http", "https", "com", "fr", "j'ai", "ça", "qu'il", "trs", "mes", "mais", "dlai", "plu", "pa", "qui", "dan"}
    tous_stopwords = stopwords_fr.union(mots_a_exclure)

    filtered_reviews = filtered_data['avis'].dropna().astype(str)
    text = " ".join(review for review in filtered_reviews)

    wordcloud = WordCloud(background_color="white", width=800, height=400, stopwords=tous_stopwords, collocations=False, max_words=200).generate(text)

    fig_wordcloud, ax = plt.subplots()
    ax.imshow(wordcloud, interpolation='bilinear')
    ax.axis('off')
    st.pyplot(fig_wordcloud)

except Exception as e:
    st.error(f"Erreur lors de la génération du nuage de mots : {str(e)}")

# Graphique temporel
st.header("Évolution des avis dans le temps")
if 'date_only' in filtered_data.columns:
    time_series = filtered_data.groupby('date_only').size()
    fig_time = px.line(time_series, x=time_series.index, y=time_series.values, labels={'x': 'Date', 'y': "Nombre d'avis"})
    st.plotly_chart(fig_time, key="time_series_chart")
else:
    st.write("La colonne 'date_only' n'est pas disponible pour l'analyse temporelle.")

# Tableau interactif avec coloration
st.header("Tableau des avis")
def color_sentiment(val):
    color = 'green' if val == 'positif' else 'red' if val == 'negatif' else 'gray'
    return f'background-color: {color}'

if not filtered_data.empty:
    styled_table = filtered_data[['avis', 'Sentiment', 'date_only', 'score', 'Subjectivité']].style.applymap(color_sentiment, subset=['Sentiment'])
    st.dataframe(styled_table)
else:
    st.write("Aucun avis disponible après application des filtres.")
