# -*- coding: utf-8 -*-
"""
INCLUSION PROFESSIONNELLE 15
Écosystème numérique intégré de vérification de l'application de la loi n° 2010-15
Application Streamlit – version initiale opérationnelle et extensible.

Installation :
    pip install streamlit pandas numpy plotly scikit-learn openpyxl matplotlib
Lancement :
    streamlit run app_inclusion_professionnelle_15.py

Dépendances facultatives :
    pip install openai pypdf python-docx
"""

from __future__ import annotations

import os
import io
import re
import json
import sqlite3
import hashlib
from pathlib import Path
from datetime import datetime, date
from typing import Any

import numpy as np
import pandas as pd
import streamlit as st

try:
    import plotly.express as px
    import plotly.graph_objects as go
    PLOTLY_OK = True
except Exception:
    PLOTLY_OK = False

try:
    from sklearn.compose import ColumnTransformer
    from sklearn.pipeline import Pipeline
    from sklearn.impute import SimpleImputer
    from sklearn.preprocessing import OneHotEncoder, StandardScaler
    from sklearn.ensemble import RandomForestClassifier, RandomForestRegressor
    from sklearn.model_selection import train_test_split
    from sklearn.metrics import (
        balanced_accuracy_score, accuracy_score, classification_report,
        confusion_matrix, mean_absolute_error, mean_squared_error, r2_score
    )
    SKLEARN_OK = True
except Exception:
    SKLEARN_OK = False

# =============================================================================
# 1. CONFIGURATION GÉNÉRALE
# =============================================================================

APP_NAME = "INCLUSION PROFESSIONNELLE 15"
APP_SUBTITLE = "Écosystème numérique intégré de vérification de l'application de la loi n° 2010-15"
LEGAL_REFERENCE = "Loi n° 2010-15 du 6 juillet 2010"
BASE_DIR = Path(os.getenv("IP15_DATA_DIR", "data_inclusion_professionnelle_15"))
UPLOAD_DIR = BASE_DIR / "bibliotheque_juridique"
EXPORT_DIR = BASE_DIR / "exports"
DB_PATH = BASE_DIR / "inclusion_professionnelle_15.sqlite3"

for directory in (BASE_DIR, UPLOAD_DIR, EXPORT_DIR):
    directory.mkdir(parents=True, exist_ok=True)

C = {
    "navy": "#10243A", "navy2": "#193B5A", "blue": "#28658A",
    "teal": "#239B91", "gold": "#C9A44C", "bg": "#F2F5F8",
    "white": "#FFFFFF", "text": "#263746", "muted": "#718096",
    "border": "#DCE4EC", "lightblue": "#EAF1F7",
    "lightteal": "#E8F4F1", "lightgold": "#F7F1E5",
    "red": "#B94A48", "green": "#26856E"
}

st.set_page_config(
    page_title=APP_NAME,
    page_icon="🇸🇳",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap');
html, body, [class*="css"] {{ font-family: 'DM Sans', sans-serif; color: #172B3A; }}
.stApp {{ background: {C['bg']}; }}
section[data-testid="stSidebar"] {{ background: #FFFFFF; border-right: 1px solid {C['border']}; }}
section[data-testid="stSidebar"] * {{ color: #172B3A !important; }}
section[data-testid="stSidebar"] .stRadio label {{ font-weight: 500; }}
h1, h2, h3, h4, p, label, span, div {{ }}
h1, h2, h3, h4 {{ color: {C['navy']} !important; }}
p, label, .stMarkdown, [data-testid="stMetricLabel"] {{ color: #263746; }}
[data-testid="stMetricValue"] {{ color: {C['navy']} !important; font-weight: 700; }}
.ip15-header {{
 background: linear-gradient(105deg, {C['navy']}, {C['navy2']});
 color: white; padding: 1.25rem 1.5rem; border-radius: 14px;
 border-left: 7px solid {C['teal']}; margin-bottom: 1rem;
}}
.ip15-header h1 {{ color: white !important; margin: 0; font-size: 1.8rem; }}
.ip15-header p {{ color: #DCE8F0 !important; margin: .35rem 0 0 0; }}
.ip15-kicker {{ color: #B9D0DF !important; text-transform: uppercase;
 letter-spacing: .08em; font-size: .72rem; font-weight: 700; }}
.ip15-card {{ background: white; border: 1px solid {C['border']};
 border-radius: 12px; padding: 1rem; margin-bottom: .65rem; }}
.ip15-note {{ background: {C['lightgold']}; border: 1px solid #E8D9B6;
 border-radius: 9px; padding: .8rem 1rem; color: {C['navy']}; }}
div.stButton > button[kind="primary"] {{ background: {C['navy']}; border-color: {C['navy']}; }}
.stProgress > div > div > div > div {{ background-color: {C['teal']}; }}
[data-testid="stDataFrame"] {{ background: white; }}
footer {{ visibility: hidden; }}
</style>
""", unsafe_allow_html=True)

# =============================================================================
# 2. RÉFÉRENTIEL DES 50 INDICATEURS – TRANSCRIPTION DU SCRIPT FOURNI
# =============================================================================

CHAPTERS = [
    {
        "id": "I", "name": "DISPOSITIONS GÉNÉRALES", "articles": "Articles 1 à 6",
        "color": C["blue"], "items": [
            ("IP15-01", 1, "Prise en compte de la définition légale du handicap dans les dispositifs institutionnels"),
            ("IP15-02", 2, "Effectivité des mesures d'égalité des chances et de non-discrimination"),
            ("IP15-03", 3, "Fonctionnement du dispositif de délivrance de la carte d'égalité des chances"),
            ("IP15-04", 4, "Participation des personnes handicapées à la vie sociale, économique et culturelle"),
            ("IP15-05", 5, "Mise en œuvre des obligations nationales de prévention, de réadaptation et d'intégration"),
            ("IP15-06", 6, "Implication des collectivités territoriales et consultation des organisations représentatives"),
        ]
    },
    {
        "id": "II", "name": "SANTÉ, ACTION SOCIALE ET PRÉVENTION", "articles": "Articles 7 à 14",
        "color": C["teal"], "items": [
            ("IP15-07", 7, "Accessibilité effective des soins médicaux nécessaires"),
            ("IP15-08", 8, "Application de la gratuité et des réductions de prestations de santé"),
            ("IP15-09", 9, "Mise en œuvre des dispositifs d'assistance et de prise en charge sociale"),
            ("IP15-10", 10, "Accès aux appareils orthopédiques et aux aides techniques"),
            ("IP15-11", 11, "Disponibilité du personnel qualifié et soutien aux équipements spécialisés"),
            ("IP15-12", 12, "Existence et mise en œuvre des programmes de prévention du handicap"),
            ("IP15-13", 13, "Réalisation des campagnes de sensibilisation et d'information"),
            ("IP15-14", 14, "Intégration de la prévention des risques dans les milieux professionnels et les programmes de développement"),
        ]
    },
    {
        "id": "III", "name": "ÉDUCATION, FORMATION ET EMPLOI", "articles": "Articles 15 à 30",
        "color": C["gold"], "items": [
            ("IP15-15", 15, "Accès effectif à l'éducation, à l'enseignement et à l'emploi"),
            ("IP15-16", 16, "Mise en place et fonctionnement des commissions départementales d'éducation spéciale"),
            ("IP15-17", 17, "Effectivité des décisions d'orientation et respect du choix de l'établissement"),
            ("IP15-18", 18, "Disponibilité des aménagements pédagogiques, humains et matériels"),
            ("IP15-19", 19, "Application des dispositions particulières relatives aux âges, examens et concours"),
            ("IP15-20", 20, "Application des réductions de frais de scolarité et attribution des bourses"),
            ("IP15-21", 21, "Disponibilité des supports adaptés : braille, bibliothèques sonores et langue des signes"),
            ("IP15-22", 22, "Accès à la formation technique et professionnelle ordinaire ou adaptée"),
            ("IP15-23", 23, "Application du quota de places en formation professionnelle et aménagement des postes"),
            ("IP15-24", 24, "Adaptation des centres de formation aux besoins spécifiques des apprenants"),
            ("IP15-25", 25, "Création et fonctionnement de filières et centres spécialisés de formation professionnelle"),
            ("IP15-26", 26, "Respect des règles d'admission, de suivi pédagogique et de non-discrimination à l'emploi"),
            ("IP15-27", 27, "Maintien dans l'emploi ou reclassement des agents devenus handicapés"),
            ("IP15-28", 28, "Transmission des déclarations relatives aux attributions, suspensions et suppressions d'emploi"),
            ("IP15-29", 29, "Application de la priorité de mutation et suivi de la proportion des emplois accessibles réservés"),
            ("IP15-30", 30, "Effectivité des aides à la création d'entreprises et aux activités économiques"),
        ]
    },
    {
        "id": "IV", "name": "ACCESSIBILITÉ, HABITAT ET TRANSPORT", "articles": "Articles 31 à 38",
        "color": C["blue"], "items": [
            ("IP15-31", 31, "Conformité de l'accessibilité des bâtiments, routes, trottoirs et espaces extérieurs"),
            ("IP15-32", 32, "Intégration des normes d'accessibilité dans les autorisations de construire et les travaux"),
            ("IP15-33", 33, "Accessibilité des transports collectifs et signalétique adaptée"),
            ("IP15-34", 34, "Accès aux transports adaptés, réductions tarifaires et avantages des accompagnateurs"),
            ("IP15-35", 35, "Respect des conditions d'exonération et d'utilisation des véhicules spécialisés"),
            ("IP15-36", 36, "Effectivité des autorisations de stationnement à proximité du lieu de travail"),
            ("IP15-37", 37, "Accès des personnes handicapées aux programmes de logements sociaux"),
            ("IP15-38", 38, "Accès à la terre et aux outils de production"),
        ]
    },
    {
        "id": "V", "name": "SPORTS, LOISIRS, ARTS ET CULTURE", "articles": "Articles 39 à 42",
        "color": C["teal"], "items": [
            ("IP15-39", 39, "Accessibilité des infrastructures sportives, culturelles et de loisirs"),
            ("IP15-40", 40, "Intégration des sports adaptés dans les centres et programmes scolaires et universitaires"),
            ("IP15-41", 41, "Accessibilité des institutions culturelles et disponibilité des équipements spécifiques"),
            ("IP15-42", 42, "Accessibilité des espaces de jeux publics aux enfants handicapés"),
        ]
    },
    {
        "id": "VI", "name": "ORGANISATION DES PERSONNES HANDICAPÉES", "articles": "Articles 43 à 44",
        "color": C["gold"], "items": [
            ("IP15-43", 43, "Consultation effective des organisations dans les politiques publiques"),
            ("IP15-44", 44, "Soutien public, contrôle et suivi des organisations et structures spécialisées"),
        ]
    },
    {
        "id": "VII", "name": "DISPOSITIONS DIVERSES ET FINALES", "articles": "Articles 45 à 50",
        "color": C["navy2"], "items": [
            ("IP15-45", 45, "Organisation de la semaine nationale et des journées de sensibilisation"),
            ("IP15-46", 46, "Prise en compte des besoins des personnes handicapées en milieu pénitentiaire"),
            ("IP15-47", 47, "Mise en place, financement et fonctionnement du fonds d'appui"),
            ("IP15-48", 48, "Mise en place et fonctionnement de la Haute Autorité"),
            ("IP15-49", 49, "Continuité de l'accès aux droits pendant la période transitoire relative à la carte"),
            ("IP15-50", 50, "Prise en compte des dispositions d'abrogation et de remplacement dans le cadre juridique"),
        ]
    },
]

INDICATORS = []
for ch in CHAPTERS:
    for code, article, label in ch["items"]:
        INDICATORS.append({
            "Code": code, "Article": article, "Chapitre": ch["id"],
            "Domaine": ch["name"], "Indicateur": label
        })
assert len(INDICATORS) == 50
INDICATOR_DF = pd.DataFrame(INDICATORS)

TOMES = [
    ("TOME I", "RÉFÉRENTIEL JURIDIQUE", "Textes, obligations et cadre réglementaire"),
    ("TOME II", "50 INDICATEURS", "Référentiel de vérification couvrant les 7 chapitres"),
    ("TOME III", "PROCÉDURES", "Manuel de vérification administrative et opérationnelle"),
    ("TOME IV", "RISQUES & PERFORMANCE", "Analyse de conformité et des résultats"),
    ("TOME V", "DÉPLOIEMENT", "Recommandations et suivi de leur exécution"),
    ("TOME VI", "VÉRIFICATION PRÉVENTIVE", "Normes professionnelles et inclusion dans les grands projets de l'État"),
]

# =============================================================================
# 3. BASE DE DONNÉES SQLITE – PERSISTANCE LOCALE
# =============================================================================

def db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with db() as con:
        con.executescript("""
        CREATE TABLE IF NOT EXISTS missions (
            id INTEGER PRIMARY KEY AUTOINCREMENT, reference TEXT, organisme TEXT,
            type_mission TEXT, objet TEXT, perimetre TEXT, date_debut TEXT,
            responsable TEXT, statut TEXT, observations TEXT, created_at TEXT
        );
        CREATE TABLE IF NOT EXISTS evaluations (
            id INTEGER PRIMARY KEY AUTOINCREMENT, mission_id INTEGER,
            organisme TEXT, code_indicateur TEXT, statut TEXT, score INTEGER,
            constat TEXT, preuve TEXT, source TEXT, responsable TEXT,
            date_evaluation TEXT, commentaire TEXT
        );
        CREATE TABLE IF NOT EXISTS recommendations (
            id INTEGER PRIMARY KEY AUTOINCREMENT, mission_id INTEGER,
            reference TEXT, organisme TEXT, recommandation TEXT, priorite TEXT,
            responsable TEXT, echeance TEXT, statut TEXT, preuve_execution TEXT,
            date_creation TEXT, historique TEXT
        );
        CREATE TABLE IF NOT EXISTS legal_docs (
            id INTEGER PRIMARY KEY AUTOINCREMENT, nom_original TEXT,
            nom_stocke TEXT, famille TEXT, reference_officielle TEXT,
            autorite TEXT, date_texte TEXT, statut TEXT, mots_cles TEXT,
            chemin TEXT, empreinte TEXT, date_ajout TEXT
        );
        CREATE TABLE IF NOT EXISTS questionnaires (
            id INTEGER PRIMARY KEY AUTOINCREMENT, mission_id INTEGER,
            organisme TEXT, domaine TEXT, question TEXT, reponse TEXT,
            commentaire TEXT, date_saisie TEXT
        );
        CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT, action TEXT,
            objet TEXT, detail TEXT, horodatage TEXT
        );
        CREATE TABLE IF NOT EXISTS matrix_rows (
            id INTEGER PRIMARY KEY AUTOINCREMENT, categorie TEXT,
            code TEXT, contenu TEXT, statut TEXT, justification TEXT,
            date_modification TEXT
        );
        """)
init_db()

def log_action(action: str, objet: str, detail: str = ""):
    with db() as con:
        con.execute(
            "INSERT INTO audit_log(action,objet,detail,horodatage) VALUES(?,?,?,?)",
            (action, objet, detail, datetime.now().isoformat(timespec="seconds"))
        )

def qdf(query: str, params: tuple = ()) -> pd.DataFrame:
    with db() as con:
        return pd.read_sql_query(query, con, params=params)

def execute(query: str, params: tuple = ()):
    with db() as con:
        con.execute(query, params)
        con.commit()

def csv_download(df: pd.DataFrame, filename: str, label: str):
    data = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(label, data, file_name=filename, mime="text/csv", use_container_width=False)

def dataframe_download_excel(df: pd.DataFrame, filename: str, label: str):
    buffer = io.BytesIO()
    try:
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Données")
        st.download_button(label, buffer.getvalue(), filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception:
        csv_download(df, filename.replace(".xlsx", ".csv"), label + " (CSV)")

# =============================================================================
# 4. OUTILS COMMUNS
# =============================================================================

def page_header(title: str, subtitle: str):
    st.markdown(
        f'<div class="ip15-header"><div class="ip15-kicker">Sénégal · Modernisation de la vérification publique</div>'
        f'<h1>{title}</h1><p>{subtitle}</p></div>',
        unsafe_allow_html=True
    )

def metric_row(items):
    cols = st.columns(len(items))
    for col, (label, value, help_text) in zip(cols, items):
        col.metric(label, value, help=help_text)

def get_missions():
    return qdf("SELECT * FROM missions ORDER BY id DESC")

def mission_options():
    missions = get_missions()
    if missions.empty:
        return {"Aucune mission enregistrée": None}
    return {
        f"#{r.id} · {r.reference or 'Sans référence'} · {r.organisme or 'Organisme non précisé'}": int(r.id)
        for r in missions.itertuples()
    }

def indicator_options():
    return {
        f"{r.Code} | Art. {r.Article} | {r.Indicateur}": r.Code
        for r in INDICATOR_DF.itertuples()
    }

def get_indicator(code):
    row = INDICATOR_DF[INDICATOR_DF["Code"] == code]
    return row.iloc[0].to_dict() if not row.empty else {}

def risk_label(score):
    if score >= 75: return "Élevé"
    if score >= 50: return "Significatif"
    if score >= 25: return "Modéré"
    return "Faible"

def safe_filename(name: str):
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", Path(name).name)
    return stem[:120] or "document"

# =============================================================================
# 5. BARRE LATÉRALE ET NAVIGATION
# =============================================================================

PAGES = [
    "Accueil & tableau de bord",
    "Déclenchement de mission",
    "Organismes & services à vérifier",
    "Analyse des risques",
    "Contrôles automatisés",
    "Assistant méthodologique IA",
    "Préparation de mission",
    "Travaux de terrain",
    "Bibliothèque juridique",
    "Référentiel des 50 indicateurs",
    "Matrices & questionnaires",
    "Suivi des recommandations",
    "Rapports & exports",
    "Paramètres & journal",
    "TOME VI – Vérification préventive des grands projets"
]

with st.sidebar:
    st.markdown(f"""
    <div style="padding:.5rem 0 1rem 0">
      <div style="font-size:1.7rem;font-weight:800;color:{C['navy']}">🇸🇳 IP15</div>
      <div style="font-size:.72rem;font-weight:700;color:{C['teal']};letter-spacing:.06em">INCLUSION PROFESSIONNELLE</div>
      <div style="font-size:.68rem;color:#657586">Vérification de l'effectivité des droits</div>
    </div>
    """, unsafe_allow_html=True)
    page = st.radio("NAVIGATION", PAGES, label_visibility="collapsed")
    st.divider()
    st.caption("Référence juridique")
    st.markdown(f"**{LEGAL_REFERENCE}**")
    st.caption("50 indicateurs · 7 chapitres · 5 tomes")
    st.caption("Version fonctionnelle initiale")

# =============================================================================
# 6. PAGE 1 – ACCUEIL & TABLEAU DE BORD
# =============================================================================

if page == PAGES[0]:
    page_header(APP_NAME, APP_SUBTITLE)
    missions = get_missions()
    evals = qdf("SELECT * FROM evaluations")
    recs = qdf("SELECT * FROM recommendations")
    docs = qdf("SELECT * FROM legal_docs")

    assessed = evals["code_indicateur"].nunique() if not evals.empty else 0
    completed = int((recs["statut"] == "Clôturée").sum()) if not recs.empty else 0
    open_recs = int((recs["statut"] != "Clôturée").sum()) if not recs.empty else 0

    metric_row([
        ("Articles couverts", "50", "Référentiel de travail fourni"),
        ("Indicateurs", "50", "IP15-01 à IP15-50"),
        ("Chapitres", "07", "Structure du script source"),
        ("Missions enregistrées", len(missions), "Données saisies dans cette installation"),
        ("Indicateurs évalués", f"{assessed}/50", "Indicateurs ayant au moins une évaluation"),
    ])
    st.markdown("### Architecture méthodologique")
    cols = st.columns(5)
    for col, (tome, title, desc) in zip(cols, TOMES):
        with col:
            st.markdown(f"""<div class="ip15-card" style="border-top:4px solid {C['teal']}">
            <div style="font-size:.75rem;color:{C['teal']};font-weight:700">{tome}</div>
            <div style="font-weight:700;color:{C['navy']};margin:.3rem 0">{title}</div>
            <div style="font-size:.78rem;color:#657586">{desc}</div></div>""", unsafe_allow_html=True)

    st.markdown("### Couverture de la loi par chapitre")
    chapter_data = pd.DataFrame([
        {"Chapitre": f"{c['id']} · {c['name']}", "Indicateurs": len(c["items"])}
        for c in CHAPTERS
    ])
    if PLOTLY_OK:
        fig = px.bar(chapter_data, x="Chapitre", y="Indicateurs", text="Indicateurs",
                     color="Indicateurs", color_continuous_scale=["#B9D6E5", C["teal"], C["navy"]])
        fig.update_layout(showlegend=False, height=340, plot_bgcolor="white", paper_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.bar_chart(chapter_data.set_index("Chapitre"))
    a, b, c = st.columns(3)
    a.metric("Recommandations ouvertes", open_recs)
    b.metric("Recommandations clôturées", completed)
    c.metric("Documents juridiques indexés", len(docs))
    st.markdown('<div class="ip15-note"><b>Garde-fou professionnel :</b> les indicateurs sont des axes de vérification. Une évaluation doit reposer sur des critères applicables, des pièces probantes, des entretiens et une appréciation contradictoire. Les scores et alertes ne constituent pas, à eux seuls, des preuves de non-conformité.</div>', unsafe_allow_html=True)

# =============================================================================
# 7. PAGE 2 – DÉCLENCHEMENT DE MISSION
# =============================================================================

elif page == PAGES[1]:
    page_header("Déclenchement de mission", "Formaliser le mandat, l'objet, le périmètre et les responsabilités.")
    with st.form("mission_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        reference = c1.text_input("Référence de la mission")
        organisme = c2.text_input("Organisme / institution / service concerné *")
        type_mission = c1.selectbox("Type de mission", [
            "Vérification de conformité", "Audit de performance", "Évaluation de politique publique",
            "Mission de conseil", "Suivi de recommandations", "Diagnostic institutionnel"
        ])
        date_debut = c2.date_input("Date prévisionnelle de début", value=date.today())
        objet = st.text_area("Objet de la mission *")
        perimetre = st.text_area("Périmètre institutionnel, géographique et temporel")
        responsable = st.text_input("Chef de mission / responsable")
        observations = st.text_area("Observations et réserves initiales")
        statut = st.selectbox("Statut", ["À cadrer", "Préparation", "En cours", "Suspendue", "Clôturée"])
        submitted = st.form_submit_button("Enregistrer la mission", type="primary")
    if submitted:
        if not organisme.strip() or not objet.strip():
            st.error("L'organisme et l'objet de la mission sont obligatoires.")
        else:
            execute("""INSERT INTO missions(reference,organisme,type_mission,objet,perimetre,date_debut,
                       responsable,statut,observations,created_at) VALUES(?,?,?,?,?,?,?,?,?,?)""",
                    (reference, organisme, type_mission, objet, perimetre, date_debut.isoformat(),
                     responsable, statut, observations, datetime.now().isoformat(timespec="seconds")))
            log_action("Création", "Mission", f"{reference} – {organisme}")
            st.success("Mission enregistrée dans la base locale.")
    st.markdown("### Checklist de cadrage")
    checks = [
        "Mandat et autorisation de mission disponibles",
        "Objet et périmètre validés",
        "Organismes et services concernés identifiés",
        "Responsabilités et interlocuteurs désignés",
        "Période et population concernées précisées",
        "Règles de confidentialité et de protection des données définies",
        "Modalités de communication et de contradiction prévues"
    ]
    for item in checks:
        st.checkbox(item, key="mission_check_" + hashlib.md5(item.encode()).hexdigest()[:8])
    missions = get_missions()
    st.markdown("### Missions enregistrées")
    if missions.empty: st.info("Aucune mission enregistrée.")
    else:
        st.dataframe(missions, use_container_width=True, hide_index=True)

# =============================================================================
# 8. PAGE 3 – ORGANISMES & SERVICES À VÉRIFIER
# =============================================================================

elif page == PAGES[2]:
    page_header("Organismes & services à vérifier", "Constituer un portefeuille d'entités et organiser la priorisation des travaux.")
    st.info("Importez vos données réelles. Aucun organisme, score ou classement n'est préchargé ou inventé.")
    up = st.file_uploader("Importer un fichier CSV ou Excel", type=["csv", "xlsx"], key="org_upload")
    if up:
        try:
            df_org = pd.read_csv(up) if up.name.lower().endswith(".csv") else pd.read_excel(up)
            st.write("Aperçu du fichier")
            st.dataframe(df_org.head(30), use_container_width=True)
            required = {"organisme", "score_risque"}
            if required.issubset(set(df_org.columns)):
                df_org["score_risque"] = pd.to_numeric(df_org["score_risque"], errors="coerce")
                valid = df_org.dropna(subset=["organisme", "score_risque"]).copy()
                valid["niveau_risque"] = valid["score_risque"].apply(lambda x: risk_label(float(x)) if 0 <= x <= 100 else "À vérifier")
                valid = valid.sort_values("score_risque", ascending=False)
                st.markdown("### Classement calculé à partir du fichier importé")
                st.dataframe(valid, use_container_width=True, hide_index=True)
                csv_download(valid, "portefeuille_organismes_ip15.csv", "Exporter le classement")
            else:
                st.warning("Colonnes obligatoires : organisme et score_risque. Colonne facultative : risque_principal, date_analyse.")
        except Exception as e:
            st.error(f"Lecture impossible : {e}")
    st.markdown("### Registre manuel des entités")
    with st.form("org_manual"):
        x1, x2 = st.columns(2)
        ent = x1.text_input("Organisme / service")
        domaine = x2.selectbox("Domaine principal", [c["name"] for c in CHAPTERS])
        description = st.text_area("Mandat, missions ou observations")
        if st.form_submit_button("Ajouter au portefeuille"):
            if ent.strip():
                execute("INSERT INTO matrix_rows(categorie,code,contenu,statut,justification,date_modification) VALUES(?,?,?,?,?,?)",
                        ("Organisme", ent, description, "À examiner", domaine, datetime.now().isoformat()))
                log_action("Ajout", "Organisme", ent)
                st.success("Entité ajoutée au registre.")
    orgs = qdf("SELECT * FROM matrix_rows WHERE categorie='Organisme' ORDER BY id DESC")
    if not orgs.empty: st.dataframe(orgs, use_container_width=True, hide_index=True)

# =============================================================================
# 9. PAGE 4 – ANALYSE DES RISQUES
# =============================================================================

elif page == PAGES[3]:
    page_header("Analyse des risques", "Explorer les facteurs de risque et, si les données le permettent, entraîner un modèle prédictif.")
    st.warning("Un modèle prédictif est un outil d'aide à la priorisation. Il ne démontre ni une violation de la loi ni une discrimination.")
    up = st.file_uploader("Importer un historique CSV / Excel", type=["csv", "xlsx"], key="risk_upload")
    if up:
        try:
            df = pd.read_csv(up) if up.name.lower().endswith(".csv") else pd.read_excel(up)
            st.dataframe(df.head(20), use_container_width=True)
            if not SKLEARN_OK:
                st.error("Scikit-learn n'est pas installé. Installez-le avec : pip install scikit-learn")
            else:
                target = st.selectbox("Variable cible (résultat déjà observé)", list(df.columns))
                features = st.multiselect("Variables explicatives", [c for c in df.columns if c != target],
                                          default=[c for c in df.columns if c != target][:min(5, len(df.columns)-1)])
                if st.button("Entraîner et évaluer le modèle", type="primary"):
                    data = df[features + [target]].dropna(subset=[target]).copy()
                    if len(data) < 20 or not features:
                        st.error("Il faut au moins 20 observations exploitables et une variable explicative.")
                    elif data[target].nunique() < 2:
                        st.error("La variable cible doit comporter au moins deux classes.")
                    else:
                        X, y = data[features], data[target]
                        cats = X.select_dtypes(include=["object", "category", "bool"]).columns.tolist()
                        nums = [x for x in features if x not in cats]
                        transformers = []
                        if nums:
                            transformers.append(("num", Pipeline([("imputer", SimpleImputer(strategy="median")),
                                                                  ("scale", StandardScaler())]), nums))
                        if cats:
                            transformers.append(("cat", Pipeline([("imputer", SimpleImputer(strategy="most_frequent")),
                                                                  ("onehot", OneHotEncoder(handle_unknown="ignore"))]), cats))
                        prep = ColumnTransformer(transformers)
                        stratify = y if y.value_counts().min() >= 2 else None
                        Xtr, Xte, ytr, yte = train_test_split(X, y, test_size=.25, random_state=42, stratify=stratify)
                        model = Pipeline([("prep", prep), ("model", RandomForestClassifier(
                            n_estimators=300, class_weight="balanced", min_samples_leaf=2, random_state=42
                        ))])
                        model.fit(Xtr, ytr)
                        pred = model.predict(Xte)
                        st.metric("Exactitude équilibrée", f"{balanced_accuracy_score(yte, pred):.1%}")
                        st.text("Rapport de classification")
                        st.code(classification_report(yte, pred, zero_division=0))
                        st.write("Matrice de confusion")
                        cm = confusion_matrix(yte, pred, labels=model.named_steps["model"].classes_)
                        st.dataframe(pd.DataFrame(cm, index=model.named_steps["model"].classes_,
                                                  columns=model.named_steps["model"].classes_))
                        st.session_state["ip15_model"] = model
                        st.session_state["ip15_features"] = features
                        st.success("Modèle entraîné en mémoire pour cette session. Réentraînement et validation indépendante nécessaires avant tout usage opérationnel.")
                if "ip15_model" in st.session_state:
                    st.markdown("### Scorer de nouvelles observations")
                    newfile = st.file_uploader("Fichier des nouvelles observations", type=["csv", "xlsx"], key="newrisk")
                    if newfile:
                        newdf = pd.read_csv(newfile) if newfile.name.lower().endswith(".csv") else pd.read_excel(newfile)
                        feats = st.session_state["ip15_features"]
                        if set(feats).issubset(newdf.columns):
                            model = st.session_state["ip15_model"]
                            out = newdf.copy()
                            out["classe_predite"] = model.predict(out[feats])
                            if hasattr(model, "predict_proba"):
                                out["confiance_max_modele"] = model.predict_proba(out[feats]).max(axis=1)
                            st.dataframe(out, use_container_width=True)
                            csv_download(out, "scoring_risques_ip15.csv", "Exporter le scoring")
                        else: st.error("Les colonnes du fichier ne correspondent pas aux variables d'entraînement.")
        except Exception as e:
            st.error(f"Erreur d'analyse : {e}")
    st.markdown("### Grille qualitative de risques")
    risk_rows = []
    for ch in CHAPTERS:
        risk_rows.append({"Domaine": ch["name"], "Question de cadrage": "Quels obstacles peuvent empêcher l'effectivité des droits ?", "Niveau": "À apprécier", "Justification": ""})
    st.dataframe(pd.DataFrame(risk_rows), use_container_width=True, hide_index=True)

# =============================================================================
# 10. PAGE 5 – CONTRÔLES AUTOMATISÉS
# =============================================================================

elif page == PAGES[4]:
    page_header("Contrôles automatisés", "Contrôles exploratoires de complétude, doublons, cohérence et valeurs atypiques.")
    st.info("Les alertes doivent être vérifiées sur pièces et replacées dans leur contexte. Elles ne constituent pas des conclusions.")
    up = st.file_uploader("Importer un fichier de données", type=["csv", "xlsx"], key="controls_upload")
    if up:
        try:
            df = pd.read_csv(up) if up.name.lower().endswith(".csv") else pd.read_excel(up)
            st.caption(f"{len(df):,} lignes · {len(df.columns)} colonnes")
            tab1, tab2, tab3, tab4 = st.tabs(["Complétude", "Doublons", "Valeurs numériques", "Synthèse"])
            with tab1:
                selected = st.multiselect("Champs à contrôler", list(df.columns), default=list(df.columns))
                miss = pd.DataFrame([{"Variable": c, "Valeurs manquantes": int(df[c].isna().sum()),
                                      "Taux manquant (%)": round(df[c].isna().mean()*100, 2)}
                                     for c in selected])
                st.dataframe(miss, use_container_width=True)
                if selected:
                    st.dataframe(df[df[selected].isna().any(axis=1)], use_container_width=True)
                    csv_download(df[df[selected].isna().any(axis=1)], "anomalies_completude_ip15.csv", "Exporter les lignes à examiner")
            with tab2:
                keys = st.multiselect("Clé(s) de doublon", list(df.columns))
                if keys:
                    dup = df[df.duplicated(subset=keys, keep=False)].sort_values(keys)
                    st.metric("Lignes impliquées dans des doublons", len(dup))
                    st.dataframe(dup, use_container_width=True)
                    csv_download(dup, "doublons_ip15.csv", "Exporter les doublons")
            with tab3:
                nums = df.select_dtypes(include=np.number).columns.tolist()
                if nums:
                    col = st.selectbox("Variable numérique", nums)
                    s = df[col].dropna()
                    q1, q3 = s.quantile(.25), s.quantile(.75)
                    iqr = q3-q1
                    lower, upper = q1-1.5*iqr, q3+1.5*iqr
                    outliers = df[(df[col] < lower) | (df[col] > upper)]
                    st.write(f"Bornes exploratoires IQR : {lower:.3f} à {upper:.3f}")
                    st.metric("Valeurs atypiques signalées", len(outliers))
                    st.dataframe(outliers, use_container_width=True)
                else: st.info("Aucune variable numérique détectée.")
            with tab4:
                summary = df.describe(include="all").transpose().reset_index(names="Variable")
                st.dataframe(summary, use_container_width=True)
                csv_download(summary, "synthese_donnees_ip15.csv", "Exporter la synthèse")
        except Exception as e: st.error(f"Fichier non exploitable : {e}")

# =============================================================================
# 11. PAGE 6 – ASSISTANT MÉTHODOLOGIQUE IA
# =============================================================================

elif page == PAGES[5]:
    page_header("Assistant méthodologique IA", "Recherche dans les 50 indicateurs et appui à la structuration des travaux.")
    st.warning("Ne saisissez pas de données personnelles sensibles ou d'informations confidentielles non autorisées. Toute réponse générée doit être contrôlée par un professionnel.")
    question = st.text_area("Votre question", placeholder="Ex. : Quels travaux préparer pour vérifier l'accessibilité des transports ?")
    api_key = st.text_input("Clé API OpenAI (facultatif, non enregistrée)", type="password")
    model_name = st.text_input("Modèle API", value="gpt-4o-mini")
    if st.button("Analyser la question", type="primary"):
        if not question.strip():
            st.error("Saisissez une question.")
        else:
            terms = [t.lower() for t in re.findall(r"\w+", question) if len(t) > 3]
            scored = []
            for item in INDICATORS:
                text = (item["Indicateur"] + " " + item["Domaine"] + f" article {item['Article']}").lower()
                score = sum(1 for t in terms if t in text)
                if score: scored.append((score, item))
            relevant = [x[1] for x in sorted(scored, key=lambda z: z[0], reverse=True)[:12]]
            context = "\n".join([f"{x['Code']} – Art. {x['Article']} – {x['Indicateur']}" for x in relevant])
            if api_key:
                try:
                    from openai import OpenAI
                    client = OpenAI(api_key=api_key)
                    prompt = f"""Tu es un assistant méthodologique pour une mission de vérification au Sénégal.
Référence du projet : {LEGAL_REFERENCE}. Réponds en français, avec prudence.
N'invente aucun article, texte réglementaire, seuil, fait ou preuve. Distingue les pistes de contrôle
des constats établis. L'inspecteur garde la responsabilité de l'analyse.
Question : {question}
Indicateurs du référentiel potentiellement pertinents :
{context or 'Aucun indicateur identifié par recherche lexicale.'}
Propose : (1) cadrage, (2) critères à vérifier à confirmer dans le texte officiel,
(3) documents à demander, (4) tests/entretiens, (5) preuves attendues et limites."""
                    response = client.chat.completions.create(
                        model=model_name, messages=[{"role": "user", "content": prompt}], temperature=0.2
                    )
                    answer = response.choices[0].message.content
                except Exception as e:
                    answer = f"Mode IA indisponible ({e}). Résultat local présenté ci-dessous."
            else:
                answer = (
                    "Mode local sans génération IA.\n\n"
                    "Démarche suggérée :\n"
                    "1. Définir l'entité, la période, les bénéficiaires et le périmètre de vérification.\n"
                    "2. Confirmer dans la version officielle de la loi et les textes d'application les obligations applicables.\n"
                    "3. Demander les procédures, statistiques, budgets, décisions, registres et rapports pertinents.\n"
                    "4. Réaliser des tests sur pièces, entretiens et observations, en recherchant des preuves traçables.\n"
                    "5. Documenter les écarts, recueillir les explications contradictoires et formuler des recommandations mesurables.\n\n"
                    "Cette réponse est une trame générale et non une analyse juridique automatisée."
                )
            st.markdown("### Réponse")
            st.write(answer)
            st.session_state["ip15_last_answer"] = answer
            if relevant:
                st.markdown("### Indicateurs rapprochés par recherche lexicale")
                st.dataframe(pd.DataFrame(relevant), use_container_width=True, hide_index=True)
            st.download_button("Télécharger la réponse TXT", answer.encode("utf-8"), "assistant_ip15.txt", "text/plain")

# =============================================================================
# 12. PAGE 7 – PRÉPARATION DE MISSION
# =============================================================================

elif page == PAGES[6]:
    page_header("Préparation de mission", "Construire le programme de travail à partir des obligations et des 50 indicateurs.")
    missions = mission_options()
    selected_mission = st.selectbox("Mission concernée", list(missions.keys()))
    mission_id = missions[selected_mission]
    with st.form("prep_form"):
        indicator_label = st.selectbox("Indicateur à préparer", list(indicator_options().keys()))
        code = indicator_options()[indicator_label]
        col1, col2 = st.columns(2)
        objectif = col1.text_area("Objectif de vérification")
        critere = col2.text_area("Critères / obligations à confirmer dans les textes applicables")
        docs_req = st.text_area("Documents et données à demander")
        procedures = st.text_area("Procédures, tests et entretiens envisagés")
        preuves = st.text_area("Preuves attendues / sources possibles")
        responsable = st.text_input("Responsable du test")
        if st.form_submit_button("Enregistrer dans la matrice de préparation"):
            content = json.dumps({"objectif": objectif, "critere": critere, "documents": docs_req,
                                  "procedures": procedures, "preuves": preuves, "responsable": responsable},
                                 ensure_ascii=False)
            execute("INSERT INTO matrix_rows(categorie,code,contenu,statut,justification,date_modification) VALUES(?,?,?,?,?,?)",
                    ("Programme de travail", code, content, "À préparer", str(mission_id or ""), datetime.now().isoformat()))
            log_action("Ajout", "Programme de travail", code)
            st.success("Élément de préparation enregistré.")
    prep = qdf("SELECT * FROM matrix_rows WHERE categorie='Programme de travail' ORDER BY id DESC")
    if not prep.empty:
        st.markdown("### Programme de travail enregistré")
        st.dataframe(prep, use_container_width=True, hide_index=True)
        csv_download(prep, "programme_travail_ip15.csv", "Exporter le programme")

# =============================================================================
# 13. PAGE 8 – TRAVAUX DE TERRAIN
# =============================================================================

elif page == PAGES[7]:
    page_header("Travaux de terrain", "Consigner les tests, les pièces examinées, les constats et la contradiction.")
    missions = mission_options()
    selected_mission = st.selectbox("Mission", list(missions.keys()), key="field_mission")
    mission_id = missions[selected_mission]
    org = st.text_input("Organisme / service vérifié")
    options = indicator_options()
    indicator_label = st.selectbox("Indicateur évalué", list(options.keys()), key="field_indicator")
    code = options[indicator_label]
    with st.form("evaluation_form"):
        status = st.selectbox("Résultat provisoire", [
            "Non évalué", "Conforme", "Partiellement conforme", "Non conforme",
            "Non applicable (à justifier)", "Information insuffisante"
        ])
        score = st.slider("Score interne de maturité / mise en œuvre (0–100)", 0, 100, 50)
        constat = st.text_area("Constat factuel")
        preuve = st.text_area("Références des pièces / preuves examinées")
        source = st.text_area("Sources, entretiens et observations")
        responsable = st.text_input("Agent chargé du test")
        commentaire = st.text_area("Réponse de l'organisme / contradiction / limites")
        if st.form_submit_button("Enregistrer l'évaluation", type="primary"):
            execute("""INSERT INTO evaluations(mission_id,organisme,code_indicateur,statut,score,constat,
                       preuve,source,responsable,date_evaluation,commentaire) VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                    (mission_id, org, code, status, score, constat, preuve, source, responsable,
                     date.today().isoformat(), commentaire))
            log_action("Évaluation", code, f"{org} – {status}")
            st.success("Évaluation enregistrée. Le résultat reste à valider dans le cadre de la mission.")
    evals = qdf("SELECT * FROM evaluations ORDER BY id DESC")
    if not evals.empty:
        st.markdown("### Registre des évaluations")
        st.dataframe(evals, use_container_width=True, hide_index=True)
        csv_download(evals, "evaluations_ip15.csv", "Exporter les évaluations")

# =============================================================================
# 14. PAGE 9 – BIBLIOTHÈQUE JURIDIQUE
# =============================================================================

elif page == PAGES[8]:
    page_header("Bibliothèque juridique", "Centraliser les textes, décisions, circulaires, procédures et documents de référence.")
    st.info("Le script fourni identifie la loi n° 2010-15 du 6 juillet 2010. Les autres références doivent être ajoutées et validées par l'équipe de mission.")
    with st.form("legal_meta"):
        files = st.file_uploader("Déposer un ou plusieurs documents", type=[
            "pdf", "doc", "docx", "odt", "xls", "xlsx", "ods", "csv", "ppt", "pptx", "odp", "txt", "rtf"
        ], accept_multiple_files=True)
        c1, c2 = st.columns(2)
        family = c1.selectbox("Famille documentaire", [
            "Loi et textes législatifs", "Décrets et textes réglementaires",
            "Arrêtés et décisions", "Conventions et normes internationales",
            "Politiques et stratégies", "Procédures internes", "Rapports et études",
            "Données statistiques", "Autres documents"
        ])
        ref = c2.text_input("Référence officielle")
        authority = c1.text_input("Autorité émettrice")
        date_text = c2.text_input("Date du texte (si connue)")
        status = st.selectbox("Statut documentaire", ["À vérifier", "En vigueur (à confirmer)", "Abrogé", "Projet", "Archive"])
        keywords = st.text_input("Mots-clés")
        save_docs = st.form_submit_button("Indexer les documents")
    if save_docs:
        if not files:
            st.error("Sélectionnez au moins un document.")
        else:
            count = 0
            for f in files:
                raw = f.getvalue()
                digest = hashlib.sha256(raw).hexdigest()
                stored = f"{datetime.now():%Y%m%d_%H%M%S}_{digest[:10]}_{safe_filename(f.name)}"
                folder = UPLOAD_DIR / re.sub(r"[^A-Za-z0-9_-]", "_", family)
                folder.mkdir(parents=True, exist_ok=True)
                path = folder / stored
                path.write_bytes(raw)
                execute("""INSERT INTO legal_docs(nom_original,nom_stocke,famille,reference_officielle,
                           autorite,date_texte,statut,mots_cles,chemin,empreinte,date_ajout)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                        (f.name, stored, family, ref, authority, date_text, status, keywords,
                         str(path), digest, datetime.now().isoformat(timespec="seconds")))
                count += 1
            log_action("Indexation", "Bibliothèque juridique", f"{count} document(s)")
            st.success(f"{count} document(s) indexé(s).")
    docs = qdf("SELECT * FROM legal_docs ORDER BY id DESC")
    if not docs.empty:
        search = st.text_input("Rechercher par nom, référence, autorité ou mot-clé")
        fam_filter = st.selectbox("Filtrer par famille", ["Toutes"] + sorted(docs["famille"].dropna().unique().tolist()))
        view = docs.copy()
        if search:
            mask = view.astype(str).apply(lambda col: col.str.contains(search, case=False, na=False)).any(axis=1)
            view = view[mask]
        if fam_filter != "Toutes": view = view[view["famille"] == fam_filter]
        st.dataframe(view.drop(columns=["chemin"], errors="ignore"), use_container_width=True, hide_index=True)
        csv_download(view.drop(columns=["chemin"], errors="ignore"), "catalogue_juridique_ip15.csv", "Exporter le catalogue")
        for r in view.itertuples():
            p = Path(r.chemin)
            if p.exists():
                with st.expander(f"{r.nom_original} · {r.famille}"):
                    st.caption(f"{r.reference_officielle or 'Référence non renseignée'} · {r.statut}")
                    st.download_button("Télécharger le document", p.read_bytes(), r.nom_original, key=f"doc_{r.id}")

# =============================================================================
# 15. PAGE 10 – RÉFÉRENTIEL DES 50 INDICATEURS
# =============================================================================

elif page == PAGES[9]:
    page_header("Référentiel des 50 indicateurs", "Référentiel de travail structuré selon les sept chapitres et les articles 1 à 50.")
    c1, c2, c3 = st.columns(3)
    chapter_filter = c1.selectbox("Chapitre", ["Tous"] + [f"{x['id']} · {x['name']}" for x in CHAPTERS])
    code_filter = c2.text_input("Recherche par code")
    text_filter = c3.text_input("Recherche dans les libellés")
    view = INDICATOR_DF.copy()
    if chapter_filter != "Tous":
        chap_id = chapter_filter.split(" · ")[0]
        view = view[view["Chapitre"] == chap_id]
    if code_filter: view = view[view["Code"].str.contains(code_filter, case=False)]
    if text_filter: view = view[view["Indicateur"].str.contains(text_filter, case=False, regex=False)]
    st.metric("Indicateurs affichés", len(view))
    st.dataframe(view, use_container_width=True, hide_index=True)
    csv_download(INDICATOR_DF, "referentiel_50_indicateurs_ip15.csv", "Télécharger les 50 indicateurs")
    st.markdown("### Répartition par chapitre")
    st.dataframe(pd.DataFrame([{
        "Chapitre": c["id"], "Intitulé": c["name"], "Articles": c["articles"],
        "Nombre d'indicateurs": len(c["items"])
    } for c in CHAPTERS]), use_container_width=True, hide_index=True)
    st.caption("Les libellés ci-dessus sont ceux du script de référentiel fourni. Les critères détaillés, modalités de calcul, seuils et références d'application restent à documenter et à valider.")

# =============================================================================
# 16. PAGE 11 – MATRICES & QUESTIONNAIRES
# =============================================================================

elif page == PAGES[10]:
    page_header("Matrices & questionnaires", "Structurer l'applicabilité, les critères, les preuves, les entretiens et les arbitrages.")
    tabs = st.tabs(["Applicabilité", "Obligations & contrôles", "Questionnaire", "Écarts & arbitrages", "Historique"])
    with tabs[0]:
        st.markdown("#### Matrice d'applicabilité")
        with st.form("applicability"):
            code = st.selectbox("Indicateur", list(indicator_options().keys()), key="matrix_code")
            applicable = st.selectbox("Applicabilité", ["À déterminer", "Applicable", "Non applicable – justification requise"])
            rationale = st.text_area("Justification et contexte institutionnel")
            if st.form_submit_button("Enregistrer l'applicabilité"):
                execute("INSERT INTO matrix_rows(categorie,code,contenu,statut,justification,date_modification) VALUES(?,?,?,?,?,?)",
                        ("Applicabilité", indicator_options()[code], rationale, applicable, "", datetime.now().isoformat()))
                log_action("Mise à jour", "Applicabilité", indicator_options()[code])
                st.success("Matrice mise à jour.")
        d = qdf("SELECT * FROM matrix_rows WHERE categorie='Applicabilité' ORDER BY id DESC")
        if not d.empty: st.dataframe(d, use_container_width=True, hide_index=True)
    with tabs[1]:
        st.markdown("#### Matrice obligations – risques – critères – preuves")
        with st.form("obligation_matrix"):
            code_label = st.selectbox("Indicateur", list(indicator_options().keys()), key="obl_code")
            obligation = st.text_area("Obligation / exigence à confirmer dans le texte officiel")
            risk = st.text_area("Risque d'inapplication ou d'ineffectivité")
            criterion = st.text_area("Critère de vérification")
            procedure = st.text_area("Procédure / test")
            evidence = st.text_area("Preuves attendues")
            matrix_status = st.selectbox("État", ["À documenter", "En cours", "Validé", "À arbitrer"])
            if st.form_submit_button("Enregistrer la matrice"):
                payload = json.dumps({"obligation": obligation, "risque": risk, "critere": criterion,
                                      "procedure": procedure, "preuve": evidence}, ensure_ascii=False)
                execute("INSERT INTO matrix_rows(categorie,code,contenu,statut,justification,date_modification) VALUES(?,?,?,?,?,?)",
                        ("Obligations-contrôles", indicator_options()[code_label], payload, matrix_status, "",
                         datetime.now().isoformat()))
                log_action("Ajout", "Obligations-contrôles", indicator_options()[code_label])
                st.success("Ligne enregistrée.")
        d = qdf("SELECT * FROM matrix_rows WHERE categorie='Obligations-contrôles' ORDER BY id DESC")
        if not d.empty: st.dataframe(d, use_container_width=True, hide_index=True)
    with tabs[2]:
        st.markdown("#### Questionnaire institutionnel")
        missions = mission_options()
        mlabel = st.selectbox("Mission", list(missions.keys()), key="quest_mission")
        with st.form("questionnaire"):
            organism = st.text_input("Organisme / service")
            domain = st.selectbox("Domaine", [c["name"] for c in CHAPTERS])
            question = st.text_area("Question")
            answer = st.text_area("Réponse de l'interlocuteur")
            comment = st.text_area("Commentaire / pièce justificative / point à approfondir")
            if st.form_submit_button("Enregistrer la réponse"):
                execute("INSERT INTO questionnaires(mission_id,organisme,domaine,question,reponse,commentaire,date_saisie) VALUES(?,?,?,?,?,?,?)",
                        (missions[mlabel], organism, domain, question, answer, comment, datetime.now().isoformat()))
                log_action("Réponse", "Questionnaire", organism)
                st.success("Réponse enregistrée.")
        qs = qdf("SELECT * FROM questionnaires ORDER BY id DESC")
        if not qs.empty: st.dataframe(qs, use_container_width=True, hide_index=True)
    with tabs[3]:
        st.markdown("#### Registre des écarts et arbitrages")
        with st.form("gap_form"):
            gap_code = st.selectbox("Indicateur", list(indicator_options().keys()), key="gap_code")
            issue = st.text_area("Écart / question à arbitrer")
            position = st.text_area("Analyse et éléments contradictoires")
            decision = st.selectbox("Statut", ["Ouvert", "En analyse", "Arbitré", "Clôturé"])
            rationale = st.text_area("Motivation / référence de la décision")
            if st.form_submit_button("Enregistrer l'arbitrage"):
                execute("INSERT INTO matrix_rows(categorie,code,contenu,statut,justification,date_modification) VALUES(?,?,?,?,?,?)",
                        ("Écart-arbitrage", indicator_options()[gap_code],
                         json.dumps({"écart": issue, "analyse": position}, ensure_ascii=False),
                         decision, rationale, datetime.now().isoformat()))
                log_action("Arbitrage", "Écart", indicator_options()[gap_code])
                st.success("Élément enregistré.")
        gaps = qdf("SELECT * FROM matrix_rows WHERE categorie='Écart-arbitrage' ORDER BY id DESC")
        if not gaps.empty: st.dataframe(gaps, use_container_width=True, hide_index=True)
    with tabs[4]:
        logs = qdf("SELECT * FROM audit_log ORDER BY id DESC LIMIT 300")
        if logs.empty: st.info("Aucune opération enregistrée.")
        else: st.dataframe(logs, use_container_width=True, hide_index=True)

# =============================================================================
# 17. PAGE 12 – SUIVI DES RECOMMANDATIONS
# =============================================================================

elif page == PAGES[11]:
    page_header("Suivi des recommandations", "Formaliser les actions correctives, responsables, échéances et preuves d'exécution.")
    missions = mission_options()
    with st.form("recommendation_form"):
        mission_label = st.selectbox("Mission", list(missions.keys()), key="rec_mission")
        c1, c2 = st.columns(2)
        ref = c1.text_input("Référence de la recommandation")
        priority = c2.selectbox("Priorité", ["Critique", "Élevée", "Moyenne", "Faible"])
        organism = st.text_input("Organisme responsable")
        recommendation = st.text_area("Libellé de la recommandation")
        c3, c4 = st.columns(2)
        owner = c3.text_input("Responsable de mise en œuvre")
        deadline = c4.date_input("Échéance", value=date.today())
        status = st.selectbox("Statut", ["Ouverte", "En cours", "Mise en œuvre déclarée", "À vérifier", "Clôturée"])
        evidence = st.text_area("Preuve d'exécution / référence du document")
        if st.form_submit_button("Enregistrer la recommandation", type="primary"):
            execute("""INSERT INTO recommendations(mission_id,reference,organisme,recommandation,priorite,
                       responsable,echeance,statut,preuve_execution,date_creation,historique)
                       VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                    (missions[mission_label], ref, organism, recommendation, priority, owner,
                     deadline.isoformat(), status, evidence, date.today().isoformat(), "Création"))
            log_action("Création", "Recommandation", ref or recommendation[:80])
            st.success("Recommandation enregistrée.")
    recs = qdf("SELECT * FROM recommendations ORDER BY id DESC")
    if not recs.empty:
        st.markdown("### Registre des recommandations")
        st.dataframe(recs, use_container_width=True, hide_index=True)
        csv_download(recs, "suivi_recommandations_ip15.csv", "Exporter le registre")
        st.markdown("### Mise à jour du statut")
        ids = recs["id"].astype(int).tolist()
        rid = st.selectbox("Recommandation à actualiser", ids)
        new_status = st.selectbox("Nouveau statut", ["Ouverte", "En cours", "Mise en œuvre déclarée", "À vérifier", "Clôturée"], key="new_rec_status")
        new_evidence = st.text_input("Nouvelle preuve / commentaire")
        if st.button("Mettre à jour", type="primary"):
            execute("UPDATE recommendations SET statut=?, preuve_execution=?, historique=? WHERE id=?",
                    (new_status, new_evidence, f"Mis à jour le {date.today().isoformat()}", int(rid)))
            log_action("Mise à jour", "Recommandation", str(rid))
            st.success("Statut mis à jour.")
            st.rerun()

# =============================================================================
# 18. PAGE 13 – RAPPORTS & EXPORTS
# =============================================================================

elif page == PAGES[12]:
    page_header("Rapports & exports", "Produire les extractions de travail et les synthèses disponibles à partir des données saisies.")
    missions = get_missions()
    evals = qdf("SELECT * FROM evaluations")
    recs = qdf("SELECT * FROM recommendations")
    docs = qdf("SELECT * FROM legal_docs")
    tabs = st.tabs(["Synthèse", "Missions", "Évaluations", "Recommandations", "Bibliothèque"])
    with tabs[0]:
        st.markdown("### Synthèse de l'activité enregistrée")
        metric_row([
            ("Missions", len(missions), "Base locale"),
            ("Évaluations", len(evals), "Enregistrements de terrain"),
            ("Recommandations", len(recs), "Registre de suivi"),
            ("Documents", len(docs), "Bibliothèque juridique"),
        ])
        if not evals.empty:
            by_status = evals.groupby("statut").size().reset_index(name="Nombre")
            if PLOTLY_OK:
                fig = px.pie(by_status, names="statut", values="Nombre", hole=.45)
                st.plotly_chart(fig, use_container_width=True)
            else: st.dataframe(by_status)
        st.caption("Cette synthèse reflète uniquement les informations enregistrées dans cette installation.")
    with tabs[1]:
        if not missions.empty:
            st.dataframe(missions, use_container_width=True, hide_index=True)
            csv_download(missions, "missions_ip15.csv", "Exporter les missions")
        else: st.info("Aucune mission.")
    with tabs[2]:
        if not evals.empty:
            st.dataframe(evals, use_container_width=True, hide_index=True)
            csv_download(evals, "evaluations_ip15.csv", "Exporter les évaluations")
        else: st.info("Aucune évaluation.")
    with tabs[3]:
        if not recs.empty:
            st.dataframe(recs, use_container_width=True, hide_index=True)
            csv_download(recs, "recommandations_ip15.csv", "Exporter les recommandations")
        else: st.info("Aucune recommandation.")
    with tabs[4]:
        if not docs.empty:
            st.dataframe(docs.drop(columns=["chemin"], errors="ignore"), use_container_width=True, hide_index=True)
            csv_download(docs.drop(columns=["chemin"], errors="ignore"), "catalogue_juridique_ip15.csv", "Exporter le catalogue")
        else: st.info("Aucun document indexé.")

# =============================================================================
# 19. PAGE 14 – PARAMÈTRES & JOURNAL
# =============================================================================

elif page == PAGES[13]:
    page_header("Paramètres & journal", "Informations techniques, sauvegarde locale et traçabilité des opérations.")
    st.markdown("### Configuration")
    st.write(f"**Dossier de données :** `{BASE_DIR.resolve()}`")
    st.write(f"**Base SQLite :** `{DB_PATH.resolve()}`")
    st.write(f"**Référentiel :** {LEGAL_REFERENCE}")
    st.write(f"**Indicateurs chargés :** {len(INDICATORS)}")
    st.write(f"**Scikit-learn :** {'Disponible' if SKLEARN_OK else 'Non installé'}")
    st.write(f"**Plotly :** {'Disponible' if PLOTLY_OK else 'Non installé'}")
    st.markdown("### Export de sauvegarde de la base")
    if DB_PATH.exists():
        st.download_button("Télécharger la base SQLite", DB_PATH.read_bytes(),
                           "sauvegarde_inclusion_professionnelle_15.sqlite3",
                           "application/x-sqlite3")
    st.markdown("### Journal des opérations")
    logs = qdf("SELECT * FROM audit_log ORDER BY id DESC LIMIT 500")
    if logs.empty: st.info("Aucune opération.")
    else:
        st.dataframe(logs, use_container_width=True, hide_index=True)
        csv_download(logs, "journal_ip15.csv", "Exporter le journal")
    st.markdown("### Principes de gouvernance")
    st.markdown("""
    - Les données de mission doivent être exactes, sourcées et limitées au besoin professionnel.
    - Les textes doivent être vérifiés dans leur version officielle et leur champ d'application.
    - Les indicateurs constituent un cadre de vérification ; ils ne valent pas preuve d'un manquement.
    - Les conclusions doivent être fondées sur des éléments probants, contradictoires et traçables.
    - L'assistance algorithmique ne remplace pas le jugement et la responsabilité des professionnels.
    """)


# =============================================================================
# 22. TOME VI – VÉRIFICATION PRÉVENTIVE ET RÉFÉRENTIELS INTERNATIONAUX
# =============================================================================
# Extension intégrée à l'application IP15. Les constantes ci-dessous proviennent
# du script Tome VI transmis par l'utilisateur. Les données opérationnelles
# (projets, contrôles, preuves, risques, anomalies, recommandations) sont saisies
# par l'utilisateur : aucun projet fictif ni résultat de conformité n'est créé.
# =============================================================================

T6_CHAPITRES = [{'numero': 1, 'titre': 'Dispositions générales', 'debut': 1, 'fin': 6},
 {'numero': 2, 'titre': 'Santé, action sociale et prévention', 'debut': 7, 'fin': 14},
 {'numero': 3, 'titre': 'Éducation, formation professionnelle et emploi', 'debut': 15, 'fin': 30},
 {'numero': 4,
  'titre': 'Accessibilité, habitat, cadre de vie, transport, communication et accès à la terre',
  'debut': 31,
  'fin': 38},
 {'numero': 5, 'titre': 'Sports, loisirs, arts et culture', 'debut': 39, 'fin': 42},
 {'numero': 6, 'titre': 'Organisation des personnes handicapées', 'debut': 43, 'fin': 44},
 {'numero': 7, 'titre': 'Dispositions transitoires et finales', 'debut': 45, 'fin': 50}]

T6_ETAPES = [['G01',
  'Identification et cadrage',
  'Avant la décision de lancement',
  'Identifier les obligations IP15 potentiellement applicables',
  'Note de cadrage, fiche projet, analyse des parties prenantes',
  'Applicabilité juridique documentée'],
 ['G02',
  'Maturation et études préalables',
  'Avant validation des études',
  "Vérifier l'intégration de l'inclusion dans les études",
  'Études de faisabilité, études sociales et techniques',
  'Risques et besoins identifiés'],
 ['G03',
  'Programmation et financement',
  'Avant arbitrage budgétaire',
  'Vérifier la prise en compte des besoins et coûts pertinents',
  'Budget, plan de financement, programmation',
  'Ressources et responsabilités identifiées'],
 ['G04',
  'Conception technique',
  'Avant validation des plans',
  "Examiner les dispositions d'accessibilité et d'inclusion",
  'Plans, cahiers techniques et avis spécialisés',
  'Exigences traduites dans la conception'],
 ['G05',
  'Passation des marchés',
  'Avant publication ou attribution',
  'Vérifier les exigences dans les dossiers de marché',
  'DAO, spécifications, critères et contrats',
  'Exigences pertinentes contractualisées'],
 ['G06',
  'Exécution des travaux ou prestations',
  'Pendant la réalisation',
  'Contrôler la mise en œuvre des dispositions prévues',
  'Rapports de chantier, PV, photos et contrôles',
  'Écarts détectés et corrigés'],
 ['G07',
  'Réception et mise en service',
  'Avant réception ou ouverture au public',
  "Vérifier la conformité et l'utilisabilité des réalisations",
  'PV de réception, essais, réserves et visites',
  'Réserves traitées avant mise en service']]

T6_REFERENTIELS = [{'Organisme': 'INTOSAI',
  'Référence': 'ISSAI 100',
  'Intitulé': "Principes fondamentaux de l'audit du secteur public",
  'Domaine': 'Audit public',
  'Nature': "Principes d'audit public",
  'Principes': 'Mandat, indépendance, jugement professionnel, assurance, risques, matérialité, '
               'éléments probants et rapport.',
  'Application IP15': 'Structurer la mission, définir les objectifs, les critères, les procédures, '
                      'les preuves et les conclusions.',
  'Étapes': 'G01 à G07',
  'Source': 'S01',
  'URL': 'https://www.issai.org/'},
 {'Organisme': 'INTOSAI',
  'Référence': 'ISSAI 130',
  'Intitulé': 'Code de déontologie',
  'Domaine': 'Éthique',
  'Nature': 'Référentiel déontologique',
  'Principes': 'Intégrité, indépendance, objectivité, compétence, comportement professionnel et '
               'confidentialité.',
  'Application IP15': "Prévenir les conflits d'intérêts et garantir l'impartialité des "
                      'vérificateurs.',
  'Étapes': 'G01 à G07',
  'Source': 'S01',
  'URL': 'https://www.issai.org/'},
 {'Organisme': 'INTOSAI',
  'Référence': 'ISSAI 200',
  'Intitulé': "Principes de l'audit financier",
  'Domaine': 'Audit financier',
  'Nature': "Principes d'audit financier",
  'Principes': "Risques d'anomalies, matérialité, éléments probants et examen des informations "
               'financières.',
  'Application IP15': 'Vérifier la budgétisation des aménagements accessibles, la réalité des '
                      'dépenses et leur traçabilité.',
  'Étapes': 'G03, G05, G06, G07',
  'Source': 'S01',
  'URL': 'https://www.issai.org/'},
 {'Organisme': 'INTOSAI',
  'Référence': 'ISSAI 300',
  'Intitulé': "Principes de l'audit de performance",
  'Domaine': 'Performance publique',
  'Nature': "Principes d'audit de performance",
  'Principes': 'Économie, efficience et efficacité ; analyse des résultats et de la valeur '
               "produite par l'action publique.",
  'Application IP15': 'Évaluer si les investissements produisent des infrastructures et services '
                      'effectivement accessibles et utilisables.',
  'Étapes': 'G01 à G07',
  'Source': 'S01',
  'URL': 'https://www.issai.org/'},
 {'Organisme': 'INTOSAI',
  'Référence': 'ISSAI 400',
  'Intitulé': "Principes de l'audit de conformité",
  'Domaine': 'Conformité',
  'Nature': "Principes d'audit de conformité",
  'Principes': 'Critères de conformité, objet de contrôle, éléments probants, évaluation des '
               'écarts et conclusions.',
  'Application IP15': 'Comparer les obligations pertinentes de la loi n°2010-15 aux décisions, '
                      'documents et réalisations du projet.',
  'Étapes': 'G01 à G07',
  'Source': 'S01',
  'URL': 'https://www.issai.org/'},
 {'Organisme': 'INTOSAI',
  'Référence': 'ISSAI 4000',
  'Intitulé': "Norme d'audit de conformité",
  'Domaine': 'Conformité',
  'Nature': "Norme d'audit de conformité",
  'Principes': 'Planification, identification des critères, collecte des preuves, évaluation des '
               'constats et rapport.',
  'Application IP15': 'Formaliser les tests de conformité IP15 et documenter les conclusions à '
                      'partir de preuves suffisantes.',
  'Étapes': 'G01 à G07',
  'Source': 'S01',
  'URL': 'https://www.issai.org/'},
 {'Organisme': 'IIA',
  'Référence': 'Global Internal Audit Standards 2024',
  'Intitulé': "Normes mondiales d'audit interne",
  'Domaine': 'Audit interne',
  'Nature': 'Normes professionnelles',
  'Principes': "15 principes organisés en cinq domaines : finalité de l'audit interne, éthique et "
               'professionnalisme, gouvernance, gestion et réalisation des missions.',
  'Application IP15': 'Renforcer la planification, la qualité des travaux, la communication des '
                      'résultats et le suivi des recommandations.',
  'Étapes': 'G01 à G07',
  'Source': 'S02',
  'URL': 'https://www.theiia.org/en/standards/'},
 {'Organisme': 'ISO',
  'Référence': 'ISO 19011:2026',
  'Intitulé': "Lignes directrices pour l'audit des systèmes de management",
  'Domaine': "Méthodologie d'audit",
  'Nature': 'Lignes directrices',
  'Principes': "Principes d'audit, gestion d'un programme d'audit, conduite des audits et "
               'compétence des auditeurs.',
  'Application IP15': 'Organiser les programmes de vérification, les équipes, les entretiens, les '
                      'inspections et les tests.',
  'Étapes': 'G01 à G07',
  'Source': 'S03',
  'URL': 'https://www.iso.org/'},
 {'Organisme': 'COSO',
  'Référence': 'Internal Control – Integrated Framework 2013',
  'Intitulé': 'Cadre intégré de contrôle interne',
  'Domaine': 'Contrôle interne',
  'Nature': 'Cadre méthodologique',
  'Principes': 'Cinq composantes : environnement de contrôle, évaluation des risques, activités de '
               'contrôle, information et communication, activités de pilotage.',
  'Application IP15': "Évaluer la gouvernance, les contrôles du maître d'ouvrage et le suivi des "
                      "exigences d'accessibilité.",
  'Étapes': 'G01 à G07',
  'Source': 'S04',
  'URL': 'https://www.coso.org/'},
 {'Organisme': 'ISO',
  'Référence': 'ISO 31000:2018',
  'Intitulé': 'Management du risque – Lignes directrices',
  'Domaine': 'Gestion des risques',
  'Nature': 'Lignes directrices',
  'Principes': 'Identification, analyse, évaluation, traitement, surveillance et communication des '
               'risques.',
  'Application IP15': "Construire la cartographie des risques d'exclusion, d'inaccessibilité, de "
                      'surcoût et de correction tardive.',
  'Étapes': 'G01 à G07',
  'Source': 'S05',
  'URL': 'https://www.iso.org/standard/65694.html'},
 {'Organisme': 'ISO',
  'Référence': 'ISO 37301:2021',
  'Intitulé': 'Systèmes de management de la conformité',
  'Domaine': 'Conformité',
  'Nature': 'Norme de management',
  'Principes': 'Identification des obligations, responsabilités, maîtrise opérationnelle, '
               'évaluation et amélioration continue.',
  'Application IP15': 'Mettre en place une veille des obligations IP15, des contrôles de '
                      'conformité et des actions correctives.',
  'Étapes': 'G01 à G07',
  'Source': 'S06',
  'URL': 'https://www.iso.org/standard/75080.html'},
 {'Organisme': 'ISO',
  'Référence': 'ISO 37000:2021',
  'Intitulé': 'Gouvernance des organisations',
  'Domaine': 'Gouvernance',
  'Nature': 'Lignes directrices',
  'Principes': 'Gouvernance responsable, finalité organisationnelle, supervision, responsabilité '
               'et création de valeur.',
  'Application IP15': "Vérifier la prise en compte de l'inclusion dans les arbitrages du maître "
                      "d'ouvrage et les mécanismes de supervision.",
  'Étapes': 'G01, G02, G03, G07',
  'Source': 'S07',
  'URL': 'https://www.iso.org/'},
 {'Organisme': 'ISO',
  'Référence': 'ISO 37001:2025',
  'Intitulé': 'Systèmes de management anti-corruption',
  'Domaine': 'Intégrité',
  'Nature': 'Norme de management',
  'Principes': 'Prévention, détection et traitement des risques de corruption ; diligence, '
               'contrôles et amélioration.',
  'Application IP15': "Examiner les risques d'intégrité dans la sélection des prestataires, les "
                      'marchés et les contrôles des prestations.',
  'Étapes': 'G03, G05, G06',
  'Source': 'S08',
  'URL': 'https://www.iso.org/'},
 {'Organisme': 'PEFA',
  'Référence': 'Cadre PEFA 2016',
  'Intitulé': 'Évaluation de la gestion des finances publiques',
  'Domaine': 'Finances publiques',
  'Nature': "Cadre d'évaluation",
  'Principes': 'Sept piliers, 31 indicateurs et 94 dimensions : fiabilité budgétaire, '
               'transparence, actifs et passifs, exécution, comptabilité, reporting et contrôle '
               'externe.',
  'Application IP15': "Examiner la programmation, la traçabilité des crédits, l'exécution "
                      "budgétaire et le contrôle des dépenses d'inclusion.",
  'Étapes': 'G03, G05, G06, G07',
  'Source': 'S09',
  'URL': 'https://www.pefa.org/'},
 {'Organisme': 'IPSASB',
  'Référence': 'Handbook 2025',
  'Intitulé': 'Normes comptables internationales du secteur public',
  'Domaine': 'Comptabilité publique',
  'Nature': 'Référentiel comptable',
  'Principes': "Normes IPSAS et cadre conceptuel de l'information financière des entités du "
               'secteur public ; édition couvrant les normes publiées au 31 janvier 2025.',
  'Application IP15': "Éclairer l'examen de la comptabilisation, de la présentation et du suivi "
                      'financier des investissements publics.',
  'Étapes': 'G03, G06, G07',
  'Source': 'S10',
  'URL': 'https://www.ipsasb.org/'},
 {'Organisme': 'UEMOA',
  'Référence': 'Directive n°01/2009',
  'Intitulé': 'Code de transparence dans la gestion des finances publiques',
  'Domaine': 'Finances publiques',
  'Nature': 'Texte communautaire',
  'Principes': 'Transparence dans la gestion des finances publiques et information budgétaire.',
  'Application IP15': "Examiner la transparence de l'information budgétaire et financière du "
                      "projet, sous réserve de l'applicabilité nationale.",
  'Étapes': 'G03, G06, G07',
  'Source': 'S11',
  'URL': 'https://www.uemoa.int/'},
 {'Organisme': 'UEMOA',
  'Référence': 'Directive n°06/2009',
  'Intitulé': 'Lois de finances',
  'Domaine': 'Finances publiques',
  'Nature': 'Texte communautaire',
  'Principes': "Cadre relatif aux lois de finances et à l'organisation des processus budgétaires.",
  'Application IP15': "Examiner la programmation, l'autorisation et l'exécution des crédits "
                      'mobilisés pour le projet.',
  'Étapes': 'G03, G06',
  'Source': 'S11',
  'URL': 'https://www.uemoa.int/'},
 {'Organisme': 'UEMOA',
  'Référence': 'Directive n°07/2009',
  'Intitulé': 'Règlement général sur la comptabilité publique',
  'Domaine': 'Comptabilité publique',
  'Nature': 'Texte communautaire',
  'Principes': 'Règles de comptabilité et de gestion des deniers publics.',
  'Application IP15': 'Examiner les opérations comptables et les contrôles financiers pertinents '
                      'pour les dépenses du projet.',
  'Étapes': 'G03, G06, G07',
  'Source': 'S11',
  'URL': 'https://www.uemoa.int/'},
 {'Organisme': 'Sénégal',
  'Référence': 'Loi n°2010-15 du 6 juillet 2010',
  'Intitulé': "Loi d'orientation sociale",
  'Domaine': 'Inclusion et handicap',
  'Nature': 'Loi nationale',
  'Principes': "Loi d'orientation sociale relative à la promotion et à la protection des droits "
               'des personnes handicapées.',
  'Application IP15': 'Identifier les articles applicables, les obligations, les acteurs '
                      'responsables et les critères de vérification.',
  'Étapes': 'G01 à G07 selon applicabilité',
  'Source': 'S12',
  'URL': 'À renseigner depuis le Journal officiel du Sénégal'},
 {'Organisme': 'Sénégal',
  'Référence': 'Décret n°2022-2295',
  'Intitulé': 'Code des marchés publics',
  'Domaine': 'Commande publique',
  'Nature': 'Réglementation nationale',
  'Principes': "Cadre national de passation, d'attribution et d'exécution des marchés publics.",
  'Application IP15': 'Examiner, selon le régime applicable, la traduction des exigences '
                      'pertinentes dans les dossiers de consultation et les marchés.',
  'Étapes': 'G03, G05, G06',
  'Source': 'S13',
  'URL': 'À renseigner depuis le Journal officiel du Sénégal'},
 {'Organisme': 'Sénégal',
  'Référence': 'Loi n°2021-23',
  'Intitulé': 'Contrats de partenariat public-privé',
  'Domaine': 'Partenariats public-privé',
  'Nature': 'Loi nationale',
  'Principes': 'Cadre juridique des contrats de partenariat public-privé.',
  'Application IP15': "Vérifier l'intégration des obligations pertinentes dans le montage "
                      'contractuel, uniquement si le projet relève du régime PPP.',
  'Étapes': 'G02, G03, G05, G06',
  'Source': 'S14',
  'URL': 'À renseigner depuis le Journal officiel du Sénégal'}]

T6_SOURCES_OFFICIELLES = [['S01', 'INTOSAI', 'https://www.issai.org/'],
 ['S02', 'IIA', 'https://www.theiia.org/en/standards/'],
 ['S03', 'ISO', 'https://www.iso.org/'],
 ['S04', 'COSO', 'https://www.coso.org/'],
 ['S05', 'ISO 31000', 'https://www.iso.org/standard/65694.html'],
 ['S06', 'ISO 37301', 'https://www.iso.org/standard/75080.html'],
 ['S07', 'ISO', 'https://www.iso.org/'],
 ['S08', 'ISO', 'https://www.iso.org/'],
 ['S09', 'PEFA', 'https://www.pefa.org/'],
 ['S10', 'IPSASB', 'https://www.ipsasb.org/'],
 ['S11', 'UEMOA', 'https://www.uemoa.int/'],
 ['S12', 'Journal officiel du Sénégal', 'À compléter'],
 ['S13', 'Journal officiel du Sénégal', 'À compléter'],
 ['S14', 'Journal officiel du Sénégal', 'À compléter']]

T6_COLONNES_PROJETS = ['Code projet',
 'Nom du projet',
 "Maître d'ouvrage",
 'Secteur',
 'Localisation',
 'Budget prévisionnel FCFA',
 'Phase actuelle',
 'Date de lancement',
 'Responsable',
 'Statut IP15',
 'Observations']

T6_COLONNES_PREUVES = ['ID preuve',
 'Code projet',
 'Article IP15',
 'Point de contrôle',
 'Nature de la preuve',
 'Description',
 'Source / emplacement',
 'Date de collecte',
 'Collecté par',
 'Fiabilité',
 'Statut de validation',
 'Observations']

T6_COLONNES_RISQUES = ['ID risque',
 'Code projet',
 'Article IP15',
 'Point de contrôle',
 'Risque identifié',
 'Cause',
 'Conséquence potentielle',
 'Probabilité (1-5)',
 'Impact (1-5)',
 'Criticité',
 'Mesure de maîtrise',
 'Responsable',
 'Échéance',
 'Statut']

T6_COLONNES_ANOMALIES = ['ID anomalie',
 'Code projet',
 'Article IP15',
 'Point de contrôle',
 'Constat',
 'Critère',
 'Cause',
 'Conséquence',
 'Niveau',
 'Preuve associée',
 'Responsable',
 'Échéance',
 'Statut',
 'Observation']

T6_COLONNES_RECOMMANDATIONS = ['ID recommandation',
 'Code projet',
 'Article IP15',
 'Anomalie associée',
 'Recommandation',
 'Responsable',
 'Échéance',
 'Indicateur de réalisation',
 'Statut',
 'Preuve de clôture',
 'Validation du vérificateur']

T6_COLONNES_GATES = ['Code',
 'Étape du projet',
 'Moment du contrôle',
 'Objectif de vérification',
 'Documents et preuves attendus',
 'Résultat attendu',
 'Projet concerné',
 'Date du contrôle',
 'Vérificateur',
 'Conclusion',
 'Observations']

T6_LISTES = {'Applicabilite': ['Applicable', 'Non applicable', 'À examiner'],
 'Statut_verification': ['Non évalué',
                         'En cours',
                         'Conforme',
                         'Conforme avec réserve',
                         'Non conforme',
                         'Insuffisance documentaire',
                         'Non applicable'],
 'Conclusion': ['Avis favorable',
                'Avis favorable avec réserves',
                'Alerte préventive majeure',
                'À approfondir'],
 'Niveau_risque': ['Faible', 'Modéré', 'Élevé', 'Critique'],
 'Fiabilite': ['Faible', 'Moyenne', 'Élevée'],
 'Statut_recommandation': ['Non démarrée',
                           'En cours',
                           'Mise en œuvre',
                           'À vérifier',
                           'Clôturée',
                           'En retard'],
 'Probabilite_impact': [1, 2, 3, 4, 5]}

T6_PAGE = "TOME VI – Vérification préventive des grands projets"

T6_COLONNES_ARTICLES = [
    "Code article", "Article", "Chapitre", "Intitulé du chapitre",
    "Indicateur Tome II", "Libellé indicateur", "Texte de la loi (source Tome I)",
    "Applicabilité au projet", "Justification de l'applicabilité",
    "Statut de vérification", "Critère / procédure", "Référence de preuve",
    "Observations", "Responsable", "Date de mise à jour"
]

T6_COLONNES_MAPPING = [
    "Organisme", "Référentiel", "Nature", "Point de contrôle",
    "Étape du projet", "Application opérationnelle", "Réserve méthodologique"
]

T6_COLONNES_PLAN = [
    "Point de contrôle", "Étape du projet", "Moment du contrôle",
    "Code article", "Article", "Indicateur Tome II", "Objectif du test",
    "Procédure de vérification", "Preuve attendue", "Résultat", "Observations"
]

T6_STATUTS_APPLICABILITE = ["À examiner", "Applicable", "Non applicable (motivé)", "Applicabilité conditionnelle"]
T6_STATUTS_VERIFICATION = [
    "Non évalué", "En cours", "Conforme", "Conforme avec réserve",
    "Non conforme", "Insuffisance documentaire", "Non applicable"
]
T6_STATUTS_CONCLUSION = [
    "Avis favorable", "Avis favorable avec réserves",
    "Alerte préventive majeure", "À approfondir"
]
T6_STATUTS_SUIVI = ["Non démarrée", "En cours", "Mise en œuvre", "À vérifier", "Clôturée", "En retard"]

def t6_init_db():
    with db() as con:
        con.executescript("""
        CREATE TABLE IF NOT EXISTS t6_projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code_projet TEXT UNIQUE NOT NULL, nom_projet TEXT NOT NULL,
            maitre_ouvrage TEXT, secteur TEXT, localisation TEXT,
            budget_fcfa REAL, phase_actuelle TEXT, date_lancement TEXT,
            responsable TEXT, statut_ip15 TEXT, observations TEXT,
            created_at TEXT
        );
        CREATE TABLE IF NOT EXISTS t6_article_reviews (
            id INTEGER PRIMARY KEY AUTOINCREMENT, code_projet TEXT NOT NULL,
            article INTEGER NOT NULL, applicabilite TEXT, justification TEXT,
            statut TEXT, critere TEXT, preuve TEXT, observations TEXT,
            responsable TEXT, date_maj TEXT
        );
        CREATE TABLE IF NOT EXISTS t6_gate_checks (
            id INTEGER PRIMARY KEY AUTOINCREMENT, code_projet TEXT NOT NULL,
            gate_code TEXT NOT NULL, date_controle TEXT, verificateur TEXT,
            conclusion TEXT, observations TEXT, date_creation TEXT
        );
        CREATE TABLE IF NOT EXISTS t6_evidence (
            id INTEGER PRIMARY KEY AUTOINCREMENT, id_preuve TEXT,
            code_projet TEXT NOT NULL, article TEXT, point_controle TEXT,
            nature TEXT, description TEXT, source_emplacement TEXT,
            date_collecte TEXT, collecteur TEXT, fiabilite TEXT,
            statut_validation TEXT, observations TEXT, date_creation TEXT
        );
        CREATE TABLE IF NOT EXISTS t6_risks (
            id INTEGER PRIMARY KEY AUTOINCREMENT, id_risque TEXT,
            code_projet TEXT NOT NULL, article TEXT, point_controle TEXT,
            risque TEXT, cause TEXT, consequence TEXT, probabilite INTEGER,
            impact INTEGER, criticite INTEGER, niveau TEXT,
            mesure_maitrise TEXT, responsable TEXT, echeance TEXT,
            statut TEXT, date_creation TEXT
        );
        CREATE TABLE IF NOT EXISTS t6_anomalies (
            id INTEGER PRIMARY KEY AUTOINCREMENT, id_anomalie TEXT,
            code_projet TEXT NOT NULL, article TEXT, point_controle TEXT,
            constat TEXT, critere TEXT, cause TEXT, consequence TEXT,
            niveau TEXT, preuve_associee TEXT, responsable TEXT,
            echeance TEXT, statut TEXT, observation TEXT, date_creation TEXT
        );
        CREATE TABLE IF NOT EXISTS t6_recommendations (
            id INTEGER PRIMARY KEY AUTOINCREMENT, id_recommandation TEXT,
            code_projet TEXT NOT NULL, article TEXT, anomalie_associee TEXT,
            recommandation TEXT, responsable TEXT, echeance TEXT,
            indicateur_realisation TEXT, statut TEXT, preuve_cloture TEXT,
            validation_verificateur TEXT, date_creation TEXT
        );
        """)
t6_init_db()

def t6_df(query, params=()):
    return qdf(query, params)

def t6_projects_df():
    return t6_df("SELECT * FROM t6_projects ORDER BY id DESC")

def t6_project_choices():
    projects = t6_projects_df()
    if projects.empty:
        return {}
    return {
        f"{r.code_projet} · {r.nom_projet}": r.code_projet
        for r in projects.itertuples()
    }

def t6_articles_df():
    rows = []
    for chapter in T6_CHAPITRES:
        for article in range(chapter["debut"], chapter["fin"] + 1):
            indicator = INDICATOR_DF[INDICATOR_DF["Article"] == article]
            indicator_code = indicator.iloc[0]["Code"] if not indicator.empty else f"IP15-{article:02d}"
            indicator_label = indicator.iloc[0]["Indicateur"] if not indicator.empty else ""
            rows.append({
                "Code article": f"IP15-ART-{article:03d}",
                "Article": article,
                "Chapitre": chapter["numero"],
                "Intitulé du chapitre": chapter["titre"],
                "Indicateur Tome II": indicator_code,
                "Libellé indicateur": indicator_label,
                "Texte de la loi (source Tome I)": "À renseigner depuis le texte officiel et le Tome I ; ne pas reconstituer le texte de loi par approximation.",
            })
    return pd.DataFrame(rows)

def t6_mapping_df():
    rows = []
    for ref in T6_REFERENTIELS:
        for gate in T6_ETAPES:
            stages = ref["Étapes"]
            if gate[0] in stages or stages in ("G01 à G07", "G01 à G07 selon applicabilité"):
                rows.append([
                    ref["Organisme"], ref["Référence"], ref["Nature"], gate[0],
                    gate[1], ref["Application IP15"],
                    "Correspondance méthodologique à confirmer selon le mandat"
                ])
    return pd.DataFrame(rows, columns=T6_COLONNES_MAPPING)

def t6_plan_df():
    rows = []
    for gate in T6_ETAPES:
        for article in range(1, 51):
            rows.append([
                gate[0], gate[1], gate[2], f"IP15-ART-{article:03d}", article,
                f"IP15-{article:02d}", "À définir à partir du Tome II",
                "À définir selon l'applicabilité", "À documenter",
                "Non évalué", ""
            ])
    return pd.DataFrame(rows, columns=T6_COLONNES_PLAN)

def t6_risk_level(score):
    if score >= 20: return "Critique"
    if score >= 15: return "Élevé"
    if score >= 8: return "Modéré"
    return "Faible"

def t6_export_workbook():
    projects = t6_projects_df()
    article_df = t6_articles_df()
    reviews = t6_df("SELECT * FROM t6_article_reviews ORDER BY code_projet, article")
    gates = t6_df("SELECT * FROM t6_gate_checks ORDER BY code_projet, gate_code")
    evidence = t6_df("SELECT * FROM t6_evidence ORDER BY id")
    risks = t6_df("SELECT * FROM t6_risks ORDER BY id")
    anomalies = t6_df("SELECT * FROM t6_anomalies ORDER BY id")
    recommendations = t6_df("SELECT * FROM t6_recommendations ORDER BY id")
    refs = pd.DataFrame(T6_REFERENTIELS)
    mapping = t6_mapping_df()
    plan = t6_plan_df()
    sources = pd.DataFrame(T6_SOURCES_OFFICIELLES, columns=["Code", "Organisme", "Adresse officielle"])
    gates_ref = pd.DataFrame(T6_ETAPES, columns=[
        "Code", "Étape du projet", "Moment du contrôle", "Objectif de vérification",
        "Documents et preuves attendus", "Résultat attendu"
    ])
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        pd.DataFrame([
            ["Objet", "Guide de vérification préventive de l'inclusion des personnes handicapées dans les grands projets de l'État"],
            ["Référence juridique centrale", LEGAL_REFERENCE],
            ["Architecture", "50 articles, 7 chapitres et 7 points de contrôle G01 à G07"],
            ["Référentiels", f"{len(T6_REFERENTIELS)} référentiels professionnels, internationaux et nationaux"],
            ["Principe", "Les normes internationales éclairent la méthode et ne remplacent pas le droit national."],
            ["Précaution", "Aucune conformité n'est présumée ; les éditions et versions des référentiels doivent être vérifiées."],
            ["Date de génération", datetime.now().strftime("%d/%m/%Y %H:%M")]
        ], columns=["Rubrique", "Description"]).to_excel(writer, index=False, sheet_name="Présentation")
        pd.DataFrame([
            ["Projets enregistrés", len(projects)],
            ["Articles évalués", len(reviews)],
            ["Contrôles préventifs documentés", len(gates)],
            ["Preuves collectées", len(evidence)],
            ["Risques enregistrés", len(risks)],
            ["Anomalies enregistrées", len(anomalies)],
            ["Recommandations enregistrées", len(recommendations)],
            ["Recommandations clôturées", int((recommendations["statut"] == "Clôturée").sum()) if not recommendations.empty else 0],
        ], columns=["Indicateur", "Valeur"]).to_excel(writer, index=False, sheet_name="Tableau de bord")
        exports = [
            ("Portefeuille projets", projects), ("Matrice IP15", article_df),
            ("Revues articles", reviews), ("Contrôles préventifs", gates),
            ("Plan de vérification", plan), ("Référentiels", refs),
            ("Mapping normes étapes", mapping), ("Sources officielles", sources),
            ("Registre preuves", evidence), ("Registre risques", risks),
            ("Anomalies", anomalies), ("Recommandations", recommendations),
            ("Étapes G01-G07", gates_ref)
        ]
        for name, frame in exports:
            frame.to_excel(writer, index=False, sheet_name=name[:31])
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.sheet_view.showGridLines = False
            for cell in ws[1]:
                cell.font = __import__("openpyxl").styles.Font(bold=True, color="FFFFFF")
                cell.fill = __import__("openpyxl").styles.PatternFill("solid", fgColor="10243A")
                cell.alignment = __import__("openpyxl").styles.Alignment(wrap_text=True, vertical="center")
            for col in ws.columns:
                letter = col[0].column_letter
                max_len = max((len(str(c.value)) for c in col if c.value is not None), default=10)
                ws.column_dimensions[letter].width = min(max(max_len + 2, 14), 48)
        from openpyxl.chart import BarChart, Reference
        chart_data = writer.book.create_sheet("Données graphiques")
        chart_data.append(["Chapitre", "Nombre d'articles"])
        for chapter in T6_CHAPITRES:
            chart_data.append([chapter["titre"], chapter["fin"] - chapter["debut"] + 1])
        chart = BarChart()
        chart.type = "bar"
        chart.style = 10
        chart.title = "Répartition des articles IP15 par chapitre"
        chart.y_axis.title = "Chapitre"
        chart.x_axis.title = "Nombre d'articles"
        chart.height = 9
        chart.width = 18
        chart.add_data(Reference(chart_data, min_col=2, min_row=1, max_row=8), titles_from_data=True)
        chart.set_categories(Reference(chart_data, min_col=1, min_row=2, max_row=8))
        writer.book["Tableau de bord"].add_chart(chart, "D2")
        chart_data.sheet_state = "hidden"
        from openpyxl.worksheet.datavalidation import DataValidation
        validations = [
            ("Revues articles", "D", "À examiner,Applicable,Non applicable (motivé),Applicabilité conditionnelle"),
            ("Revues articles", "F", "Non évalué,En cours,Conforme,Conforme avec réserve,Non conforme,Insuffisance documentaire,Non applicable"),
            ("Contrôles préventifs", "F", "Avis favorable,Avis favorable avec réserves,Alerte préventive majeure,À approfondir"),
            ("Registre preuves", "K", "Faible,Moyenne,Élevée"),
            ("Registre risques", "I", "1,2,3,4,5"),
            ("Registre risques", "J", "1,2,3,4,5"),
            ("Anomalies", "J", "Faible,Modéré,Élevé,Critique"),
            ("Recommandations", "J", "Non démarrée,En cours,Mise en œuvre,À vérifier,Clôturée,En retard"),
        ]
        for sheet_name, column, options in validations:
            dv = DataValidation(type="list", formula1=f'"{options}"', allow_blank=True)
            dv.error = "Choisissez une valeur proposée dans la liste."
            writer.book[sheet_name].add_data_validation(dv)
            dv.add(f"{column}2:{column}501")
    return buffer.getvalue()

def t6_export_csv_zip():
    import zipfile
    frames = {
        "portefeuille_projets": t6_projects_df(),
        "matrice_50_articles": t6_articles_df(),
        "revues_articles": t6_df("SELECT * FROM t6_article_reviews"),
        "controles_preventifs": t6_df("SELECT * FROM t6_gate_checks"),
        "plan_verification": t6_plan_df(),
        "referentiels": pd.DataFrame(T6_REFERENTIELS),
        "mapping_normes_etapes": t6_mapping_df(),
        "sources_officielles": pd.DataFrame(T6_SOURCES_OFFICIELLES, columns=["Code","Organisme","Adresse officielle"]),
        "preuves": t6_df("SELECT * FROM t6_evidence"),
        "risques": t6_df("SELECT * FROM t6_risks"),
        "anomalies": t6_df("SELECT * FROM t6_anomalies"),
        "recommandations": t6_df("SELECT * FROM t6_recommendations"),
    }
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, frame in frames.items():
            zf.writestr(f"{name}.csv", frame.to_csv(index=False, encoding="utf-8-sig"))
    return buffer.getvalue()

def t6_export_png_zip():
    """Génère quatre graphiques PNG haute résolution à partir du référentiel et des données saisies."""
    import zipfile
    import matplotlib.pyplot as plt
    from matplotlib.patches import Patch
    output = io.BytesIO()
    refs = pd.DataFrame(T6_REFERENTIELS)
    chapters = pd.DataFrame([
        {"Chapitre": f"Chapitre {x['numero']} – {x['titre']}", "Articles": x["fin"]-x["debut"]+1}
        for x in T6_CHAPITRES
    ])
    gates = pd.DataFrame(T6_ETAPES, columns=["Code","Étape","Moment","Objectif","Preuves","Résultat"])
    colors = ["#10243A", "#245B83", "#239B91", "#C9A227", "#667085", "#27864B", "#7B5AA6"]
    images = {}
    # 1. Cycle G01-G07
    fig, ax = plt.subplots(figsize=(16, 5))
    xs = np.arange(len(gates))
    ax.plot(xs, np.ones(len(xs)), color="#245B83", marker="o", linewidth=3, markersize=13)
    for i, row in gates.iterrows():
        ax.text(i, 1.06, row["Code"], ha="center", va="bottom", fontweight="bold", color="#10243A")
        ax.text(i, .94, row["Étape"], ha="center", va="top", fontsize=8, wrap=True, color="#10243A")
    ax.set_title("Cycle des sept points de contrôle préventif IP15", color="#10243A", pad=24)
    ax.set_xlim(-.4, 6.4); ax.set_ylim(.75, 1.25); ax.axis("off")
    fig.tight_layout(); buf=io.BytesIO(); fig.savefig(buf,format="png",dpi=300,bbox_inches="tight"); plt.close(fig)
    images["01_cycle_verification_preventive.png"] = buf.getvalue()
    # 2. Répartition des articles
    fig, ax = plt.subplots(figsize=(14, 7))
    ax.barh(chapters["Chapitre"], chapters["Articles"], color="#245B83")
    ax.set_title("Répartition des 50 articles de la loi n°2010-15 par chapitre", color="#10243A")
    ax.set_xlabel("Nombre d'articles"); ax.grid(axis="x",alpha=.2); fig.tight_layout()
    buf=io.BytesIO(); fig.savefig(buf,format="png",dpi=300,bbox_inches="tight"); plt.close(fig)
    images["02_repartition_50_articles.png"] = buf.getvalue()
    # 3. Cartographie des référentiels par domaine
    domain_counts = refs.groupby("Domaine").size().sort_values()
    fig, ax = plt.subplots(figsize=(14, 8))
    ax.barh(domain_counts.index, domain_counts.values, color="#239B91")
    ax.set_title("Cartographie des référentiels professionnels et internationaux", color="#10243A")
    ax.set_xlabel("Nombre de références"); ax.grid(axis="x",alpha=.2); fig.tight_layout()
    buf=io.BytesIO(); fig.savefig(buf,format="png",dpi=300,bbox_inches="tight"); plt.close(fig)
    images["03_cartographie_referentiels.png"] = buf.getvalue()
    # 4. Chaîne de traçabilité
    labels = ["Obligations", "Contrôles", "Preuves", "Risques / constats", "Recommandations", "Clôture"]
    fig, ax = plt.subplots(figsize=(15, 4))
    x = np.arange(len(labels))
    ax.plot(x, np.ones(len(x)), color="#10243A", linewidth=3, marker="o", markersize=13)
    for i, label in enumerate(labels):
        ax.text(i, 1.07, label, ha="center", va="bottom", fontsize=9, color="#10243A")
    ax.set_title("Chaîne de traçabilité et de responsabilisation IP15", color="#10243A", pad=24)
    ax.set_xlim(-.4, 5.4); ax.set_ylim(.85, 1.25); ax.axis("off")
    fig.tight_layout(); buf=io.BytesIO(); fig.savefig(buf,format="png",dpi=300,bbox_inches="tight"); plt.close(fig)
    images["04_chaine_tracabilite.png"] = buf.getvalue()
    with zipfile.ZipFile(output,"w",zipfile.ZIP_DEFLATED) as zf:
        for name, content in images.items():
            zf.writestr(name,content)
    return output.getvalue()

def t6_charts():
    projects = t6_projects_df()
    reviews = t6_df("SELECT * FROM t6_article_reviews")
    risks = t6_df("SELECT * FROM t6_risks")
    gates = t6_df("SELECT * FROM t6_gate_checks")
    refs = pd.DataFrame(T6_REFERENTIELS)
    st.markdown("#### Visualisations du dispositif préventif")
    if PLOTLY_OK:
        stages = pd.DataFrame(T6_ETAPES, columns=["Code","Étape","Moment","Objectif","Preuves","Résultat"])
        fig = px.scatter(stages, x=list(range(1, 8)), y=[1]*7, text="Code",
                         hover_name="Étape", title="Cycle des sept points de contrôle préventif IP15")
        fig.update_traces(marker=dict(size=18, color=C["teal"]), textposition="top center")
        fig.update_yaxes(visible=False)
        fig.update_xaxes(title="Progression du cycle projet", tickvals=list(range(1,8)), ticktext=stages["Code"])
        fig.update_layout(height=300, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
        chapter_counts = pd.DataFrame([
            {"Chapitre": f"Chapitre {x['numero']} – {x['titre']}",
             "Articles": x["fin"]-x["debut"]+1} for x in T6_CHAPITRES
        ])
        st.plotly_chart(px.bar(chapter_counts, x="Chapitre", y="Articles",
                               title="Répartition des 50 articles par chapitre",
                               color_discrete_sequence=[C["blue"]]), use_container_width=True)
        domain_counts = refs.groupby("Domaine", as_index=False).size().rename(columns={"size":"Référentiels"})
        st.plotly_chart(px.bar(domain_counts, x="Référentiels", y="Domaine", orientation="h",
                               title="Référentiels par domaine",
                               color_discrete_sequence=[C["teal"]]), use_container_width=True)
        chain = pd.DataFrame({"Étape":["Obligations","Contrôles","Preuves","Risques / constats","Recommandations","Clôture"],
                              "Ordre":[1,2,3,4,5,6]})
        fig_chain = px.line(chain, x="Ordre", y=[1]*6, text="Étape",
                            title="Chaîne de traçabilité et de responsabilisation")
        fig_chain.update_traces(mode="markers+lines+text", marker=dict(size=14, color=C["gold"]),
                                textposition="top center", line=dict(color=C["navy"], width=3))
        fig_chain.update_yaxes(visible=False)
        fig_chain.update_xaxes(visible=False)
        fig_chain.update_layout(height=300, showlegend=False)
        st.plotly_chart(fig_chain, use_container_width=True)
    else:
        st.info("Plotly n'est pas installé. Les tableaux restent disponibles ; installez plotly pour activer les graphiques interactifs.")
        st.bar_chart(pd.DataFrame({"Articles":[x["fin"]-x["debut"]+1 for x in T6_CHAPITRES]},
                                  index=[x["titre"] for x in T6_CHAPITRES]))
    if not gates.empty:
        st.markdown("**Contrôles renseignés par étape**")
        counts = gates.groupby("gate_code").size().reindex([x[0] for x in T6_ETAPES], fill_value=0)
        st.bar_chart(counts)
    if not risks.empty:
        st.markdown("**Répartition des risques enregistrés**")
        st.bar_chart(risks["niveau"].value_counts())
    if not reviews.empty:
        st.markdown("**Statuts des vérifications des articles**")
        st.bar_chart(reviews["statut"].value_counts())

def t6_render_page():
    page_header(
        "TOME VI – VÉRIFICATION PRÉVENTIVE IP15",
        "Normes professionnelles et référentiels internationaux · Inclusion des personnes handicapées dans les grands projets de l'État"
    )
    projects = t6_projects_df()
    evidence = t6_df("SELECT * FROM t6_evidence")
    risks = t6_df("SELECT * FROM t6_risks")
    anomalies = t6_df("SELECT * FROM t6_anomalies")
    recommendations = t6_df("SELECT * FROM t6_recommendations")
    reviews = t6_df("SELECT * FROM t6_article_reviews")
    gates = t6_df("SELECT * FROM t6_gate_checks")
    metric_row([
        ("Projets enregistrés", len(projects), "Aucun projet n'est préchargé."),
        ("Articles évalués", len(reviews), "Revues article/projet enregistrées."),
        ("Contrôles G01-G07", len(gates), "Contrôles préventifs documentés."),
        ("Preuves collectées", len(evidence), "Éléments probants enregistrés."),
        ("Risques", len(risks), "Risques enregistrés."),
        ("Anomalies", len(anomalies), "Constats enregistrés."),
    ])
    st.markdown(
        '<div class="ip15-note"><b>Règle de prudence :</b> les référentiels internationaux sont des appuis méthodologiques. '
        'Ils ne remplacent pas les obligations du droit sénégalais. L’applicabilité doit être motivée projet par projet ; '
        'une absence de pièce ne constitue pas automatiquement une non-conformité.</div>',
        unsafe_allow_html=True
    )
    tabs = st.tabs([
        "Tableau de bord", "Projets", "50 articles IP15", "Contrôles G01-G07",
        "Référentiels & mapping", "Preuves", "Risques", "Anomalies",
        "Recommandations", "Exports"
    ])

    with tabs[0]:
        t6_charts()
        st.markdown("#### Synthèse des référentiels intégrés")
        st.dataframe(pd.DataFrame(T6_REFERENTIELS)[["Organisme","Référence","Domaine","Étapes"]],
                     use_container_width=True, hide_index=True)
        st.caption("Les éditions et versions doivent être confirmées à la date de chaque mission.")

    with tabs[1]:
        st.markdown("#### Enregistrer un projet réel")
        with st.form("t6_project_form", clear_on_submit=True):
            c1,c2,c3 = st.columns(3)
            code = c1.text_input("Code projet *", placeholder="Ex. PROJ-2027-001")
            name = c2.text_input("Nom du projet *")
            owner = c3.text_input("Maître d'ouvrage")
            sector = c1.selectbox("Secteur", ["Infrastructures","Éducation","Santé","Transport","Habitat","Énergie","Eau et assainissement","Numérique","Sports et culture","Autre"])
            location = c2.text_input("Localisation")
            budget = c3.number_input("Budget prévisionnel (FCFA)", min_value=0.0, step=1000000.0)
            phase = c1.selectbox("Phase actuelle", [x[1] for x in T6_ETAPES])
            start_date = c2.date_input("Date de lancement", value=None)
            manager = c3.text_input("Responsable du projet")
            project_status = st.selectbox("Statut IP15", ["Non évalué","À instruire","En vérification","Avec réserves","Alerte préventive","Clôturé"])
            obs = st.text_area("Observations")
            submitted = st.form_submit_button("Enregistrer le projet", type="primary")
            if submitted:
                if not code.strip() or not name.strip():
                    st.error("Le code et le nom du projet sont obligatoires.")
                else:
                    try:
                        execute("""INSERT INTO t6_projects(code_projet,nom_projet,maitre_ouvrage,secteur,localisation,
                            budget_fcfa,phase_actuelle,date_lancement,responsable,statut_ip15,observations,created_at)
                            VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                            (code.strip(),name.strip(),owner,sector,location,budget,phase,
                             start_date.isoformat() if start_date else "",manager,project_status,obs,
                             datetime.now().isoformat(timespec="seconds")))
                        log_action("Création","Tome VI – projet",f"{code} · {name}")
                        st.success("Projet enregistré.")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("Ce code projet existe déjà.")
        if not projects.empty:
            st.dataframe(projects, use_container_width=True, hide_index=True)
            csv_download(projects, "tome6_portefeuille_projets.csv", "Exporter le portefeuille")
        else:
            st.info("Aucun projet n'est enregistré. Saisissez uniquement des projets réels.")

    with tabs[2]:
        st.markdown("#### Matrice de traçabilité des 50 articles")
        base_articles = t6_articles_df()
        project_map = t6_project_choices()
        if not project_map:
            st.warning("Enregistrez d'abord un projet pour documenter son applicabilité.")
        else:
            selected_project_label = st.selectbox("Projet concerné", list(project_map), key="t6_article_project")
            selected_project = project_map[selected_project_label]
            with st.form("t6_article_review_form", clear_on_submit=True):
                c1,c2,c3 = st.columns(3)
                article_no = c1.selectbox("Article", list(range(1,51)))
                app_status = c2.selectbox("Applicabilité", T6_STATUTS_APPLICABILITE)
                verify_status = c3.selectbox("Statut de vérification", T6_STATUTS_VERIFICATION)
                justification = st.text_area("Justification de l'applicabilité")
                criterion = st.text_area("Critère / procédure de vérification")
                proof_ref = st.text_input("Référence de preuve")
                reviewer = st.text_input("Responsable / vérificateur")
                comments = st.text_area("Observations")
                if st.form_submit_button("Enregistrer la revue de l'article"):
                    execute("""INSERT INTO t6_article_reviews(code_projet,article,applicabilite,justification,
                        statut,critere,preuve,observations,responsable,date_maj)
                        VALUES(?,?,?,?,?,?,?,?,?,?)""",
                        (selected_project,article_no,app_status,justification,verify_status,
                         criterion,proof_ref,comments,reviewer,datetime.now().isoformat(timespec="seconds")))
                    log_action("Création","Tome VI – revue article",f"{selected_project} · Article {article_no}")
                    st.success("Revue enregistrée.")
                    st.rerun()
            current = t6_df("SELECT * FROM t6_article_reviews WHERE code_projet=? ORDER BY article,id DESC",(selected_project,))
            if not current.empty:
                latest = current.drop_duplicates("article", keep="first")
                base_articles = base_articles.merge(latest[["article","applicabilite","justification","statut","critere","preuve","observations","responsable","date_maj"]],
                    left_on="Article",right_on="article",how="left").drop(columns=["article"])
            st.dataframe(base_articles, use_container_width=True, hide_index=True)
            csv_download(base_articles, f"{selected_project}_matrice_50_articles.csv", "Exporter la matrice du projet")
        with st.expander("Texte intégral de l'article sélectionné"):
            n = st.selectbox("Article à consulter", list(range(1,51)), key="t6_article_text")
            st.info("Le texte intégral de cet article n’est pas présent dans la version de base du script. Joignez ou renseignez la version officielle dans la bibliothèque juridique / Tome I avant de conclure.")

    with tabs[3]:
        st.markdown("#### Sept points de contrôle préventif")
        project_map = t6_project_choices()
        if not project_map:
            st.info("Enregistrez un projet avant d'ouvrir un contrôle préventif.")
        else:
            with st.form("t6_gate_form", clear_on_submit=True):
                proj_label = st.selectbox("Projet", list(project_map), key="t6_gate_project")
                gate_code_selected = st.selectbox(
                    "Point de contrôle",
                    [x[0] for x in T6_ETAPES],
                    format_func=lambda code: f"{code} – {next(x[1] for x in T6_ETAPES if x[0] == code)}"
                )
                gate = next(x for x in T6_ETAPES if x[0] == gate_code_selected)
                dcol1,dcol2 = st.columns(2)
                control_date = dcol1.date_input("Date du contrôle")
                checker = dcol2.text_input("Vérificateur")
                conclusion = st.selectbox("Conclusion", T6_STATUTS_CONCLUSION)
                gate_obs = st.text_area("Observations, réserves et références de pièces")
                if st.form_submit_button("Enregistrer le contrôle"):
                    execute("""INSERT INTO t6_gate_checks(code_projet,gate_code,date_controle,verificateur,
                        conclusion,observations,date_creation) VALUES(?,?,?,?,?,?,?)""",
                        (project_map[proj_label],gate[0],control_date.isoformat(),checker,conclusion,
                         gate_obs,datetime.now().isoformat(timespec="seconds")))
                    log_action("Création","Tome VI – contrôle préventif",f"{project_map[proj_label]} · {gate[0]}")
                    st.success("Contrôle préventif enregistré.")
                    st.rerun()
        gates = t6_df("SELECT * FROM t6_gate_checks ORDER BY id DESC")
        if not gates.empty: st.dataframe(gates,use_container_width=True,hide_index=True)
        gate_ref_df = pd.DataFrame(T6_ETAPES,columns=["Code","Étape du projet","Moment du contrôle","Objectif","Documents et preuves attendus","Résultat attendu"])
        st.dataframe(gate_ref_df,use_container_width=True,hide_index=True)

    with tabs[4]:
        st.markdown("#### Référentiels professionnels et internationaux")
        refs_df = pd.DataFrame(T6_REFERENTIELS)
        search = st.text_input("Rechercher un référentiel, organisme ou domaine", key="t6_ref_search")
        if search:
            mask = refs_df.astype(str).apply(lambda col: col.str.contains(search,case=False,na=False)).any(axis=1)
            refs_df = refs_df[mask]
        st.dataframe(refs_df,use_container_width=True,hide_index=True)
        st.markdown("#### Correspondance référentiels / étapes G01-G07")
        mapping_df = t6_mapping_df()
        st.dataframe(mapping_df,use_container_width=True,hide_index=True)
        st.markdown("#### Sources officielles")
        st.dataframe(pd.DataFrame(T6_SOURCES_OFFICIELLES,columns=["Code","Organisme","Adresse officielle"]),
                     use_container_width=True,hide_index=True)
        st.caption("Les intitulés, éditions, versions, dates d'entrée en vigueur et liens doivent être contrôlés sur les sites officiels avant utilisation.")

    with tabs[5]:
        st.markdown("#### Registre des éléments probants")
        project_map = t6_project_choices()
        if project_map:
            with st.form("t6_evidence_form", clear_on_submit=True):
                proj = st.selectbox("Projet",list(project_map),key="t6_ev_project")
                c1,c2,c3 = st.columns(3)
                evid_id = c1.text_input("ID preuve")
                article = c2.selectbox("Article",[""]+[str(i) for i in range(1,51)])
                gate_code = c3.selectbox("Point de contrôle",[""]+[x[0] for x in T6_ETAPES])
                nature = st.selectbox("Nature de la preuve",["Document","Observation physique","Entretien","Photographie","Test technique","Donnée administrative","Autre"])
                description = st.text_area("Description de la preuve")
                source = st.text_input("Source / emplacement")
                c1,c2,c3 = st.columns(3)
                collected = c1.date_input("Date de collecte")
                collector = c2.text_input("Collecté par")
                reliability = c3.selectbox("Fiabilité",["Faible","Moyenne","Élevée"])
                validation = st.selectbox("Statut de validation",["À valider","Validée","Rejetée","À compléter"])
                ev_obs = st.text_area("Observations")
                if st.form_submit_button("Enregistrer la preuve"):
                    execute("""INSERT INTO t6_evidence(id_preuve,code_projet,article,point_controle,nature,
                        description,source_emplacement,date_collecte,collecteur,fiabilite,statut_validation,observations,date_creation)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (evid_id or f"PREUVE-{datetime.now():%Y%m%d%H%M%S}",project_map[proj],article,gate_code,nature,
                         description,source,collected.isoformat(),collector,reliability,validation,ev_obs,
                         datetime.now().isoformat(timespec="seconds")))
                    log_action("Création","Tome VI – preuve",project_map[proj])
                    st.success("Preuve enregistrée."); st.rerun()
        else: st.info("Enregistrez un projet avant de collecter des preuves.")
        evidence = t6_df("SELECT * FROM t6_evidence ORDER BY id DESC")
        if not evidence.empty: st.dataframe(evidence,use_container_width=True,hide_index=True)

    with tabs[6]:
        st.markdown("#### Registre des risques")
        project_map = t6_project_choices()
        if project_map:
            with st.form("t6_risk_form", clear_on_submit=True):
                proj = st.selectbox("Projet",list(project_map),key="t6_risk_project")
                c1,c2,c3 = st.columns(3)
                risk_id = c1.text_input("ID risque")
                article = c2.selectbox("Article",[""]+[str(i) for i in range(1,51)],key="t6_risk_article")
                gate_code = c3.selectbox("Point de contrôle",[""]+[x[0] for x in T6_ETAPES],key="t6_risk_gate")
                risk_name = st.text_area("Risque identifié")
                cause = st.text_area("Cause")
                consequence = st.text_area("Conséquence potentielle")
                c1,c2 = st.columns(2)
                probability = c1.select_slider("Probabilité",options=[1,2,3,4,5],value=3)
                impact = c2.select_slider("Impact",options=[1,2,3,4,5],value=3)
                score = probability*impact
                st.metric("Criticité calculée",f"{score}/25",t6_risk_level(score))
                mitigation = st.text_area("Mesure de maîtrise")
                c1,c2,c3 = st.columns(3)
                owner = c1.text_input("Responsable")
                due = c2.date_input("Échéance")
                status = c3.selectbox("Statut",["Ouvert","En traitement","Maîtrisé","Accepté","Clos"])
                if st.form_submit_button("Enregistrer le risque"):
                    execute("""INSERT INTO t6_risks(id_risque,code_projet,article,point_controle,risque,cause,
                        consequence,probabilite,impact,criticite,niveau,mesure_maitrise,responsable,echeance,statut,date_creation)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (risk_id or f"RISQUE-{datetime.now():%Y%m%d%H%M%S}",project_map[proj],article,gate_code,
                         risk_name,cause,consequence,probability,impact,score,t6_risk_level(score),mitigation,
                         owner,due.isoformat(),status,datetime.now().isoformat(timespec="seconds")))
                    log_action("Création","Tome VI – risque",project_map[proj])
                    st.success("Risque enregistré."); st.rerun()
        else: st.info("Enregistrez un projet avant d'identifier les risques.")
        risks = t6_df("SELECT * FROM t6_risks ORDER BY criticite DESC,id DESC")
        if not risks.empty: st.dataframe(risks,use_container_width=True,hide_index=True)

    with tabs[7]:
        st.markdown("#### Registre des anomalies et constats")
        project_map = t6_project_choices()
        if project_map:
            with st.form("t6_anomaly_form", clear_on_submit=True):
                proj = st.selectbox("Projet",list(project_map),key="t6_anom_project")
                c1,c2,c3 = st.columns(3)
                anom_id = c1.text_input("ID anomalie")
                article = c2.selectbox("Article",[""]+[str(i) for i in range(1,51)],key="t6_anom_article")
                gate_code = c3.selectbox("Point de contrôle",[""]+[x[0] for x in T6_ETAPES],key="t6_anom_gate")
                finding = st.text_area("Constat")
                criterion = st.text_area("Critère / exigence")
                cause = st.text_area("Cause")
                consequence = st.text_area("Conséquence")
                c1,c2,c3 = st.columns(3)
                level = c1.selectbox("Niveau",["Faible","Modéré","Élevé","Critique"])
                proof = c2.text_input("Preuve associée")
                owner = c3.text_input("Responsable")
                due = st.date_input("Échéance")
                status = st.selectbox("Statut",["Ouverte","En analyse","Contestée","Corrigée","Clôturée"])
                obs = st.text_area("Observation")
                if st.form_submit_button("Enregistrer l'anomalie"):
                    execute("""INSERT INTO t6_anomalies(id_anomalie,code_projet,article,point_controle, constat,
                        critere,cause,consequence,niveau,preuve_associee,responsable,echeance,statut,observation,date_creation)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (anom_id or f"ANOM-{datetime.now():%Y%m%d%H%M%S}",project_map[proj],article,gate_code,
                         finding,criterion,cause,consequence,level,proof,owner,due.isoformat(),status,obs,
                         datetime.now().isoformat(timespec="seconds")))
                    log_action("Création","Tome VI – anomalie",project_map[proj])
                    st.success("Anomalie enregistrée."); st.rerun()
        else: st.info("Enregistrez un projet avant de documenter une anomalie.")
        anomalies = t6_df("SELECT * FROM t6_anomalies ORDER BY id DESC")
        if not anomalies.empty: st.dataframe(anomalies,use_container_width=True,hide_index=True)
        st.caption("Distinguer l'insuffisance documentaire, le risque potentiel et la non-conformité établie.")

    with tabs[8]:
        st.markdown("#### Suivi des recommandations")
        project_map = t6_project_choices()
        if project_map:
            with st.form("t6_recommendation_form", clear_on_submit=True):
                proj = st.selectbox("Projet",list(project_map),key="t6_rec_project")
                c1,c2,c3 = st.columns(3)
                rec_id = c1.text_input("ID recommandation")
                article = c2.selectbox("Article",[""]+[str(i) for i in range(1,51)],key="t6_rec_article")
                anomaly = c3.text_input("Anomalie associée")
                recommendation = st.text_area("Recommandation")
                c1,c2 = st.columns(2)
                owner = c1.text_input("Responsable")
                due = c2.date_input("Échéance")
                indicator = st.text_input("Indicateur de réalisation")
                status = st.selectbox("Statut",T6_STATUTS_SUIVI)
                proof = st.text_input("Preuve de clôture")
                validation = st.text_input("Validation du vérificateur")
                if st.form_submit_button("Enregistrer la recommandation"):
                    execute("""INSERT INTO t6_recommendations(id_recommandation,code_projet,article,anomalie_associee,
                        recommandation,responsable,echeance,indicateur_realisation,statut,preuve_cloture,
                        validation_verificateur,date_creation) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (rec_id or f"REC-{datetime.now():%Y%m%d%H%M%S}",project_map[proj],article,anomaly,
                         recommendation,owner,due.isoformat(),indicator,status,proof,validation,
                         datetime.now().isoformat(timespec="seconds")))
                    log_action("Création","Tome VI – recommandation",project_map[proj])
                    st.success("Recommandation enregistrée."); st.rerun()
        else: st.info("Enregistrez un projet avant de créer des recommandations.")
        recommendations = t6_df("SELECT * FROM t6_recommendations ORDER BY id DESC")
        if not recommendations.empty: st.dataframe(recommendations,use_container_width=True,hide_index=True)

    with tabs[9]:
        st.markdown("#### Exports consolidés du Tome VI")
        st.write("Le classeur rassemble les projets, les 50 articles, les contrôles, le plan de vérification, les référentiels, les preuves, les risques, les anomalies et les recommandations.")
        try:
            excel_bytes = t6_export_workbook()
            st.download_button("Télécharger le classeur Excel complet du Tome VI",
                excel_bytes,"TOME_VI_IP15_VERIFICATION_PREVENTIVE.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary")
        except Exception as exc:
            st.error(f"Erreur lors de la préparation du classeur Excel : {exc}")
        try:
            csv_bytes = t6_export_csv_zip()
            st.download_button("Télécharger tous les registres CSV (.zip)",csv_bytes,
                "TOME_VI_IP15_EXPORTS_CSV.zip","application/zip")
            png_bytes = t6_export_png_zip()
            st.download_button("Télécharger les quatre graphiques PNG haute résolution (.zip)",png_bytes,
                "TOME_VI_IP15_GRAPHIQUES_PNG.zip","application/zip")
        except Exception as exc:
            st.error(f"Erreur lors de la préparation des CSV : {exc}")
        st.markdown("#### Exports unitaires")
        for title, query, filename in [
            ("Projets","SELECT * FROM t6_projects","tome6_projets.csv"),
            ("Revues articles","SELECT * FROM t6_article_reviews","tome6_revues_articles.csv"),
            ("Contrôles préventifs","SELECT * FROM t6_gate_checks","tome6_controles_preventifs.csv"),
            ("Preuves","SELECT * FROM t6_evidence","tome6_preuves.csv"),
            ("Risques","SELECT * FROM t6_risks","tome6_risques.csv"),
            ("Anomalies","SELECT * FROM t6_anomalies","tome6_anomalies.csv"),
            ("Recommandations","SELECT * FROM t6_recommendations","tome6_recommandations.csv")
        ]:
            frame = t6_df(query)
            csv_download(frame,filename,f"Exporter : {title}")
        st.caption("Les exports ne constituent pas une certification de conformité. Les données restent dans la base SQLite locale de l'application.")

# Extension du menu sans supprimer les modules des Tomes I à V.
if T6_PAGE not in PAGES:
    PAGES.append(T6_PAGE)

# Affichage du module Tome VI.
if page == T6_PAGE:
    t6_render_page()


# =============================================================================
# 20. PIED DE PAGE
# =============================================================================

st.markdown("---")
st.caption(
    "INCLUSION PROFESSIONNELLE 15 · Projet de modernisation de la vérification publique · "
    "Référentiel de travail fondé sur les scripts fournis · "
    "MOHAMED FALILOU FALL"
)
