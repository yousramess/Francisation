import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import unicodedata
import os
from io import BytesIO
from datetime import datetime

# -------------------------
# CONFIG
# -------------------------

st.set_page_config(
    page_title="Outils Excel",
    page_icon="logo.png",
    layout="wide"
)

# Navigation via query params
page = st.query_params.get("page", "accueil")

def changer_page(page_name):
    st.query_params["page"] = page_name
    st.rerun()

# -------------------------
# CSS GLOBAL
# -------------------------

st.markdown("""
<style>
.fake-btn {
    display: flex;
    align-items: center;
    justify-content: center;
    text-align: center;
    width: 100%;
    min-height: 120px;
    border-radius: 24px;
    text-decoration: none;
    color: white !important;
    font-size: 28px;
    font-weight: 700;
    box-shadow: 0 10px 24px rgba(0,0,0,0.18);
    transition: all 0.25s ease;
    padding: 20px;
}

.fake-btn:hover {
    transform: translateY(-4px);
    box-shadow: 0 16px 30px rgba(0,0,0,0.22);
    color: white !important;
}

.fake-btn-blue {
    background: #163d8f;
}

.fake-btn-blue:hover {
    background: #f58220;
}

.fake-btn-orange {
    background: #f58220;
}

.fake-btn-orange:hover {
    background: #163d8f;
}
</style>
""", unsafe_allow_html=True)

# -------------------------
# NAVIGATION
# -------------------------
if "page" not in st.session_state:
    st.session_state.page = "accueil"

def changer_page(page):
    st.session_state.page = page

# -------------------------
# PAGE ACCUEIL
# -------------------------
def accueil():
    top1, top2 = st.columns([5, 1])

    with top1:
        st.title("Choisissez une action")
        st.write("Sélectionnez l’outil que vous souhaitez utiliser.")

    with top2:
        st.image("logo.png", width=250)

    st.write("")
    st.write("")

    espace1, col1, col2, espace2 = st.columns([1, 2, 2, 1])

    with col1:
        st.markdown(
            """
            <a class="fake-btn fake-btn-blue" href="/?page=conversion">
                📄 Convertir PDF en Excel
            </a>
            """,
            unsafe_allow_html=True
        )

    with col2:
        st.markdown(
            """
            <a class="fake-btn fake-btn-blue" href="/?page=comparaison">
                🔍 Comparer deux fichiers Excel
            </a>
            """,
            unsafe_allow_html=True
        )

# -------------------------
# CONVERSION
# -------------------------
def split_nom_prenom(nom_complet):
    if pd.isna(nom_complet):
        return "", ""

    texte = str(nom_complet).strip()

    if "," in texte:
        nom_famille, prenom = texte.split(",", 1)
        return nom_famille.strip(), prenom.strip()

    morceaux = texte.split()
    if len(morceaux) == 1:
        return morceaux[0], ""
    return morceaux[-1], " ".join(morceaux[:-1])


def app_conversion():
    col1, col2 = st.columns([4, 1])

    with col1:
        if st.button("⬅ Retour", key="retour_conversion"):
           changer_page("accueil")
        st.title("Outil de conversion PDF vers Excel")

    with col2:
        st.image("logo.png", width=180)

    st.subheader("PDF → Excel")

    uploaded_file = st.file_uploader("Téléverse ton fichier PDF", type=["pdf"])

    if uploaded_file:
        all_rows = []

        headers = [
            "Personne",
            "Nom, Prénom",
            "S",
            "Adresse courriel",
            "Téléphone maison",
            "Céllulaire ou Téléphone autre"
        ]

        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    for row in table:
                        if row and len(row) == 6:
                            all_rows.append(row)

        if all_rows:
            df_source = pd.DataFrame(all_rows, columns=headers)

            lignes_finales = []

            for _, row in df_source.iterrows():
                nom_famille, prenom = split_nom_prenom(row["Nom, Prénom"])

                lignes_finales.append({
                    "Ref.Indiv": row["Personne"],
                    "Nom": row["Nom, Prénom"],
                    "Nom de famille": nom_famille,
                    "Prénom": prenom,
                    "Genre": row["S"],
                    "Email": row["Adresse courriel"],
                    "Mobile": row["Céllulaire ou Téléphone autre"],
                    "Téléphone autre": row["Téléphone maison"],
                    "Francisation": "VRAI",
                    "Usager CARI": "VRAI",
                    "Étiquette": ""
                })

            df_final = pd.DataFrame(lignes_finales)

            st.success("Conversion réussie ✅")
            st.dataframe(df_final)

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_final.to_excel(writer, index=False)

            nom = os.path.splitext(uploaded_file.name)[0]

            st.download_button(
                "Télécharger Excel",
                output.getvalue(),
                f"{nom}_Excel.xlsx"
            )

# -------------------------
# COMPARAISON
# -------------------------
def normaliser_nom_colonne(col):
    col = str(col).lower()
    col = ''.join(
        c for c in unicodedata.normalize('NFD', col)
        if unicodedata.category(c) != 'Mn'
    )
    col = re.sub(r'[^a-z0-9]', '', col)
    return col


def trouver_colonne(df):
    for col in df.columns:
        if "refindiv" in normaliser_nom_colonne(col):
            return col
    return None


def app_comparaison():
    col1, col2 = st.columns([4, 1])

    with col1:
        if st.button("⬅ Retour", key="retour_comparaison"):
           changer_page("accueil")
        st.title("Comparaison Excel")

    with col2:
        st.image("logo.png", width=180)

    f1 = st.file_uploader("Fichier 1", type=["xlsx"])
    f2 = st.file_uploader("Fichier 2", type=["xlsx"])

    if f1 and f2:
        df1 = pd.read_excel(f1)
        df2 = pd.read_excel(f2)

        col1 = trouver_colonne(df1)
        col2 = trouver_colonne(df2)

        nouvelles = df2[~df2[col2].isin(df1[col1])]

        st.write(f"Nouvelles lignes : {len(nouvelles)}")
        st.dataframe(nouvelles)

# -------------------------
# ROUTER
# -------------------------
if page == "accueil":
    accueil()
elif page == "conversion":
    app_conversion()
elif page == "comparaison":
    app_comparaison()
else:
    accueil()
