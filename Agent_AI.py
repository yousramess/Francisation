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

# -------------------------
# NAVIGATION
# -------------------------
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
# PAGE ACCUEIL
# -------------------------
def accueil():
    top1, top2 = st.columns([5, 1])

    with top1:
        st.title("Choisissez une action")
        st.write("Sélectionnez l’outil que vous souhaitez utiliser.")

    with top2:
        if os.path.exists("logo.png"):
            st.image("logo.png", width=180)

    st.write("")
    st.write("")

    espace1, col1, col2, espace2 = st.columns([1, 2, 2, 1])

    with col1:
        st.markdown(
            """
            <a class="fake-btn fake-btn-blue" href="/?page=conversion">
                📄 Convertir la liste PDF en Excel
            </a>
            """,
            unsafe_allow_html=True
        )

    with col2:
        st.markdown(
            """
            <a class="fake-btn fake-btn-blue" href="/?page=comparaison">
                🔍 Détecter les nouveaux étudiants
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
        if os.path.exists("logo.png"):
            st.image("logo.png", width=150)

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
            for page_pdf in pdf.pages:
                table = page_pdf.extract_table()
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
            st.dataframe(df_final, use_container_width=True)

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_final.to_excel(writer, index=False)
            output.seek(0)

            nom = os.path.splitext(uploaded_file.name)[0]

            st.download_button(
                label="Télécharger Excel",
                data=output.getvalue(),
                file_name=f"{nom}_Excel.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Aucun tableau valide trouvé dans le PDF.")

# -------------------------
# COMPARAISON - FONCTIONS
# -------------------------
def normaliser_nom_colonne(col):
    col = str(col).strip().lower()
    col = ''.join(
        c for c in unicodedata.normalize('NFD', col)
        if unicodedata.category(c) != 'Mn'
    )
    col = re.sub(r'[^a-z0-9]', '', col)
    return col


def trouver_colonne_ref_indiv(df):
    variantes = {
        "refindiv",
        "referenceindiv",
        "refindividuel",
        "refindividu"
    }

    for col in df.columns:
        if normaliser_nom_colonne(col) in variantes:
            return col

    return None


def nettoyer_ref(serie):
    return (
        serie.astype(str)
        .str.strip()
        .replace(["nan", "None", "none", "null", ""], pd.NA)
    )


def comparer_fichiers(df1, df2):
    col_ref_1 = trouver_colonne_ref_indiv(df1)
    col_ref_2 = trouver_colonne_ref_indiv(df2)

    if not col_ref_1:
        raise ValueError("Colonne 'Ref.Indiv' introuvable dans le fichier 1.")
    if not col_ref_2:
        raise ValueError("Colonne 'Ref.Indiv' introuvable dans le fichier 2.")

    df1 = df1.copy()
    df2 = df2.copy()

    df1[col_ref_1] = nettoyer_ref(df1[col_ref_1])
    df2[col_ref_2] = nettoyer_ref(df2[col_ref_2])

    refs_fichier1 = set(df1[col_ref_1].dropna().unique())

    nouvelles_lignes = df2[
        df2[col_ref_2].notna() & (~df2[col_ref_2].isin(refs_fichier1))
    ].copy()

    return nouvelles_lignes, col_ref_1, col_ref_2


def dataframe_to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output.getvalue()

# -------------------------
# PAGE COMPARAISON
# -------------------------
def app_comparaison():
    col1, col2 = st.columns([4, 1])

    with col1:
        if st.button("⬅ Retour", key="retour_comparaison"):
            changer_page("accueil")
        st.title("Comparaison de 2 fichiers Excel (Ref.Indiv)")

    with col2:
        if os.path.exists("logo.png"):
            st.image("logo.png", width=180)

    st.write(
        "Cette application compare deux fichiers Excel sur la colonne **Ref.Indiv** "
        "et extrait les lignes du **2e fichier** qui n'existent pas dans le **1er**."
    )

    col1_ui, col2_ui = st.columns(2)

    with col1_ui:
        fichier1 = st.file_uploader(
            "Fichier 1 (référence)",
            type=["xlsx", "xls"],
            key="fichier1"
        )

    with col2_ui:
        fichier2 = st.file_uploader(
            "Fichier 2 (à comparer)",
            type=["xlsx", "xls"],
            key="fichier2"
        )

    if fichier1 and fichier2:
        try:
            df1 = pd.read_excel(fichier1)
            df2 = pd.read_excel(fichier2)

            st.success("Fichiers chargés avec succès ✅")

            with st.expander("Aperçu fichier 1"):
                st.dataframe(df1.head(10), use_container_width=True)

            with st.expander("Aperçu fichier 2"):
                st.dataframe(df2.head(10), use_container_width=True)

            if st.button("Lancer la comparaison", key="btn_comparaison"):
                nouvelles_lignes, col_ref_1, col_ref_2 = comparer_fichiers(df1, df2)

                st.info(f"Colonne détectée fichier 1 : {col_ref_1}")
                st.info(f"Colonne détectée fichier 2 : {col_ref_2}")

                st.subheader("Résultat")
                st.write(f"Nombre de nouvelles lignes : **{len(nouvelles_lignes)}**")

                if len(nouvelles_lignes) > 0:
                    st.dataframe(nouvelles_lignes, use_container_width=True)

                    nom_original = os.path.splitext(fichier2.name)[0]
                    date_str = datetime.now().strftime("%Y%m%d")
                    nom_sortie = f"{nom_original}_New_{date_str}.xlsx"

                    excel_bytes = dataframe_to_excel_bytes(nouvelles_lignes)

                    st.download_button(
                        label="Télécharger le fichier résultat",
                        data=excel_bytes,
                        file_name=nom_sortie,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("Aucune nouvelle ligne à ajouter.")

        except Exception as e:
            st.error(f"Erreur : {e}")
    else:
        st.warning("Veuillez téléverser les 2 fichiers Excel.")

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
