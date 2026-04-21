import streamlit as st

# Navigation
if "page" not in st.session_state:
    st.session_state.page = "accueil"

# -------------------------
# PAGE ACCUEIL
# -------------------------
def accueil():
    st.title("Choisissez une action")

    col1, col2 = st.columns(2)

    with col1:
        if st.button("📄 Convertir PDF en Excel"):
            st.session_state.page = "conversion"

    with col2:
        if st.button("🔍 Comparer deux Excel"):
            st.session_state.page = "comparaison"

# -------------------------
# APP CONVERSION
import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)


col1, col2 = st.columns([1,4])

col1, col2 = st.columns([4, 1])

with col1:
    st.title("Outil de conversion PDF vers Excel")

with col2:
    st.image("logo.png", width=220)

st.set_page_config(
    page_title="Convertisseur PDF vers Excel",
    page_icon="logo.png"
)

st.write(
    "Cette application convertit le PDF en Excel, prêt à être importé dans Odoo."
)


st.subheader("PDF → Excel")

uploaded_file = st.file_uploader("Téléverse ton fichier PDF", type="pdf")


def split_nom_prenom(nom_complet):
    if pd.isna(nom_complet):
        return "", ""

    texte = str(nom_complet).strip()
    if not texte:
        return "", ""

    # Format du PDF: "Nom, Prénom"
    if "," in texte:
        nom_famille, prenom = texte.split(",", 1)
        return nom_famille.strip(), prenom.strip()

    # secours si jamais il n'y a pas de virgule
    morceaux = texte.split()
    if len(morceaux) == 1:
        return morceaux[0], ""
    return morceaux[-1], " ".join(morceaux[:-1])


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
                    # garder seulement les vraies lignes de 6 colonnes
                    if row and len(row) == 6:
                        all_rows.append(row)

    if all_rows:
        df_source = pd.DataFrame(all_rows, columns=headers)

        # nettoyer lignes vides
        df_source = df_source.dropna(how="all").reset_index(drop=True)
        df_source = df_source[
            df_source["Personne"].astype(str).str.strip() != ""
        ].reset_index(drop=True)

        # construire le tableau final
        lignes_finales = []

        for _, row in df_source.iterrows():
            nom_complet = row["Nom, Prénom"]
            nom_famille, prenom = split_nom_prenom(nom_complet)

            lignes_finales.append({
                "Ref.Indiv": row["Personne"],                         # copie Personne
                "Nom": row["Nom, Prénom"],                           # copie Nom, Prénom
                "Nom de famille": nom_famille,                       # découpe
                "Prénom": prenom,                                    # découpe
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
            df_final.to_excel(writer, index=False, sheet_name="Data")

        #st.download_button(
         #   label="Télécharger Excel",
          #  data=output.getvalue(),
           # file_name="conversion_cari.xlsx",
            #mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        #)
        
        nom_fichier = uploaded_file.name
        nom_sans_ext = os.path.splitext(nom_fichier)[0]

        st.download_button(
          label="Télécharger Excel",
          data=output.getvalue(),
          file_name=f"{nom_sans_ext}_Excel.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
       )

    else:
        st.warning("Aucun tableau détecté dans le PDF")
# -------------------------
def app_conversion():
    if st.button("⬅ Retour"):
        st.session_state.page = "accueil"

    st.title("PDF → Excel")

    # 👉 COLLE TON CODE DE CONVERSION ICI
    pdf_file = st.file_uploader("Upload PDF", type=["pdf"])

    if pdf_file:
        st.write("Traitement...")

# -------------------------
# APP COMPARAISON
import streamlit as st
import pandas as pd
import io
import re
import unicodedata
import os
from datetime import datetime

st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

st.set_page_config(page_title="Comparaison Excel - Ref.Indiv", layout="wide")

# 🔹 Normaliser nom de colonne
def normaliser_nom_colonne(col):
    col = str(col).strip().lower()
    col = ''.join(
        c for c in unicodedata.normalize('NFD', col)
        if unicodedata.category(c) != 'Mn'
    )
    col = re.sub(r'[^a-z0-9]', '', col)
    return col


# 🔹 Trouver colonne Ref.Indiv
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


# 🔹 Nettoyage valeurs
def nettoyer_ref(serie):
    return (
        serie.astype(str)
        .str.strip()
        .replace(["nan", "None", "none", "null", ""], pd.NA)
    )


# 🔹 Comparaison
def comparer_fichiers(df1, df2):
    col1 = trouver_colonne_ref_indiv(df1)
    col2 = trouver_colonne_ref_indiv(df2)

    if not col1:
        raise ValueError("Colonne 'Ref.Indiv' introuvable dans le fichier 1")
    if not col2:
        raise ValueError("Colonne 'Ref.Indiv' introuvable dans le fichier 2")

    df1 = df1.copy()
    df2 = df2.copy()

    df1[col1] = nettoyer_ref(df1[col1])
    df2[col2] = nettoyer_ref(df2[col2])

    refs_fichier1 = set(df1[col1].dropna().unique())

    nouvelles_lignes = df2[
        df2[col2].notna() & (~df2[col2].isin(refs_fichier1))
    ].copy()

    return nouvelles_lignes, col1, col2


# 🔹 Export Excel
def dataframe_to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output.getvalue()


# ================== UI ==================
col1, col2 = st.columns([1,4])

col1, col2 = st.columns([4, 1])

with col1:
    st.title("Comparaison de 2 fichiers Excel (Ref.Indiv)")

with col2:
    st.image("logo.png", width=250)

st.set_page_config(
    page_title="Comparaison Excel",
    page_icon="logo.png"
)

st.write(
    "Cette application compare deux fichiers Excel sur la colonne **Ref.Indiv** "
    "et extrait les lignes du **2e fichier** qui n'existent pas dans le **1er**."
)

col1_ui, col2_ui = st.columns(2)

with col1_ui:
    fichier1 = st.file_uploader("Fichier 1 (référence)", type=["xlsx", "xls"])

with col2_ui:
    fichier2 = st.file_uploader("Fichier 2 (à comparer)", type=["xlsx", "xls"])


if fichier1 and fichier2:
    try:
        df1 = pd.read_excel(fichier1)
        df2 = pd.read_excel(fichier2)

        st.success("Fichiers chargés avec succès")

        with st.expander("Aperçu fichier 1"):
            st.dataframe(df1.head(10), use_container_width=True)

        with st.expander("Aperçu fichier 2"):
            st.dataframe(df2.head(10), use_container_width=True)

        if st.button("Lancer la comparaison"):

            nouvelles_lignes, col1, col2 = comparer_fichiers(df1, df2)

            st.info(f"Colonne détectée fichier 1 : {col1}")
            st.info(f"Colonne détectée fichier 2 : {col2}")

            st.subheader("Résultat")

            st.write(f"Nombre de nouvelles lignes : **{len(nouvelles_lignes)}**")

            if len(nouvelles_lignes) > 0:

                st.dataframe(nouvelles_lignes, use_container_width=True)

                # 🔹 Nom du fichier dynamique
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
def app_comparaison():
    if st.button("⬅ Retour"):
        st.session_state.page = "accueil"

    st.title("Comparaison Excel")

    # 👉 COLLE TON CODE DE COMPARAISON ICI
    file1 = st.file_uploader("Fichier 1", type=["xlsx"], key="f1")
    file2 = st.file_uploader("Fichier 2", type=["xlsx"], key="f2")

    if file1 and file2:
        st.write("Comparaison...")

# -------------------------
# ROUTER
# -------------------------
if st.session_state.page == "accueil":
    accueil()

elif st.session_state.page == "conversion":
    app_conversion()

elif st.session_state.page == "comparaison":
    app_comparaison()
