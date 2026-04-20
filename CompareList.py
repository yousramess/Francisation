import streamlit as st
import pandas as pd
import io
import re
import unicodedata
import os
from datetime import datetime



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
