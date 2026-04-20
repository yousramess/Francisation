import streamlit as st
import pandas as pd
import io
import re
import unicodedata

st.set_page_config(page_title="Comparaison Excel - Réf.Ind", layout="wide")


def normaliser_nom_colonne(col):
    """Normalise le nom d'une colonne pour détecter Num.Réf malgré les variantes."""
    col = str(col).strip().lower()
    col = ''.join(
        c for c in unicodedata.normalize('NFD', col)
        if unicodedata.category(c) != 'Mn'
    )
    col = re.sub(r'[^a-z0-9]', '', col)
    return col


def trouver_colonne_num_ref(df):
    """Trouve automatiquement la colonne Ref.Indiv"""
    variantes_possibles = {
        "Ref.Indiv",
        "Referennce Indiv",
    }

    for col in df.columns:
        if normaliser_nom_colonne(col) in variantes_possibles:
            return col
    return None


def nettoyer_num_ref(serie):
    """Nettoie les valeurs Ref.Indiv pour la comparaison."""
    return (
        serie.astype(str)
        .str.strip()
        .replace(["nan", "None", "none", "null", ""], pd.NA)
    )


def comparer_fichiers(df1, df2):
    col1 = trouver_colonne_num_ref(df1)
    col2 = trouver_colonne_num_ref(df2)

    if not col1:
        raise ValueError("La colonne 'Ref.Indiv' est introuvable dans le fichier 1.")
    if not col2:
        raise ValueError("La colonne 'Ref.Indiv' est introuvable dans le fichier 2.")

    df1 = df1.copy()
    df2 = df2.copy()

    df1[col1] = nettoyer_num_ref(df1[col1])
    df2[col2] = nettoyer_num_ref(df2[col2])

    refs_fichier1 = set(df1[col1].dropna().unique())

    nouvelles_lignes = df2[
        df2[col2].notna() & (~df2[col2].isin(refs_fichier1))
    ].copy()

    return nouvelles_lignes, col1, col2


def dataframe_to_excel_bytes(df):
    """Convertit un DataFrame en fichier Excel téléchargeable."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Nouvelles_lignes")
    output.seek(0)
    return output.getvalue()


# Interface
st.title("Comparaison de 2 fichiers Excel par Référence Indiv")
st.write(
    "Téléverse 2 fichiers Excel. L'application va comparer la colonne **Ref.Indiv** "
    "et extraire les lignes du **2e fichier** dont le numéro n'existe pas dans le **1er**."
)

col_a, col_b = st.columns(2)

with col_a:
    fichier1 = st.file_uploader("Téléverser le fichier 1 (référence)", type=["xlsx", "xls"], key="f1")

with col_b:
    fichier2 = st.file_uploader("Téléverser le fichier 2 (à vérifier)", type=["xlsx", "xls"], key="f2")

if fichier1 and fichier2:
    try:
        df1 = pd.read_excel(fichier1)
        df2 = pd.read_excel(fichier2)

        st.success("Les 2 fichiers ont été chargés avec succès.")

        with st.expander("Aperçu du fichier 1"):
            st.dataframe(df1.head(10), use_container_width=True)

        with st.expander("Aperçu du fichier 2"):
            st.dataframe(df2.head(10), use_container_width=True)

        if st.button("Comparer les fichiers"):
            nouvelles_lignes, col1, col2 = comparer_fichiers(df1, df2)

            st.info(f"Colonne détectée dans fichier 1 : **{col1}**")
            st.info(f"Colonne détectée dans fichier 2 : **{col2}**")

            st.subheader("Résultat de la comparaison")
            st.write(f"Nombre de nouvelles lignes trouvées : **{len(nouvelles_lignes)}**")

            if len(nouvelles_lignes) > 0:
                st.dataframe(nouvelles_lignes, use_container_width=True)

                excel_bytes = dataframe_to_excel_bytes(nouvelles_lignes)
                
                nom_original = os.path.splitext(fichier2.name)[0]  # enlève .xlsx
                nom_sortie = f"{nom_original}_New.xlsx"

                st.download_button(
                    label="Télécharger le fichier résultat",
                    data=excel_bytes,
                    file_name=nom_sortie,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            else:
                st.warning("Aucune nouvelle ligne à ajouter. Tous les Ref.Indiv du fichier 2 existent déjà dans le fichier 1.")

    except Exception as e:
        st.error(f"Erreur : {e}")
else:
    st.warning("Veuillez téléverser les 2 fichiers Excel.")

