import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO


st.set_page_config(page_title="Francisation PDF → Excel", layout="wide")
st.title("PDF → Excel Converter")

uploaded_file = st.file_uploader("Upload ton PDF", type="pdf")


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
                "Cellulaire": row["Céllulaire ou Téléphone autre"],
                "Téléphone autre": row["Téléphone maison"],
      
                "Francisation": "VRAI",
                "CARI Usager": "VRAI",
                "ÉTIQUETTE": ""

                
            })

        df_final = pd.DataFrame(lignes_finales)

        st.success("Conversion réussie ✅")
        st.dataframe(df_final, use_container_width=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_final.to_excel(writer, index=False, sheet_name="Data")

        st.download_button(
            label="Télécharger Excel",
            data=output.getvalue(),
            file_name="conversion_cari.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.warning("Aucun tableau détecté dans le PDF")