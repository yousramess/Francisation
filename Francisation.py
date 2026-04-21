import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

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
