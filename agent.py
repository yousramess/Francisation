import pdfplumber, pandas as pd

with pdfplumber.open("test.pdf") as pdf:
    for page in pdf.pages:
        tables = page.extract_tables()
        # tables[0] = premier tableau de la page

df = pd.DataFrame(tables[0][1:], columns=tables[0][0])
# Séparer nom / prénom
df["Prénom"] = df["Nom,Prénom"].str.split().str[0]
df["Nom"]    = df["Nom,Prénom"].str.split().str[-1]
df.to_excel("sortieTest.xlsx", index=False)