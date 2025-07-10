import streamlit as st
import pandas as pd
from fpdf import FPDF
from io import BytesIO

# Optionnel : si tu veux utiliser SharePoint
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

# -----------------------
# 📄 Fonction PDF
# -----------------------
def generer_pdf_intervention(row):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, f"Intervention {row['InterventionNumber']}", ln=True, align="C")
    pdf.ln(10)

    pdf.set_font("Arial", size=10)
    lignes = [
        ("Date", row.get("InterventionDate", "")),
        ("Adresse", f"{row.get('Street', '')} {row.get('Number', '')}, {row.get('Zip', '')} {row.get('City', '')}"),
        ("Description", row.get("Description", ""))
    ]

    for titre, valeur in lignes:
        pdf.cell(50, 10, f"{titre} :", ln=0)
        pdf.multi_cell(0, 10, str(valeur))

    pdf_bytes = pdf.output(dest='S').encode('latin1')
    return BytesIO(pdf_bytes)

# -----------------------
# ⚙️ Configuration Streamlit
# -----------------------
st.set_page_config(page_title="Interventions", layout="wide")
st.title("📅 Rechercher des interventions")

# -----------------------
# 🛜 Télécharger depuis SharePoint si activé
# -----------------------
USE_SHAREPOINT = True  # Passe à False pour tester en local

def telecharger_excel_sharepoint(site_url, client_id, client_secret, fichier_sharepoint):
    ctx = ClientContext(site_url).with_credentials(ClientCredential(client_id, client_secret))
    fichier = ctx.web.get_file_by_server_relative_url(fichier_sharepoint)
    buffer = BytesIO()
    fichier.download(buffer).execute_query()
    buffer.seek(0)
    return buffer

# -----------------------
# 📁 Charger les données
# -----------------------
@st.cache_data
def charger_donnees(fichier):
    df = pd.read_excel(fichier, sheet_name="POWERBI V_Intervention")
    df["InterventionDate"] = pd.to_datetime(df["InterventionDate"], errors="coerce")
    df = df.dropna(subset=["InterventionDate"])
    return df

# Infos SharePoint (à adapter)
if USE_SHAREPOINT:
    site_url = st.secrets["sharepoint"]["site_url"]
    client_id = st.secrets["sharepoint"]["client_id"]
    client_secret = st.secrets["sharepoint"]["client_secret"]
    fichier_sharepoint = st.secrets["sharepoint"]["fichier_sharepoint"]
    fichier_excel = telecharger_excel_sharepoint(site_url, client_id, client_secret, fichier_sharepoint)
else:
    fichier_excel = r"C:\Users\Arnaud Mathieu\OneDrive - incendiebw.be\arnaud\4AMU2\PAU\Stat NEW\DB rapport 2324.xlsx"

df = charger_donnees(fichier_excel)

# -----------------------
# 📅 Sélection de dates
# -----------------------
col1, col2 = st.columns(2)
with col1:
    date_debut = st.date_input("Date de début", value=df["InterventionDate"].min().date())
with col2:
    date_fin = st.date_input("Date de fin", value=df["InterventionDate"].max().date())

# -----------------------
# 🏙️ Champ libre pour la ville
# -----------------------
ville_recherche = st.text_input("🔍 Rechercher une commune (partie du nom)", "")

# -----------------------
# 🔍 Lancer la recherche
# -----------------------
if st.button("🔍 Valider et rechercher"):
    masque = (
        (df["InterventionDate"].dt.date >= date_debut) &
        (df["InterventionDate"].dt.date <= date_fin) &
        (df["Urgency"] != "Ambulances")
    )

    if ville_recherche.strip():
        masque = masque & (df["City"].str.contains(ville_recherche, case=False, na=False))

    resultat = df.loc[masque, [
        "Id", "InterventionNumber", "InterventionDate", "Street", "Number", "Zip", "City", "Description",
        "Responsible StationName",
    ]]

    st.success(f"{len(resultat)} intervention(s) trouvée(s)")

    for _, row in resultat.iterrows():
        col1, col2 = st.columns([5, 1])
        with col1:
            st.markdown(
                f"**🧯 Intervention {row['InterventionNumber']}**  \n"
                f"📍 {row['Street']} {row['Number']}, {row['Zip']} {row['City']}  \n"
                f"📋 {row['Description']}  \n"
                f"📅 {row['InterventionDate'].date()}  \n"
                f"🚒 {row['Responsible StationName']}"
            )
        with col2:
            pdf_buffer = generer_pdf_intervention(row)
            st.download_button(
                label="📄 PDF",
                data=pdf_buffer,
                file_name=f"Intervention_{row['InterventionNumber']}.pdf",
                mime="application/pdf",
                key=f"btn_{row['Id']}"
            )
        st.markdown("---")
else:
    st.info("Choisissez une période et éventuellement une commune (partielle), puis cliquez sur **Valider et rechercher**.")