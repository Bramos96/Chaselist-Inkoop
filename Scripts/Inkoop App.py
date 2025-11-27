import streamlit as st
import os
import subprocess
from datetime import datetime


st.set_page_config(page_title="Inkoop App", layout="centered")

# Map met de .py-scripts
SCRIPTS_FOLDER = r"C:\Users\bram.gerrits\Desktop\Automations\Inkoop\Chaselist Madelon\Scripts"
CHASELIST_SCRIPT = os.path.join(SCRIPTS_FOLDER, "Chaselist.py")

# Map waar Chaselist.py zijn bestanden verwacht (BASE_FOLDER in Chaselist.py)
DATA_FOLDER = r"C:\Users\bram.gerrits\Desktop\Automations\Inkoop\Chaselist Madelon"
CHASE_PREFIX = "Chase"


# ─────────────────────────────
# HULPFUNCTIES
# ─────────────────────────────

def save_uploaded_chase(uploaded_file) -> str:
    """
    Sla het geüploade Chase-bestand op in DATA_FOLDER met een naam
    die met 'Chase' begint, zodat Chaselist.py 'm als nieuwste herkent.
    """
    os.makedirs(DATA_FOLDER, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{CHASE_PREFIX}_from_app_{timestamp}.xlsx"
    dest_path = os.path.join(DATA_FOLDER, filename)

    bytes_data = uploaded_file.read()
    with open(dest_path, "wb") as f:
        f.write(bytes_data)

    return dest_path



def run_chaselist_script() -> subprocess.CompletedProcess:
    """
    Run Chaselist.py exact zoals je dat in een terminal zou doen:
    python Chaselist.py

    Chaselist.py zoekt zelf het nieuwste Chase*-bestand in BASE_FOLDER
    en houdt alle bestaande logica / voorwaarden / opmaak aan.
    """
    result = subprocess.run(
        ["python", CHASELIST_SCRIPT],
        capture_output=True,
        text=True
    )
    return result

# ─────────────────────────────
# UI
# ─────────────────────────────

st.subheader("1️⃣ Upload Chase-bestand")

uploaded = st.file_uploader("Kies een Chase-bestand (.xlsx)", type=["xlsx"])

if uploaded is not None:
    st.success(f"✅ Bestand geselecteerd: {uploaded.name}")

    # Optioneel: even tonen welke naam hij gaat krijgen
    st.caption("Na upload wordt het bestand in dezelfde map als Chaselist.py gezet.")

    st.subheader("2️⃣ Start Chaselist-script")

    if st.button("📤 Mails voorbereiden met Chaselist.py"):
        # 1) Bestand opslaan in BASE_FOLDER
        saved_path = save_uploaded_chase(uploaded)
        st.info(f"Bestand opgeslagen als:\n{saved_path}")

        # 2) Chaselist.py draaien
        with st.spinner("Chaselist.py wordt uitgevoerd… Outlook kan zo mails openen."):
            result = run_chaselist_script()

        # 3) Output laten zien
        st.subheader("🔎 Script-uitvoer (stdout)")
        if result.stdout:
            st.text(result.stdout)
        else:
            st.write("Geen stdout-output.")

        if result.returncode == 0:
            st.success("✅ Chaselist.py is succesvol uitgevoerd. Check je Outlook (mails staan klaar).")
        else:
            st.error("❌ Er is een fout opgetreden in Chaselist.py.")
            st.subheader("⚠️ Foutdetails (stderr)")
            st.text(result.stderr)

else:
    st.info("📁 Upload eerst een Chase-bestand om te beginnen.")
