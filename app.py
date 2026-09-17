import os
import tempfile
import zipfile
import streamlit as st
from core.validator import validate_file
from core.excel_cleaner import ExcelCleaner

st.set_page_config(page_title="Excel Import Cleaner", page_icon="📊", layout="centered")

st.title("📊 EXCEL IMPORT CLEANER")
st.caption("Préparez vos fichiers Excel pour un import sans erreurs.")

# Zone de dépôt de fichiers
uploaded_files = st.file_uploader(
    "Glissez-déposez vos fichiers Excel ici (.xlsx, .xlsm)",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True
)

# Options
st.subheader("Options de nettoyage")
mode = st.radio(
    "Mode de conversion :",
    ["Automatique (Dates → texte, Valeurs → texte)", "Dates uniquement → texte JJ/MM/AAAA"]
)
auto_mode = (mode == "Automatique (Dates → texte, Valeurs → texte)")

# Bouton de traitement
if uploaded_files and st.button("PRÉPARER LE(S) FICHIER(S)", type="primary"):
    with st.spinner("Traitement en cours..."):
        with tempfile.TemporaryDirectory() as tmpdirname:
            cleaned_files = []
            
            for uploaded_file in uploaded_files:
                input_path = os.path.join(tmpdirname, uploaded_file.name)
                with open(input_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                valid, msg = validate_file(input_path)
                if not valid:
                    st.error(f"Erreur ({uploaded_file.name}) : {msg}")
                    continue
                
                base, ext = os.path.splitext(uploaded_file.name)
                output_name = f"{base}_PRET_A_IMPORTER{ext}"
                output_path = os.path.join(tmpdirname, output_name)
                
                cleaner = ExcelCleaner(
                    input_path,
                    selected_sheets=None,
                    options={'auto_mode': auto_mode},
                    log_callback=lambda m: None,
                    progress_callback=lambda p, s=None: None
                )
                cleaner.process(output_path)
                cleaned_files.append((output_name, output_path))
            
            st.success("✅ Traitement terminé !")
            
            if len(cleaned_files) == 1:
                name, path = cleaned_files[0]
                with open(path, "rb") as f:
                    st.download_button(
                        label="📥 Télécharger le fichier nettoyé",
                        data=f.read(),
                        file_name=name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            elif len(cleaned_files) > 1:
                zip_path = os.path.join(tmpdirname, "fichiers_nettoyes.zip")
                with zipfile.ZipFile(zip_path, "w") as zipf:
                    for name, path in cleaned_files:
                        zipf.write(path, arcname=name)
                with open(zip_path, "rb") as f:
                    st.download_button(
                        label="📥 Télécharger tous les fichiers (ZIP)",
                        data=f.read(),
                        file_name="fichiers_nettoyes.zip",
                        mime="application/zip"
                    )
