
import streamlit as st
import pandas as pd
import openpyxl
import unicodedata
import re
from openpyxl.styles import Font, PatternFill, Border, Side
from io import BytesIO

st.set_page_config(page_title="OPT Cleaner", layout="centered")

# Custom CSS for design
st.markdown("""
<style>
    html, body {
        background-color: #f9fafb;
    }
    .main {
        background-color: white;
        padding: 2rem;
        border-radius: 1rem;
        box-shadow: 0 0 20px rgba(0,0,0,0.05);
    }
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    h1 {
        color: #841b60;
    }
    .stButton>button {
        background-color: #841b60;
        color: white;
        font-weight: bold;
        border-radius: 0.5rem;
        padding: 0.5rem 1rem;
        border: none;
    }
    .stButton>button:hover {
        background-color: #6d1750;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 OPT Excel Cleaner")
st.markdown("Déposez vos fichiers OPT (XLS/XLSX/TSV) et un fichier de référence. Chaque fichier sera traité et formaté automatiquement.")

# Fonctions de nettoyage
def remove_accents_advanced(text):
    if isinstance(text, str):
        text = unicodedata.normalize('NFKD', text)
        text = re.sub(r'[\u0300-\u036f]', '', text)
        text = text.replace("Œ", "OE").replace("œ", "oe")
        return text.strip()
    return text

def format_name_advanced(value):
    if isinstance(value, str):
        value = remove_accents_advanced(value)
        return " ".join(word.capitalize() for word in value.split())
    return value

def convert_tsv_to_xlsx(input_bytes):
    df = pd.read_csv(BytesIO(input_bytes), sep="\t", encoding="ISO-8859-1", engine="python")
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    output.seek(0)
    return output

def process_opt_file(input_bytes, reference_bytes):
    input_xl = pd.ExcelFile(BytesIO(input_bytes))
    df_opt = input_xl.parse(input_xl.sheet_names[0])
    partenaire_col = next((col for col in df_opt.columns if "partenaire -" in col.lower()), None)
    if not partenaire_col:
        st.error("Aucune colonne contenant 'Partenaire - ' trouvée.")
        return None

    df_opt = df_opt[df_opt[partenaire_col] == True]
    df_opt.columns = [remove_accents_advanced(col) for col in df_opt.columns]
    col_mapping = {
        "Civilite": ["Civilite", "CivilitE", "Civ"],
        "Nom": ["Nom"],
        "Prenom": ["Prenom", "PrEnom"],
        "Adresse": ["Adresse"],
        "Code Postal": ["Code Postal", "CP"],
        "Ville": ["Ville"],
        "Pays": ["Pays"],
        "Email": ["Email", "Mail", "Courriel"]
    }
    final_mapping = {}
    for correct, possibles in col_mapping.items():
        for name in possibles:
            if name in df_opt.columns:
                final_mapping[name] = correct
    df_opt.rename(columns=final_mapping, inplace=True)
    df_opt = df_opt[list(final_mapping.values())]
    for col in df_opt.columns:
        if col != "Email":
            df_opt[col] = df_opt[col].astype(str).apply(format_name_advanced)
    df_opt.drop_duplicates(subset=["Email"], inplace=True)
    df_opt = df_opt[~df_opt["Ville"].str.lower().str.contains("emerainville", na=False)]

    ref_wb = openpyxl.load_workbook(BytesIO(reference_bytes))
    ref_ws = ref_wb.active
    col_widths = {col: ref_ws.column_dimensions[col].width for col in ref_ws.column_dimensions}

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_opt.to_excel(writer, sheet_name="Optin", index=False)
        wb = writer.book
        ws = wb["Optin"]
        fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        border = Border(left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))
        for cell in ws[1]:
            cell.fill = fill
            cell.font = Font(bold=True)
            cell.border = border
        for col_letter, width in col_widths.items():
            if width:
                ws.column_dimensions[col_letter].width = width
        for col in ws.iter_cols():
            if col[0].value == "Code Postal":
                for cell in col[1:]:
                    if isinstance(cell.value, str):
                        val = cell.value.replace(" ", "").strip()
                        if val.isdigit():
                            cell.value = int(val)
                    cell.number_format = '0'
                break
        wb.save(output)
    output.seek(0)
    return output

# Interface
uploaded_files = st.file_uploader("📂 Fichiers OPT (xls, xlsx, tsv)", type=["xls", "xlsx", "tsv"], accept_multiple_files=True)
ref_file = st.file_uploader("📁 Fichier de référence (.xlsx)", type="xlsx")

if uploaded_files and ref_file:
    for file in uploaded_files:
        ext = file.name.split('.')[-1].lower()
        if ext in ['tsv', 'xls']:
            converted = convert_tsv_to_xlsx(file.read())
            result = process_opt_file(converted.read(), ref_file.read())
        else:
            result = process_opt_file(file.read(), ref_file.read())
        if result:
            st.success(f"✅ Fichier traité : {file.name}")
            st.download_button(f"⬇️ Télécharger : output_{file.name}.xlsx", result, file_name=f"output_{file.name}.xlsx")
