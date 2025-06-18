import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter
import unicodedata
import re
from io import BytesIO
from zipfile import ZipFile
from pathlib import Path

# --- OPT SCRIPT FUNCTIONS ---

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

def convert_tsv_to_xlsx(input_file: BytesIO) -> BytesIO:
    df = pd.read_csv(input_file, sep="\t", encoding="ISO-8859-1", engine="python")
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    out.seek(0)
    return out

def process_opt_stream(df, ref_bytes):
    df.columns = [remove_accents_advanced(col) for col in df.columns]
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
    for correct_name, possible_names in col_mapping.items():
        for possible_name in possible_names:
            if possible_name in df.columns:
                final_mapping[possible_name] = correct_name
    df.rename(columns=final_mapping, inplace=True)
    df = df[list(final_mapping.values())]
    for col in df.columns:
        if col != "Email":
            df[col] = df[col].astype(str).apply(format_name_advanced)
    df.drop_duplicates(subset=["Email"], inplace=True)
    df = df[~df["Ville"].str.lower().str.contains("emerainville|ozoir la ferriere", na=False)]

    ref_wb = openpyxl.load_workbook(BytesIO(ref_bytes))
    ref_ws = ref_wb.active
    column_widths = {col_letter: ref_ws.column_dimensions[col_letter].width for col_letter in ref_ws.column_dimensions}

    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Optin", index=False)
        wb = writer.book
        ws = wb["Optin"]
        header_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = Font(bold=True)
            cell.border = border
        for col_letter, width in column_widths.items():
            if width:
                ws.column_dimensions[col_letter].width = width
        for col in ws.iter_cols():
            if col[0].value == "Code Postal":
                for cell in col[1:]:
                    if isinstance(cell.value, str):
                        cleaned_value = cell.value.replace(" ", "").strip()
                        if cleaned_value.isdigit():
                            cell.value = int(cleaned_value)
                    cell.number_format = '0'
                break
        wb.save(out)
    out.seek(0)
    return out

# --- STREAMLIT INTERFACE ---

st.set_page_config(page_title="OPT & GAG Processor", layout="centered")
st.title("📊 Traitement fichiers OPT / GAG")

treatment = st.selectbox("Type de traitement :", ["Opt-in partenaire (OPT)"])  # GAG à venir
uploaded_files = st.file_uploader("📂 Fichiers (xls, xlsx, tsv)", type=["xls", "xlsx", "tsv"], accept_multiple_files=True)
ref_opt_path = Path("ref_opt.xlsx")
ref_opt_bytes = ref_opt_path.read_bytes() if ref_opt_path.exists() else None

if st.button("🚀 Lancer le traitement") and uploaded_files:
    zip_buf = BytesIO()
    zip_writer = ZipFile(zip_buf, "w")
    for f in uploaded_files:
        raw = f.read()
        ext = f.name.lower().split(".")[-1]
        if ext == "tsv" or ext == "xls":
            raw = convert_tsv_to_xlsx(BytesIO(raw)).read()
        df = pd.ExcelFile(BytesIO(raw)).parse(0)
        if treatment.startswith("Opt"):
            output = process_opt_stream(df, ref_opt_bytes)
            name_part = re.search(r"OPT\s*(.*?)\.", f.name, re.I)
            partenaire = name_part.group(1).strip().upper() if name_part else "PARTENAIRE"
            final_name = f"OPT CA.FR - {partenaire} GJ 2025.xlsx"
            st.download_button(f"⬇️ Télécharger : {final_name}", output, file_name=final_name)
            zip_writer.writestr(final_name, output.read())
    zip_writer.close()
    st.download_button("📦 Télécharger ZIP complet", zip_buf.getvalue(), file_name="Traitements_OPT.zip")
