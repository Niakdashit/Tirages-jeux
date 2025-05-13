
import streamlit as st
import pandas as pd
import openpyxl
import unicodedata
import re
from io import BytesIO
from zipfile import ZipFile, BadZipFile
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="OPT / GAG Processor", layout="centered")

st.title("📊 Traitement fichiers OPT & GAG")
st.markdown(
    "Nettoyez vos fichiers Excel/TSV pour **FemmeActuelle.fr** ou **CuisineActuelle.fr** "
    "puis téléchargez-les individuellement ou en ZIP."
)

# ---------- Sélections utilisateur ----------
treatment = st.selectbox(
    "Quel traitement souhaitez‑vous appliquer ?",
    ["Opt-in partenaire (OPT)", "Tirages gagnants (GAG)"]
)

brand = st.radio("Choisissez votre marque :", ["FemmeActuelle.fr", "CuisineActuelle.fr"])
brand_prefix = "FA.FR" if brand == "FemmeActuelle.fr" else "CA.FR"

uploaded_files = st.file_uploader(
    "📂 Fichiers à traiter", type=["xls", "xlsx", "tsv"], accept_multiple_files=True
)

download_individual = st.checkbox("📥 Télécharger fichiers un par un", value=True)
download_zip       = st.checkbox("📦 Télécharger tous les fichiers dans une archive ZIP")

# ---------- utilitaires communs ----------
def remove_accents(txt):
    return ''.join(
        c for c in unicodedata.normalize('NFKD', txt) if not unicodedata.combining(c)
    ) if isinstance(txt, str) else txt

def convert_tsv(bytes_in: bytes) -> bytes:
    df = pd.read_csv(BytesIO(bytes_in), sep="\t", encoding="ISO-8859-1", engine="python")
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    out.seek(0)
    return out.read()

# ---------- Traitement OPT ----------
def run_opt(input_bytes: bytes, ref_bytes: bytes) -> bytes:
    df = pd.ExcelFile(BytesIO(input_bytes)).parse(0)

    part_col = next((c for c in df.columns if "partenaire -" in c.lower()), None)
    if not part_col:
        st.error("Colonne 'Partenaire -' non trouvée.")
        return None

    df = df[df[part_col] == True]
    df.columns = [remove_accents(c) for c in df.columns]

    mapping = {
        "Civilite": ["Civilite", "CivilitE", "Civ"],
        "Nom": ["Nom"],
        "Prenom": ["Prenom", "PrEnom"],
        "Adresse": ["Adresse"],
        "Code Postal": ["Code Postal", "CP"],
        "Ville": ["Ville"],
        "Pays": ["Pays"],
        "Email": ["Email", "Mail", "Courriel"],
    }
    rename = {v: k for k, vals in mapping.items() for v in vals if v in df.columns}
    df.rename(columns=rename, inplace=True)
    df = df[list(rename.values())]

    for c in df.columns:
        if c != "Email":
            df[c] = df[c].astype(str).apply(
                lambda x: " ".join(word.capitalize() for word in remove_accents(x).split())
            )

    df.drop_duplicates(subset=["Email"], inplace=True)
    df = df[~df["Ville"].str.lower().str.contains("emerainville", na=False)]

    ref_wb = openpyxl.load_workbook(BytesIO(ref_bytes))
    col_widths = {
        col: dim.width for col, dim in ref_wb.active.column_dimensions.items()
    }

    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Optin", index=False)
        ws = w.book["Optin"]
        header_fill = PatternFill("solid", fgColor="DCE6F1")
        border = Border(*(Side(style="thin"),)*4)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = Font(bold=True)
            cell.border = border
        for col, wth in col_widths.items():
            ws.column_dimensions[col].width = wth or ws.column_dimensions[col].width
    out.seek(0)
    return out.getvalue()

# ---------- Traitement GAG ----------
def run_gag(input_bytes: bytes, ref_bytes: bytes) -> bytes:
    try:
        df = pd.read_excel(BytesIO(input_bytes), dtype=str)
    except BadZipFile:
        df = pd.read_csv(BytesIO(input_bytes), sep="\t", encoding="utf-8", on_bad_lines="skip", dtype=str)

    df = df.shift(axis=1)
    df.rename(columns={
        "Merci de nous transmettre ici votre numéro de téléphone": "Tel"
    }, inplace=True)

    keep = [
        "Civilité", "Nom", "Prénom", "Adresse", "Complément d'adresse",
        "Code Postal", "Ville", "Pays", "Tel", "Email"
    ]
    df = df[[c for c in keep if c in df.columns]]

    df = df.applymap(remove_accents)
    df.columns = [c.capitalize() if c.lower() != "email" else c for c in df.columns]

    if "Tel" in df.columns:
        df["Tel"] = df["Tel"].astype(str).str.replace(".0", "", regex=False).str.zfill(10)
        df["Tel"] = df["Tel"].apply(lambda x: " ".join([x[i:i+2] for i in range(0, len(x), 2)]))

    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Gagnants", index=False)
        w.book.create_sheet("Réservistes")
    out.seek(0)
    return out.getvalue()

# ---------- Chargement fichiers de référence (dans le repo) ----------
with open("ref_opt.xlsx", "rb") as f:
    ref_opt = f.read()
with open("ref_gag.xlsx", "rb") as f:
    ref_gag = f.read()

# ---------- Exécution ----------
if uploaded_files:
    zip_buffer = BytesIO()
    zf = ZipFile(zip_buffer, "w")

    for file in uploaded_files:
        raw_bytes = file.read()
        ext = file.name.split(".")[-1].lower()

        if ext in {"tsv", "xls"}:
            raw_bytes = convert_tsv(raw_bytes)

        # Extraction nom partenaire
        base = re.sub(r"\.(xls|xlsx|tsv)$", "", file.name, flags=re.I)
        match = re.search(r"OPT\s*(.*)", base, re.I)
        partner = match.group(1).strip().upper() if match else "PARTENAIRE"

        out_filename = f"OPT {brand_prefix} - {partner} GJ 2025.xlsx"
        out_bytes = run_opt(raw_bytes, ref_opt) if treatment.startswith("Opt-in") else run_gag(raw_bytes, ref_gag)

        if out_bytes:
            if download_individual:
                st.download_button(f"⬇️ {out_filename}", out_bytes, file_name=out_filename)
            if download_zip:
                zf.writestr(out_filename, out_bytes)

    if download_zip:
        zf.close()
        st.download_button("📦 Télécharger ZIP", zip_buffer.getvalue(),
                           file_name=f"OPT {brand_prefix} - GJ 2025.zip")
