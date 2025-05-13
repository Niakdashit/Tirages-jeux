
import streamlit as st
import pandas as pd
import openpyxl
import unicodedata
import re
from io import BytesIO
from zipfile import ZipFile, BadZipFile
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="OPT / GAG Processor", layout="centered")

st.title("📊 Traitement OPT & GAG")
st.markdown(
    "Déposez vos fichiers, choisissez le traitement puis cliquez sur **Lancer** "
    "pour générer vos fichiers normalisés."
)

# ---------- Sélections ----------
treatment = st.selectbox(
    "Quel traitement souhaitez‑vous appliquer ?",
    ["Opt-in partenaire (OPT)", "Tirages gagnants (GAG)"]
)

brand = st.radio("Choisissez votre marque :", ["FemmeActuelle.fr", "CuisineActuelle.fr"])
brand_prefix = "FA.FR" if brand == "FemmeActuelle.fr" else "CA.FR"

uploaded_files = st.file_uploader(
    "📂 Fichiers à traiter (XLS, XLSX, TSV)", type=["xls", "xlsx", "tsv"], accept_multiple_files=True
)

launch = st.button("🚀 Lancer le traitement")

# ---------- Utils ----------
def remove_accents(txt):
    return ''.join(
        c for c in unicodedata.normalize('NFKD', txt) if not unicodedata.combining(c)
    ) if isinstance(txt, str) else txt

def convert_tsv(tsv_bytes: bytes) -> bytes:
    df = pd.read_csv(BytesIO(tsv_bytes), sep="\t", encoding="ISO-8859-1", engine="python")
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    buf.seek(0)
    return buf.read()

# ---------- Charger références ----------
with open("ref_opt.xlsx", "rb") as f:
    ref_opt_bytes = f.read()
with open("ref_gag.xlsx", "rb") as f:
    ref_gag_bytes = f.read()

# ---------- Fonction OPT ----------
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

    for col in df.columns:
        if col != "Email":
            df[col] = df[col].astype(str).apply(
                lambda x: " ".join(word.capitalize() for word in remove_accents(x).split())
            )

    df.drop_duplicates(subset=["Email"], inplace=True)
    df = df[~df["Ville"].str.lower().str.contains("emerainville", na=False)]

    # Conversion du Code Postal en nombre (sans décimales)
    if "Code Postal" in df.columns:
        df["Code Postal"] = df["Code Postal"].astype(str).str.replace(" ", "").str[:5]
        df["Code Postal"] = pd.to_numeric(df["Code Postal"], errors="coerce").fillna(0).astype(int)

    ref_wb = openpyxl.load_workbook(BytesIO(ref_bytes))
    ref_ws = ref_wb.active
    col_widths = {col: dim.width for col, dim in ref_ws.column_dimensions.items()}

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Optin")
        wb = w.book
        ws = wb["Optin"]

        header_fill = PatternFill("solid", fgColor="DCE6F1")
        thin = Side(style="thin")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = Font(bold=True)
            cell.border = border

        # Largeurs de colonne
        for letter, width in col_widths.items():
            if width:
                ws.column_dimensions[letter].width = width

        # Format code postal
        for col in ws.iter_cols():
            if col[0].value == "Code Postal":
                for c in col[1:]:
                    c.number_format = "0"
                break
    buf.seek(0)
    return buf.getvalue()

# ---------- Fonction GAG ----------
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

    # Code postal en nombre
    if "Code Postal" in df.columns:
        df["Code Postal"] = df["Code Postal"].astype(str).str.replace(" ", "").str[:5]
        df["Code Postal"] = pd.to_numeric(df["Code Postal"], errors="coerce").fillna(0).astype(int)

    if "Tel" in df.columns:
        df["Tel"] = df["Tel"].astype(str).str.replace(".0", "", regex=False).str.zfill(10)
        df["Tel"] = df["Tel"].apply(lambda x: " ".join([x[i:i+2] for i in range(0, len(x), 2)]))

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Gagnants", index=False)
        wb = w.book
        ws = wb["Gagnants"]
        header_fill = PatternFill("solid", fgColor="E3F2FD")
        for cell in ws[1]:
            cell.fill = header_fill

        wb.create_sheet("Réservistes")

        # Largeurs auto
        for col in ws.columns:
            length = max(len(str(c.value)) if c.value else 0 for c in col) + 2
            ws.column_dimensions[get_column_letter(col[0].column)].width = length
    buf.seek(0)
    return buf.getvalue()

# ---------- Lancement quand bouton cliqué ----------
if launch and uploaded_files:
    zip_buf = BytesIO()
    zip_writer = ZipFile(zip_buf, "w")

    for up_file in uploaded_files:
        raw = up_file.read()
        ext = up_file.name.split(".")[-1].lower()
        if ext in {"tsv", "xls"}:
            raw = convert_tsv(raw)

        base_name = re.sub(r"\.(xls|xlsx|tsv)$", "", up_file.name, flags=re.I)
        match = re.search(r"OPT\s*(.*)", base_name, re.I)
        partner = (match.group(1).strip().upper() if match else "PARTENAIRE")

        final_name = f"OPT {brand_prefix} - {partner} GJ 2025.xlsx"

        out_bytes = run_opt(raw, ref_opt_bytes) if treatment.startswith("Opt-in") else run_gag(raw, ref_gag_bytes)
        if not out_bytes:
            continue

        if download_individual:
            st.download_button(f"⬇️ {final_name}", out_bytes, file_name=final_name)

        if download_zip:
            zip_writer.writestr(final_name, out_bytes)

    if download_zip:
        zip_writer.close()
        st.download_button(
            "📦 Télécharger ZIP",
            zip_buf.getvalue(),
            file_name=f"OPT {brand_prefix} - GJ 2025.zip"
        )

