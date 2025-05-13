
import streamlit as st
import pandas as pd
import openpyxl
import unicodedata, re
from io import BytesIO
from zipfile import ZipFile, BadZipFile
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="OPT & GAG Processor", layout="centered")

st.title("📊 Traitement OPT & GAG")
st.caption("Chargez vos fichiers, cliquez **Lancer le traitement**, puis téléchargez les résultats.")

# ------------------- Sélections -------------------
treatment = st.selectbox(
    "Traitement :", ["Opt-in partenaire (OPT)", "Tirages gagnants (GAG)"]
)

brand = st.radio("Marque :", ["FemmeActuelle.fr", "CuisineActuelle.fr"])
brand_prefix = "FA.FR" if brand == "FemmeActuelle.fr" else "CA.FR"

uploaded_files = st.file_uploader(
    "📂 Fichiers (xls, xlsx, tsv)", type=["xls", "xlsx", "tsv"], accept_multiple_files=True
)

# Options téléchargements (toujours définies AVANT le bouton)
dl_individual = st.checkbox("📥 Téléchargement unitaire", value=True)
dl_zip       = st.checkbox("📦 ZIP global")

run_btn = st.button("🚀 Lancer le traitement")

# ------------------- Helpers -------------------
def remove_acc(txt):
    return ''.join(c for c in unicodedata.normalize('NFKD', txt) if not unicodedata.combining(c)) if isinstance(txt, str) else txt

def convert_tsv(b: bytes) -> bytes:
    df = pd.read_csv(BytesIO(b), sep="\t", encoding="ISO-8859-1", engine="python")
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    buf.seek(0)
    return buf.read()

# Références
ref_opt_bytes = Path("ref_opt.xlsx").read_bytes()
ref_gag_bytes = Path("ref_gag.xlsx").read_bytes()

# ------------------- OPT -------------------
def process_opt(data: bytes, ref: bytes) -> bytes:
    df = pd.ExcelFile(BytesIO(data)).parse(0)
    col_part = next((c for c in df.columns if "partenaire -" in c.lower()), None)
    if not col_part:
        st.error("❌ 'Partenaire -' non trouvé")
        return None
    df = df[df[col_part] == True]

    df.columns = [remove_acc(c) for c in df.columns]

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
    rename = {orig: std for std, alts in mapping.items() for orig in alts if orig in df.columns}
    df.rename(columns=rename, inplace=True)
    df = df[list(rename.values())]

    for c in df.columns:
        if c != "Email":
            df[c] = df[c].astype(str).apply(lambda x: " ".join(word.capitalize() for word in remove_acc(x).split()))

    df.drop_duplicates(subset=["Email"], inplace=True)
    df = df[~df["Ville"].str.lower().str.contains("emerainville", na=False)]

    if "Code Postal" in df.columns:
        df["Code Postal"] = df["Code Postal"].astype(str).str.replace(" ", "").str[:5]
        df["Code Postal"] = pd.to_numeric(df["Code Postal"], errors="coerce").fillna(0).astype(int)

    ref_ws = openpyxl.load_workbook(BytesIO(ref)).active
    ref_widths = {c: dim.width for c, dim in ref_ws.column_dimensions.items()}

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Optin")
        ws = w.book["Optin"]
        header_fill = PatternFill("solid", fgColor="DCE6F1")
        thin = Side(style="thin")
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = Font(bold=True)
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for col, wth in ref_widths.items():
            if wth:
                ws.column_dimensions[col].width = wth

        for col in ws.iter_cols():
            if col[0].value == "Code Postal":
                for cel in col[1:]:
                    cel.number_format = "0"
                break
    buf.seek(0)
    return buf.getvalue()

# ------------------- GAG -------------------
def process_gag(data: bytes, ref: bytes) -> bytes:
    try:
        df = pd.read_excel(BytesIO(data), dtype=str)
    except BadZipFile:
        df = pd.read_csv(BytesIO(data), sep="\t", encoding="utf-8", on_bad_lines="skip", dtype=str)

    df = df.shift(axis=1)
    df.rename(columns={"Merci de nous transmettre ici votre numéro de téléphone": "Tel"}, inplace=True)

    keep = ["Civilité", "Nom", "Prénom", "Adresse", "Complément d'adresse",
            "Code Postal", "Ville", "Pays", "Tel", "Email"]
    df = df[[c for c in keep if c in df.columns]]

    df = df.applymap(remove_acc)
    df.columns = [c.capitalize() if c.lower() != "email" else c for c in df.columns]

    if "Code Postal" in df.columns:
        df["Code Postal"] = df["Code Postal"].astype(str).str.replace(" ", "").str[:5]
        df["Code Postal"] = pd.to_numeric(df["Code Postal"], errors="coerce").fillna(0).astype(int)

    if "Tel" in df.columns:
        df["Tel"] = df["Tel"].astype(str).str.replace(".0", "", regex=False).str.zfill(10)
        df["Tel"] = df["Tel"].apply(lambda x: " ".join([x[i:i+2] for i in range(0, len(x), 2)]))

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Gagnants")
        wb = w.book
        wb.create_sheet("Réservistes")
        ws = wb["Gagnants"]
        header_fill = PatternFill("solid", fgColor="E3F2FD")
        for cell in ws[1]:
            cell.fill = header_fill

        for col in ws.columns:
            length = max(len(str(c.value)) if c.value else 0 for c in col) + 2
            ws.column_dimensions[get_column_letter(col[0].column)].width = length
    buf.seek(0)
    return buf.getvalue()

# ------------------- Exécution -------------------
if run_btn and uploaded_files:
    zip_buffer = BytesIO()
    zip_writer = ZipFile(zip_buffer, "w")

    for f in uploaded_files:
        raw = f.read()
        if f.name.lower().endswith((".tsv", ".xls")):
            raw = convert_tsv(raw)

        base = re.sub(r"\.(xls|xlsx|tsv)$", "", f.name, flags=re.I)
        m = re.search(r"OPT\s*(.*)", base, flags=re.I)
        partner = m.group(1).strip().upper() if m else "PARTENAIRE"

        result = process_opt(raw, ref_opt_bytes) if treatment.startswith("Opt-in") else process_gag(raw, ref_gag_bytes)
        if not result:
            continue

        output_name = f"OPT {brand_prefix} - {partner} GJ 2025.xlsx"
        if dl_individual:
            st.download_button(f"⬇️ {output_name}", result, file_name=output_name)

        if dl_zip:
            zip_writer.writestr(output_name, result)

    if dl_zip:
        zip_writer.close()
        st.download_button(
            "📦 Télécharger ZIP global",
            zip_buffer.getvalue(),
            file_name=f"OPT {brand_prefix} - GJ 2025.zip"
        )
