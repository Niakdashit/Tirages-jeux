
import streamlit as st

st.set_page_config(page_title="Débogage App", layout="centered")
st.title("🔍 Débogage interface Streamlit")

try:
    st.success("✅ L'application démarre correctement.")
    traitement = st.selectbox("Quel traitement souhaitez-vous tester ?", ["OPT", "GAG"])
    marque = st.radio("Marque :", ["FemmeActuelle.fr", "CuisineActuelle.fr"])
    nb_gagnants = st.number_input("Nombre de gagnants (GAG uniquement)", min_value=1, value=10) if traitement == "GAG" else None
    fichiers = st.file_uploader("Déposez vos fichiers (xls, xlsx, tsv)", type=["xls", "xlsx", "tsv"], accept_multiple_files=True)
    if st.button("🚀 Lancer"):
        st.info(f"Traitement = {traitement}")
        st.info(f"Marque = {marque}")
        if traitement == "GAG":
            st.info(f"Nb gagnants = {nb_gagnants}")
        if fichiers:
            st.success(f"{len(fichiers)} fichier(s) prêt(s) à être traité(s).")
        else:
            st.warning("Aucun fichier uploadé.")
except Exception as e:
    st.error(f"❌ Erreur critique au chargement : {e}")
