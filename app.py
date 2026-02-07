import streamlit as st
import pandas as pd

from processing import (
    read_excel_any,
    clean_stage_1,
    clean_stage_2,
    to_excel_bytes,
    DEFAULT_ETATS_AUTORISES
)

st.set_page_config(
    page_title="WFM Rapport Cleaner",
    page_icon="📊",
    layout="wide"
)

# ---------- HEADER ----------
st.markdown(
    """
    <div style="padding: 14px 16px; border-radius: 14px; background: #111827;">
      <div style="font-size: 22px; font-weight: 700; color: white;">📊 WFM Rapport Cleaner</div>
      <div style="color: #D1D5DB; margin-top: 4px;">
        Pipeline complet : <b>Brut → Nettoyage (Fichier 1) → Moy Temps Total + Filtres (Fichier final)</b>
      </div>
    </div>
    """,
    unsafe_allow_html=True
)

st.write("")

# ---------- SIDEBAR (PARAMS) ----------
with st.sidebar:
    st.header("⚙️ Paramètres")
    st.caption("Ajuste les filtres du fichier final.")

    max_moy_seconds = st.slider("Max Moy Temps Total (secondes)", 5, 600, 120, 5)
    min_occ = st.number_input("Min Occurances", min_value=1, max_value=999, value=3, step=1)

    st.write("---")
    st.subheader("États autorisés")
    etats = st.multiselect(
        "Choisis les états",
        options=DEFAULT_ETATS_AUTORISES,
        default=DEFAULT_ETATS_AUTORISES
    )

    st.write("---")
    show_stage1 = st.checkbox("Afficher le fichier 1 (nettoyé)", value=True)
    preview_rows = st.slider("Lignes d’aperçu", 5, 200, 50, 5)

# ---------- MAIN UPLOAD ----------
col1, col2 = st.columns([1.4, 1])

with col1:
    uploaded = st.file_uploader(
        "📥 Uploade ton fichier brut (.xls ou .xlsx)",
        type=["xls", "xlsx"]
    )
    st.caption("Astuce : si le fichier brut change de structure (colonnes Unnamed), on ajustera le mapping.")

with col2:
    st.info(
        "✅ Sorties générées :\n"
        "- **rapport_nettoye.xlsx** (Fichier 1)\n"
        "- **rapport_final_moy_temps_filtre.xlsx** (Fichier final)",
        icon="ℹ️"
    )

if not uploaded:
    st.stop()

# ---------- PROCESS ----------
try:
    file_bytes = uploaded.getvalue()
    df_raw = read_excel_any(file_bytes, uploaded.name)

    df_stage1 = clean_stage_1(df_raw)
    df_final = clean_stage_2(
        df_stage1,
        etats_autorises=etats,
        min_occurrences=int(min_occ),
        max_moy_seconds=int(max_moy_seconds)
    )

except Exception as e:
    st.error(f"Erreur pendant le traitement : {e}")
    st.stop()

# ---------- KPI ----------
k1, k2, k3, k4 = st.columns(4)
k1.metric("Lignes (Brut)", f"{len(df_raw):,}".replace(",", " "))
k2.metric("Lignes (Fichier 1)", f"{len(df_stage1):,}".replace(",", " "))
k3.metric("Lignes (Final)", f"{len(df_final):,}".replace(",", " "))
k4.metric("États sélectionnés", str(len(etats)))

st.write("")

# ---------- DOWNLOADS ----------
cA, cB = st.columns([1, 1])

with cA:
    st.subheader("⬇️ Téléchargements")
    bytes_stage1 = to_excel_bytes(df_stage1, sheet_name="rapport_nettoye")
    st.download_button(
        label="Télécharger rapport_nettoye.xlsx (Fichier 1)",
        data=bytes_stage1,
        file_name="rapport_nettoye.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    bytes_final = to_excel_bytes(df_final, sheet_name="rapport_final")
    st.download_button(
        label="Télécharger rapport_final_moy_temps_filtre.xlsx (Fichier final)",
        data=bytes_final,
        file_name="rapport_final_moy_temps_filtre.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with cB:
    st.subheader("🧾 Règles appliquées (Final)")
    st.write(
        f"""
- Moy Temps Total ≤ **{max_moy_seconds} sec**
- Occurances ≥ **{min_occ}**
- États autorisés : **{len(etats)}**
        """
    )

st.write("")

# ---------- PREVIEW TABLES ----------
tab1, tab2 = st.tabs(["📄 Aperçu Final", "🧼 Aperçu Fichier 1"])

with tab1:
    st.dataframe(df_final.head(preview_rows), use_container_width=True)
    st.caption("Aperçu du fichier final filtré.")

with tab2:
    if show_stage1:
        st.dataframe(df_stage1.head(preview_rows), use_container_width=True)
        st.caption("Aperçu du fichier 1 (nettoyé).")
    else:
        st.info("Option désactivée dans la sidebar.")
