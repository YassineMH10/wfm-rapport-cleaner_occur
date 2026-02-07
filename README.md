# 📊 WFM Rapport Cleaner (Streamlit)

Application Streamlit pour transformer un fichier brut (.xls/.xlsx) en :
1) **rapport_nettoye.xlsx** (nettoyage + normalisation)
2) **rapport_final_moy_temps_filtre.xlsx** (Moy Temps Total + filtres)

## ✅ Fonctionnalités
- Upload `.xls` / `.xlsx`
- Nettoyage colonnes / lignes vides
- Extraction `Log Téléphonie1` depuis `Nom Agent` (regex "Agent 7014")
- Renommage 2e "Pause" en "Pause générique"
- Conversion `Temps total` type `1h2'3` → `01:02:03`
- Calcul `Moy Temps Total = Temps total / Occurances`
- Filtres paramétrables (Streamlit sidebar)
- Téléchargement des 2 fichiers Excel

## 🚀 Lancer en local
```bash
pip install -r requirements.txt
streamlit run app.py
