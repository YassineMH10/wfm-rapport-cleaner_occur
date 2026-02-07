import re
import io
import numpy as np
import pandas as pd


DEFAULT_ETATS_AUTORISES = [
    "Aucun contexte démarré", "Back Office", "BUG IT", "Break", "Détachement", "Mailing",
    "Meeting", "Numérotation", "OJT", "Pause générique", "Rappel", "Training"
]


def read_excel_any(file_bytes: bytes, filename: str) -> pd.DataFrame:
    ext = filename.split(".")[-1].lower()
    bio = io.BytesIO(file_bytes)
    if ext == "xls":
        return pd.read_excel(bio, engine="xlrd")
    return pd.read_excel(bio, engine="openpyxl")


def to_hms(val):
    """Convertit formats type 1h2'3 en 01:02:03. Sinon retourne tel quel."""
    if isinstance(val, str):
        m = re.match(r"^\s*(\d+)h(\d+)'(\d+)\s*$", val)
        if m:
            h, mi, s = m.groups()
            return f"{int(h):02}:{int(mi):02}:{int(s):02}"
    return val


def hhmmss_to_seconds(timestr):
    try:
        h, m, s = map(int, str(timestr).split(":"))
        return h * 3600 + m * 60 + s
    except:
        return np.nan


def seconds_to_hhmmss(seconds):
    if pd.isna(seconds):
        return ""
    seconds = int(round(float(seconds)))
    h = seconds // 3600
    m = (seconds % 3600) // 60
    s = seconds % 60
    return f"{h:02d}:{m:02d}:{s:02d}"


def rename_second_pause(df: pd.DataFrame) -> pd.DataFrame:
    """Renomme la 2e occurrence de 'Pause' (et suivantes) en 'Pause générique' par agent."""
    def _apply(group):
        pauses = group[group["Etat"] == "Pause"].index
        if len(pauses) > 1:
            group.loc[pauses[1:], "Etat"] = "Pause générique"
        return group
    return df.groupby("Log Téléphonie1", group_keys=False).apply(_apply)


def clean_stage_1(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    CODE 1 : Nettoyage brut -> fichier 1 (rapport_nettoye.xlsx)
    """
    df = df_raw.copy()

    # Supprimer lignes/colonnes vides
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")

    if "Unnamed: 0" not in df.columns:
        raise ValueError("Colonne attendue 'Unnamed: 0' introuvable (structure du fichier brut différente).")

    df["Unnamed: 0"] = df["Unnamed: 0"].ffill()

    # Renommer colonnes (si présentes)
    rename_map = {
        "Unnamed: 0": "Nom Agent",
        "Unnamed: 1": "Etat",
        "Unnamed: 4": "Occurances",
        "Unnamed: 6": "Temps total"
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    needed = ["Nom Agent", "Etat", "Occurances", "Temps total"]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise ValueError(f"Colonnes manquantes après renommage: {missing}")

    # Drop colonnes inutiles si existent
    for col in ["Unnamed: 3", "Unnamed: 5"]:
        if col in df.columns:
            df = df.drop(columns=[col])

    # Drop colonnes Unnamed >=7
    cols_drop = []
    for c in df.columns:
        if isinstance(c, str) and c.startswith("Unnamed:"):
            try:
                idx = int(c.split(":")[1])
                if idx >= 7:
                    cols_drop.append(c)
            except:
                pass
    if cols_drop:
        df = df.drop(columns=cols_drop)

    # Filtrer Etat non vide
    df = df[df["Etat"].notna()].copy()

    # Log Téléphonie1
    df.insert(0, "Log Téléphonie1", df["Nom Agent"].astype(str).str.extract(r"Agent\s+(\d{4})")[0])

    # Pause -> Pause générique (2e pause)
    df = rename_second_pause(df)

    # Temps total -> hh:mm:ss
    df["Temps total"] = df["Temps total"].apply(to_hms)

    # Normalisation Etat
    df["Etat"] = df["Etat"].replace({
        "Attente": "Attente global",
        "Pause": "Pause global",
        "Preview": "Histo Mailing"
    })

    # Exclure "en attente"
    df = df[df["Etat"].astype(str).str.lower() != "en attente"].copy()

    # Occurances numérique
    df["Occurances"] = pd.to_numeric(df["Occurances"], errors="coerce")

    return df


def clean_stage_2(
    df_stage1: pd.DataFrame,
    etats_autorises=None,
    min_occurrences: int = 3,
    max_moy_seconds: int = 120
) -> pd.DataFrame:
    """
    CODE 2 : fichier 1 -> fichier final (moy temps + filtres)
    """
    if etats_autorises is None:
        etats_autorises = DEFAULT_ETATS_AUTORISES

    df = df_stage1.copy()

    # conversions
    df["Temps total (sec)"] = df["Temps total"].apply(hhmmss_to_seconds)
    df["Occurances"] = pd.to_numeric(df["Occurances"], errors="coerce")

    # éviter div/0
    df = df[df["Occurances"].fillna(0) > 0].copy()

    df["Moy Temps Total (sec)"] = df["Temps total (sec)"] / df["Occurances"]
    df["Moy Temps Total"] = df["Moy Temps Total (sec)"].apply(seconds_to_hhmmss)

    # filtre temps
    df = df[df["Moy Temps Total (sec)"].fillna(10**9) <= max_moy_seconds].copy()

    # filtre Etat + occurences
    df = df[
        (df["Etat"].isin(etats_autorises)) &
        (df["Occurances"] >= min_occurrences)
    ].copy()

    # drop techniques
    df.drop(columns=["Temps total (sec)", "Moy Temps Total (sec)"], inplace=True, errors="ignore")

    return df


def to_excel_bytes(df: pd.DataFrame, sheet_name="data") -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        # Mise en forme simple pro (largeurs)
        ws = writer.sheets[sheet_name]
        for i, col in enumerate(df.columns):
            # largeur = max(len(col), moyenne d'échantillon) bornée
            sample = df[col].astype(str).head(200)
            width = max(len(col), int(sample.map(len).mean() if len(sample) else len(col))) + 4
            width = min(max(width, 12), 45)
            ws.set_column(i, i, width)
        ws.freeze_panes(1, 0)
    return out.getvalue()
