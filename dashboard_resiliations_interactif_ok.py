#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np
import re, math
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from pathlib import Path
st.set_page_config(page_title="Résiliations — Dashboard", layout="wide")

# ----------------- CHEMINS -----------------
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DATA_PATH = DATA_DIR / "DEM_VOLO_M2.xlsx"
SHEET = "Sheet1"

# ----------------- AIDES -----------------
LABELS_MOIS_FR   = ["Jan","Fév","Mar","Avr","Mai","Juin","Juil","Aoû","Sept","Oct","Nov","Déc"]
LABELS_MOIS_ASCII= ["Jan","Fev","Mars","Avr","Mai","Juin","Juil","Aout","Sept","Oct","Nov","Dec"]

def month_from_long(val):
    if pd.isna(val): return np.nan
    s = str(val).strip()
    m = re.match(r"^(\d{1,2})", s)
    if m: return int(m.group(1))
    base = s.lower()[:3]
    mapping = {"jan":1,"fév":2,"fev":2,"mar":3,"avr":4,"mai":5,"jun":6,"jui":7,"aoû":8,"aou":8,"sep":9,"oct":10,"nov":11,"déc":12,"dec":12}
    return mapping.get(base, np.nan)

# ----------------- DONNÉES -----------------
@st.cache_data(show_spinner=False)
def load_data(path: Path, sheet: str):
    """
    Charge le fichier Excel. En ligne (Streamlit Cloud), si le fichier n'est pas trouvé
    dans ./data, on propose un téléversement via l'UI.
    """
    try:
        # Essai 1 : lire depuis le dépôt (./data/DEM_VOLO_M2.xlsx)
        return pd.read_excel(path, sheet_name=sheet)
    except Exception:
        st.warning("📄 Fichier non trouvé dans le dépôt. Téléverse le fichier Excel pour continuer.")
        uploaded = st.file_uploader("Importer DEM_VOLO_M2.xlsx", type=["xlsx"], accept_multiple_files=False)
        if uploaded is None:
            st.info("⏳ En attente du fichier…")
            st.stop()
        return pd.read_excel(uploaded, sheet_name=sheet)

# Chargement
df = load_data(DATA_PATH, SHEET)
df.columns = [str(c).strip() for c in df.columns]
if "un" not in df.columns:
    df["un"] = 1

# Mois → numérique
if "Mois saisie" in df.columns:
    df["mois_num"] = pd.to_numeric(df["Mois saisie"], errors="coerce")
elif "Mois saisie long" in df.columns:
    df["mois_num"] = df["Mois saisie long"].apply(month_from_long)

# Labels mois
df["mois_lbl_full"] = df["mois_num"].apply(lambda i: f"{int(i):02d} - {LABELS_MOIS_ASCII[int(i)-1]}" if pd.notna(i) else "")

# Dimensions
YEARS = sorted(df["Annee saisie"].dropna().astype(int).unique())
PRODS = sorted(df["PRODUIT"].dropna().astype(str).unique()) if "PRODUIT" in df.columns else []

# ----------------- UI -----------------
st.title("📊 Résiliations — Dashboard Streamlit")
st.sidebar.header("Filtres")

sel_prods = st.sidebar.multiselect("Produits", options=PRODS, default=PRODS)
sel_year_bar  = st.sidebar.selectbox("Année (barres)", YEARS, index=len(YEARS)-1)
sel_year_line = st.sidebar.selectbox("Année (courbe)", YEARS, index=len(YEARS)-2 if len(YEARS)>1 else 0)
sel_mean_years= st.sidebar.multiselect("Années pour la moyenne", YEARS, default=[y for y in [2005,2006] if y in YEARS])

df_filt = df[df["PRODUIT"].astype(str).isin(sel_prods)] if sel_prods else df.copy()

# ----------------- 1. Heatmap -----------------
st.subheader("Toutes offres — selon date de saisie")
pivot = df_filt.groupby(["mois_num","Annee saisie"])["un"].sum().unstack("Annee saisie").reindex(range(1,13))
fig1 = px.imshow(-pivot.fillna(0), 
                labels=dict(x="Année", y="Mois", color="Sorties"),
                x=pivot.columns, y=LABELS_MOIS_FR, 
                text_auto=True, aspect="auto", color_continuous_scale="Blues")
st.plotly_chart(fig1, use_container_width=True)

# ----------------- 2. Classe x Produit -----------------
if {"CLASSE_CLIENT","PRODUIT"}.issubset(df_filt.columns):
    st.subheader("Valeur client en étoiles des démissionnaires")
    g = df_filt.groupby(["CLASSE_CLIENT","PRODUIT"])["un"].sum().reset_index()
    fig2 = px.bar(g, x="CLASSE_CLIENT", y="un", color="PRODUIT", barmode="group")
    fig2.update_yaxes(dtick=2000)
    st.plotly_chart(fig2, use_container_width=True)

# ----------------- 3. Mois barres + lignes -----------------
st.subheader(f"Sorties {sel_year_bar}, {sel_year_line} et Moyenne démissions {min(sel_mean_years)}-{max(sel_mean_years)}")
serie_bar  = df_filt[df_filt["Annee saisie"]==sel_year_bar].groupby("mois_num")["un"].sum().reindex(range(1,13), fill_value=0)
serie_line = df_filt[df_filt["Annee saisie"]==sel_year_line].groupby("mois_num")["un"].sum().reindex(range(1,13), fill_value=0)
mean = (df_filt[df_filt["Annee saisie"].isin(sel_mean_years)]
        .groupby(["Annee saisie","mois_num"])["un"].sum()
        .groupby("mois_num").mean().reindex(range(1,13), fill_value=0))

fig3 = go.Figure()
fig3.add_bar(x=df["mois_lbl_full"].unique(), y=serie_bar.values, name=f"Sorties {sel_year_bar}")
fig3.add_trace(go.Scatter(x=df["mois_lbl_full"].unique(), y=serie_line.values, mode="lines+markers", name=f"Sorties {sel_year_line}"))
fig3.add_trace(go.Scatter(x=df["mois_lbl_full"].unique(), y=mean.values, mode="lines+markers", name="Moyenne"))
fig3.update_yaxes(dtick=1000)
st.plotly_chart(fig3, use_container_width=True)

# ----------------- 4. DV x Produit -----------------
if {"DV","PRODUIT"}.issubset(df_filt.columns):
    st.subheader("Résiliations par Direction des Ventes (DV) et produit")
    g2 = df_filt.groupby(["DV","PRODUIT"])["un"].sum().reset_index()
    fig4 = px.bar(g2, x="DV", y="un", color="PRODUIT", barmode="group")
    fig4.update_yaxes(dtick=500)
    st.plotly_chart(fig4, use_container_width=True)

# ----------------- 5. Annexe -----------------
st.subheader("Annexe — Écarts (réf. 2005–2006)")
annexe = pd.DataFrame({
    "Indicateur": ["Sorties 2008 — écart",
                   "Moyenne démissions 2005–2006 — écart",
                   "Démissions 2007 — écart"],
    "Valeur": [-133, 693, 613]
})
st.table(annexe)