import pandas as pd
import numpy as np
from openpyxl import load_workbook
import streamlit as st
from io import BytesIO

# ========== Interface utilisateur ==========
st.title("📊 Générateur d'écritures analytiques - BLDD")

fichier_entree = st.file_uploader("📂 Importer le fichier Excel BLDD", type=["xlsx"])
date_ecriture = st.date_input("📅 Date d'écriture")
journal = st.text_input("📒 Journal", value="VT")
libelle_base = st.text_input("📝 Libellé", value="VENTES BLDD")

compte_ca = st.text_input("💰 Compte CA", value="70110000")
compte_com_dist = st.text_input("💰 Compte commissions distribution", value="62280000")
compte_com_diff = st.text_input("💰 Compte commissions diffusion", value="62280001")

# 🔹 Saisie des taux
taux_dist = st.number_input("Taux distribution (%)", value=12.5) / 100
taux_diff = st.number_input("Taux diffusion (%)", value=9.0) / 100

# 🔹 Saisie des montants totaux
com_distribution_total = st.number_input("Montant total commissions distribution", value=1000.00, format="%.2f")
com_diffusion_total = st.number_input("Montant total commissions diffusion", value=500.00, format="%.2f")

# ========== Traitement ==========
if fichier_entree is not None:
    df = pd.read_excel(fichier_entree, header=9, dtype={"ISBN": str})
    df.columns = df.columns.str.strip()
    df = df.dropna(subset=["ISBN"]).copy()

    df["ISBN"] = df["ISBN"].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    df["ISBN"] = df["ISBN"].str.replace('-', '', regex=False).str.replace(' ', '', regex=False)

    for c in ["Vente", "Net", "Facture"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).round(2)

    # ========== Distribution ==========
    raw_dist = df["Vente"] * taux_dist
    sum_raw_dist = raw_dist.sum()
    scaled_dist = raw_dist * (com_distribution_total / sum_raw_dist)

    cents_floor = np.floor(scaled_dist * 100).astype(int)
    remainders = (scaled_dist * 100) - cents_floor
    target_cents = int(round(com_distribution_total * 100))
    diff = target_cents - cents_floor.sum()
    idx_sorted = np.argsort(-remainders.values)
    adjust = np.zeros(len(df), dtype=int)
    if diff > 0:
        adjust[idx_sorted[:diff]] = 1
    elif diff < 0:
        adjust[idx_sorted[len(df)+diff:]] = -1

    df["Commission_distribution"] = (cents_floor + adjust) / 100.0

    # ========== Diffusion ==========
    raw_diff = df["Net"] * taux_diff
    sum_raw_diff = raw_diff.sum()
    scaled_diff = raw_diff * (com_diffusion_total / sum_raw_diff)

    cents_floor = np.floor(scaled_diff * 100.0).astype(int)
    remainders = (scaled_diff * 100.0) - cents_floor
    target_cents = int(round(com_diffusion_total * 100))
    diff = target_cents - cents_floor.sum()
    idx_sorted = np.argsort(-remainders.values)
    adjust = np.zeros(len(df), dtype=int)
    if diff > 0:
        adjust[idx_sorted[:diff]] = 1
    elif diff < 0:
        adjust[idx_sorted[len(df)+diff:]] = -1

    df["Commission_diffusion"] = (cents_floor + adjust) / 100.0

    # ========== Construction écritures ==========
    ecritures = []
    total_facture_global = df["Facture"].sum().round(2)

    # CA global
    ecritures.append({
        "Date": date_ecriture.strftime("%d/%m/%Y"), "Journal": journal, "Compte": compte_ca,
        "Libelle": f"{libelle_base} - CA global", "ISBN": "",
        "Débit": total_facture_global, "Crédit": 0.0
    })
    # CA ISBN
    for _, r in df.iterrows():
        ecritures.append({
            "Date": date_ecriture.strftime("%d/%m/%Y"), "Journal": journal, "Compte": compte_ca,
            "Libelle": f"{libelle_base} - CA ISBN", "ISBN": r["ISBN"],
            "Débit": 0.0, "Crédit": round(float(r["Facture"]), 2)
        })

    # Commissions distribution
    total_dist = df["Commission_distribution"].sum().round(2)
    ecritures.append({
        "Date": date_ecriture.strftime("%d/%m/%Y"), "Journal": journal, "Compte": compte_com_dist,
        "Libelle": f"{libelle_base} - Com. distribution global", "ISBN": "",
        "Débit": 0.0, "Crédit": total_dist
    })
    for _, r in df.iterrows():
        ecritures.append({
            "Date": date_ecriture.strftime("%d/%m/%Y"), "Journal": journal, "Compte": compte_com_dist,
            "Libelle": f"{libelle_base} - Com. distribution ISBN", "ISBN": r["ISBN"],
            "Débit": round(float(r["Commission_distribution"]), 2), "Crédit": 0.0
        })

    # Commissions diffusion
    total_diff = df["Commission_diffusion"].sum().round(2)
    ecritures.append({
        "Date": date_ecriture.strftime("%d/%m/%Y"), "Journal": journal, "Compte": compte_com_diff,
        "Libelle": f"{libelle_base} - Com. diffusion global", "ISBN": "",
        "Débit": 0.0, "Crédit": total_diff
    })
    for _, r in df.iterrows():
        ecritures.append({
            "Date": date_ecriture.strftime("%d/%m/%Y"), "Journal": journal, "Compte": compte_com_diff,
            "Libelle": f"{libelle_base} - Com. diffusion ISBN", "ISBN": r["ISBN"],
            "Débit": round(float(r["Commission_diffusion"]), 2), "Crédit": 0.0
        })

    df_ecr = pd.DataFrame(ecritures)

    # Vérification équilibre
    total_debit = round(df_ecr["Débit"].sum(), 2)
    total_credit = round(df_ecr["Crédit"].sum(), 2)

    if total_debit != total_credit:
        st.error(f"⚠️ Écriture déséquilibrée : Débit={total_debit}, Crédit={total_credit}")
    else:
        st.success("✅ Écritures équilibrées !")

    # ========== Export & téléchargement ==========
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_ecr.to_excel(writer, index=False, sheet_name="Ecritures")
    buffer.seek(0)

    st.download_button(
        label="📥 Télécharger les écritures (Excel)",
        data=buffer,
        file_name="Ecritures_Pennylane.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Aperçu dans l’appli
    st.subheader("👀 Aperçu des écritures générées")
    st.dataframe(df_ecr)