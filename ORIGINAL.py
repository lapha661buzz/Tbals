import streamlit as st
import pandas as pd
import io
import time

st.set_page_config(page_title="Fichier Traiter", layout="wide")

# ───────────────────────── RESET ─────────────────────────
# La clé "upload_session" change à chaque reset → Streamlit recrée les uploaders vides
if "upload_session" not in st.session_state:
    st.session_state.upload_session = 0

def reset_session():
    session_id = st.session_state.upload_session + 1
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    st.session_state.upload_session = session_id
    st.rerun()

uid = st.session_state.upload_session  # suffixe unique pour les clés d'uploaders

# ───────────────────────── EN-TÊTE ─────────────────────────
col_titre, col_reset = st.columns([5, 1])
with col_titre:
    st.title("📊 Fichier Traiter")
with col_reset:
    st.write("")
    if st.button("🔄 Nouveau traitement", type="secondary", use_container_width=True,
                 help="Réinitialise l'application pour charger de nouveaux fichiers"):
        reset_session()

# ───────────────────────── LECTURE ─────────────────────────
@st.cache_data(show_spinner=False)
def read_pipe(file_bytes):
    text = ""
    for enc in ("utf-8", "latin1", "cp1252"):
        try:
            text = file_bytes.decode(enc)
            break
        except:
            continue
    lines = [l for l in text.splitlines() if l.strip()]
    if not lines:
        return pd.DataFrame()
    delim = "|" if "|" in lines[0] else "\t"
    header = [c.strip() for c in lines[0].split(delim)]
    rows = [l.split(delim) for l in lines[1:]]
    n = len(header)
    rows = [r[:n] + [""] * (n - len(r)) for r in rows]
    df = pd.DataFrame(rows, columns=header)
    df = df.loc[:, ~df.columns.duplicated()]
    return df

# ───────────────────────── OUTILS ─────────────────────────
def parse_amount(series):
    return pd.to_numeric(
        series.astype(str).str.replace(" ", "").str.replace(",", ".", regex=False),
        errors="coerce"
    )

def find_col(df, possible_names):
    for name in possible_names:
        if name in df.columns:
            return name
    return None

def traiter_stmt(df):
    key_col = find_col(df, ["CONSOL.KEY"])
    if key_col:
        df = df[df[key_col].str.strip() != ""].copy()
        crf_col = find_col(df, ["CRF.TYPE"])
        df.insert(
            df.columns.get_loc(key_col) + 1,
            "CONSOL.KEY.PRIME",
            df[key_col].astype(str) + "." + df.get(crf_col, "").astype(str)
        )
    return df

# ───────────────────────── EXCEL ─────────────────────────
@st.cache_data(show_spinner=False)
def build_excel(stmt, spec, categ, ss, all_df, tcd=None):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        stmt.to_excel(writer, "STMT", index=False)
        spec.to_excel(writer, "SPEC", index=False)
        categ.to_excel(writer, "CATEG", index=False)
        ss.to_excel(writer, "STMT+SPEC", index=False)
        all_df.to_excel(writer, "STMT+SPEC+CATEG", index=False)
        if tcd is not None:
            tcd.to_excel(writer, "TCD_TRANSACTION", index=False)
    buf.seek(0)
    return buf.getvalue()

# ───────────────────────── UPLOAD ─────────────────────────
c1, c2, c3 = st.columns(3)
with c1:
    stmt_file = st.file_uploader("Fichier STMT", key=f"uploader_stmt_{uid}")
with c2:
    spec_file = st.file_uploader("Fichier SPEC", key=f"uploader_spec_{uid}")
with c3:
    categ_file = st.file_uploader("Fichier CATEG", key=f"uploader_categ_{uid}")

# ── Compteur de progression ──
nb_charges = sum(1 for f in [stmt_file, spec_file, categ_file] if f is not None)

if nb_charges < 3:
    if nb_charges == 0:
        st.info("Veuillez charger les 3 fichiers pour commencer.")
    else:
        noms = ["STMT", "SPEC", "CATEG"]
        fichiers = [stmt_file, spec_file, categ_file]
        manquants = [noms[i] for i, f in enumerate(fichiers) if f is None]
        st.warning(f"📂 {nb_charges}/3 fichier(s) chargé(s) — En attente de : **{', '.join(manquants)}**")
        st.progress(nb_charges / 3)
    st.stop()

# ───────────────────────── TRAITEMENT EN COURS ─────────────────────────
with st.status("⏳ Traitement en cours...", expanded=True) as status:
    st.write("📥 Lecture des fichiers...")
    stmt_bytes = stmt_file.read()
    spec_bytes = spec_file.read()
    categ_bytes = categ_file.read()

    stmt_raw = read_pipe(stmt_bytes)
    st.write(f"✔️ STMT lu — {len(stmt_raw)} lignes")

    spec = read_pipe(spec_bytes)
    st.write(f"✔️ SPEC lu — {len(spec)} lignes")

    categ = read_pipe(categ_bytes)
    st.write(f"✔️ CATEG lu — {len(categ)} lignes")

    st.write("⚙️ Application des transformations STMT...")
    stmt = traiter_stmt(stmt_raw)

    st.write("🔗 Fusion des données...")
    combined_ss = pd.concat([stmt, spec], ignore_index=True)
    combined_all = pd.concat([stmt, spec, categ], ignore_index=True)

    amount_col = find_col(combined_all, ["AMOUNT.LCY"])
    trans_col = find_col(combined_all, ["TRANSACTION.CODE"])

    if amount_col:
        for df in (stmt, spec, categ, combined_ss, combined_all):
            df[amount_col] = parse_amount(df[amount_col])

    st.write("✅ Traitement terminé.")
    status.update(label="✅ Traitement terminé", state="complete", expanded=False)

# ───────────────────────── ANALYSE SOMME ─────────────────────────
tcd_probleme = None

if amount_col:
    total = combined_all[amount_col].sum()
    st.metric("Somme Totale AMOUNT.LCY", f"{total:,.2f}")

    # 🟢 CAS SOMME = 0
    if abs(total) < 1e-6:
        st.success("✅ Somme = 0 → Contrôle CAL activé")
        cal_file = st.file_uploader("📂 Charger fichier CAL", type=["txt", "csv"],
                                    key=f"uploader_cal_{uid}")
        if not cal_file:
            st.stop()

        montant_recherche = st.number_input("💰 Montant à rechercher (Ex: 198500)", value=0.0, format="%.2f")

        with st.spinner("⏳ Lecture du fichier CAL..."):
            cal = read_pipe(cal_file.read())

        id_col    = find_col(cal, ["@ID", "ID"])
        ktype_col = find_col(cal, ["K.TYPE", "KTYPE"])
        credit_col = find_col(cal, ["CREDIT.MOVEMENT"])
        debit_col  = find_col(cal, ["DEBIT.MOVEMENT"])

        if not all([id_col, ktype_col, credit_col, debit_col]):
            st.error("❌ Colonnes CAL manquantes.")
            st.stop()

        cal[credit_col] = parse_amount(cal[credit_col])
        cal[debit_col]  = parse_amount(cal[debit_col])
        cal["Som"]    = cal[credit_col].fillna(0) + cal[debit_col].fillna(0)
        cal["CONSCAL"] = cal[id_col].astype(str) + "." + cal[ktype_col].astype(str)

        key_col_prime = "CONSOL.KEY.PRIME"
        if key_col_prime in combined_ss.columns:
            with st.spinner("🔎 Recherche en cours..."):
                pivot = combined_ss.groupby(key_col_prime, sort=False)[amount_col].sum().reset_index()
                pivot.columns = ["KEY", "MONTANT"]
                merge = pivot.merge(cal[["CONSCAL", "Som"]], left_on="KEY", right_on="CONSCAL", how="left").fillna(0)
                merge["DIFF"] = merge["MONTANT"] - merge["Som"]

                target = abs(montant_recherche)
                mask_diff = (merge["DIFF"].abs() - target).abs() < 0.01
                mask_som  = (merge["Som"].abs() - target).abs() < 0.01
                problem = merge[mask_diff | mask_som]

            if montant_recherche != 0:
                if problem.empty:
                    st.warning(f"⚠️ Aucun résultat pour {montant_recherche}")
                else:
                    st.success(f"✅ {len(problem)} ligne(s) trouvée(s)")
                    st.dataframe(problem)
                    res_buf = io.BytesIO()
                    problem.to_excel(res_buf, index=False)
                    st.download_button("⬇️ Télécharger résultat", res_buf.getvalue(), "resultat.xlsx")
        st.stop()

    # 🔴 CAS SOMME != 0
    else:
        st.error(f"⚠️ Déséquilibre détecté : {total:,.2f}")

        if trans_col:
            st.subheader("📝 Analyse du déséquilibre par Code Transaction")

            with st.spinner("📊 Calcul du TCD..."):
                tcd = combined_all.groupby(trans_col)[amount_col].sum().reset_index()
                tcd.columns = [trans_col, "Somme de AMOUNT.LCY"]
                tcd_probleme = tcd[tcd["Somme de AMOUNT.LCY"].abs() > 0.001].copy()

            col_tcd, _ = st.columns([1, 1])
            with col_tcd:
                st.write("**Tableau Croisé Dynamique (TCD)**")
                st.dataframe(tcd_probleme)

            st.divider()
            st.subheader("🔎 Recherche des lignes sources (STMT + CATEG)")

            code_a_chercher = st.selectbox(
                "Choisir un code transaction à analyser :",
                tcd_probleme[trans_col].unique()
            )

            if code_a_chercher:
                with st.spinner(f"🔍 Recherche des lignes pour {code_a_chercher}..."):
                    stmt_categ = pd.concat([stmt, categ], ignore_index=True)
                    lignes_sources = stmt_categ[stmt_categ[trans_col] == code_a_chercher]

                st.write(f"Détails pour le code **{code_a_chercher}** :")
                st.dataframe(lignes_sources)

                buf_err = io.BytesIO()
                lignes_sources.to_excel(buf_err, index=False)
                st.download_button(
                    "⬇️ Télécharger ces lignes",
                    buf_err.getvalue(),
                    f"Erreur_{code_a_chercher}.xlsx"
                )

            buf_tcd = io.BytesIO()
            tcd_probleme.to_excel(buf_tcd, index=False)
            st.sidebar.download_button(
                "⬇️ Télécharger le TCD complet",
                buf_tcd.getvalue(),
                "TCD_Analyse.xlsx"
            )

# ───────────────────────── GÉNÉRATION EXCEL FINAL ─────────────────────────
st.divider()

with st.spinner("📦 Génération du fichier Excel complet..."):
    excel_data = build_excel(stmt, spec, categ, combined_ss, combined_all, tcd=tcd_probleme)

st.download_button(
    "⬇️ Télécharger Fichier Excel Complet",
    excel_data,
    "Fichier_Traiter_Global.xlsx"
)

# Aperçus
tabs = st.tabs(["STMT", "SPEC", "CATEG", "STMT+SPEC", "STMT+SPEC+CATEG"])
dfs  = [stmt, spec, categ, combined_ss, combined_all]
for tab, df in zip(tabs, dfs):
    tab.dataframe(df.head(50))