import pandas as pd
import streamlit as st

st.set_page_config(page_title="Qualification de leads", page_icon="🏠", layout="wide")

COULEURS_SCORE = {"chaud": "#C6E0B4", "tiède": "#FFE699", "froid": "#F4CCCC", "hors_sujet": "#D9D9D9"}
EMOJI_SCORE = {"chaud": "🔥", "tiède": "🌤️", "froid": "❄️", "hors_sujet": "🚫"}


@st.cache_data
def charger_donnees(fichier):
    df = pd.read_excel(fichier)
    return df


def afficher_badge(score):
    couleur = COULEURS_SCORE.get(score, "#EEEEEE")
    emoji = EMOJI_SCORE.get(score, "")
    return f"""<span style="background-color:{couleur}; padding:3px 10px;
    border-radius:12px; font-weight:600;">{emoji} {score}</span>"""


st.title("🏠 Qualification de leads immobiliers")
st.caption("Tableau de bord généré automatiquement à partir des demandes reçues")

fichier = st.file_uploader("Charger un fichier de résultats (.xlsx)", type=["xlsx"])

if fichier is None:
    st.info("Charge le fichier resultats_leads.xlsx généré par le script pour voir le tableau de bord.")
    st.stop()

df = charger_donnees(fichier)

# --- Indicateurs clés (les messages hors sujet ne comptent pas comme des leads) ---
nb_hors_sujet = (df["Score"] == "hors_sujet").sum()
df_leads = df[df["Score"] != "hors_sujet"]

col1, col2, col3, col4 = st.columns(4)
total = len(df_leads)
nb_chaud = (df_leads["Score"] == "chaud").sum()
nb_tiede = (df_leads["Score"] == "tiède").sum()
nb_froid = (df_leads["Score"] == "froid").sum()

col1.metric("Total leads", total)
col2.metric("🔥 Chauds", nb_chaud, f"{nb_chaud / total:.0%}" if total else "0%")
col3.metric("🌤️ Tièdes", nb_tiede, f"{nb_tiede / total:.0%}" if total else "0%")
col4.metric("❄️ Froids", nb_froid, f"{nb_froid / total:.0%}" if total else "0%")

if nb_hors_sujet:
    st.caption(f"🚫 {nb_hors_sujet} message(s) écarté(s) automatiquement (hors sujet, spam, ou non pertinent) — non comptés ci-dessus.")

st.divider()

# --- Filtres ---
with st.sidebar:
    st.header("Filtres")
    scores_choisis = st.multiselect(
        "Score",
        options=["chaud", "tiède", "froid", "hors_sujet"],
        default=["chaud", "tiède", "froid"],
        help="« hors_sujet » regroupe les messages écartés automatiquement (spam, non pertinent) — décoché par défaut.",
    )
    villes_disponibles = sorted(df["Ville"].dropna().unique().tolist())
    villes_choisies = st.multiselect("Ville", options=villes_disponibles, default=villes_disponibles)

df_filtre = df[
    df["Score"].isin(scores_choisis)
    & (df["Ville"].isin(villes_choisies) | df["Ville"].isna())
]

# --- Liste des leads, triés par priorité ---
ordre_score = {"chaud": 0, "tiède": 1, "froid": 2, "hors_sujet": 3}
df_filtre = df_filtre.assign(_ordre=df_filtre["Score"].map(ordre_score)).sort_values("_ordre")

st.subheader(f"Leads ({len(df_filtre)})")

for _, ligne in df_filtre.iterrows():
    with st.container(border=True):
        c1, c2 = st.columns([3, 1])
        with c1:
            st.markdown(f"**{ligne['Lead']}** — {ligne['Ville'] if pd.notna(ligne['Ville']) else 'Ville non communiquée'}")
            st.markdown(afficher_badge(ligne["Score"]), unsafe_allow_html=True)
        with c2:
            st.markdown(f"**Budget :** {ligne['Budget']}")
            st.markdown(f"**Délai :** {ligne['Délai']}")

        st.write(ligne["Résumé"])

        with st.expander("Voir le détail et l'action recommandée"):
            st.markdown(f"**Type de bien :** {ligne['Type de bien']}")
            st.markdown(f"**Justification du score :** {ligne['Justification']}")
            st.markdown(f"**Action recommandée :** {ligne['Action recommandée']}")
            if ligne.get("À relire") == "OUI":
                st.warning("Cette ligne a été signalée pour relecture manuelle.")

st.divider()
st.caption("Dashboard généré automatiquement — les scores et recommandations sont produits par IA et méritent une vérification humaine sur les leads les plus prioritaires.")