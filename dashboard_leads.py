import json
import re
import time

import pandas as pd
import streamlit as st
from openai import OpenAI

st.set_page_config(page_title="Qualification de leads", page_icon="🏠", layout="wide")

COULEURS_SCORE = {"chaud": "#C6E0B4", "tiède": "#FFE699", "froid": "#F4CCCC", "hors_sujet": "#D9D9D9"}
EMOJI_SCORE = {"chaud": "🔥", "tiède": "🌤️", "froid": "❄️", "hors_sujet": "🚫"}

SCHEMA_ANALYSE = {
    "type": "object",
    "properties": {
        "nom_probable": {"type": ["string", "null"]},
        "budget": {"type": ["string", "null"]},
        "delai": {"type": ["string", "null"]},
        "type_de_bien": {"type": ["string", "null"]},
        "ville": {"type": ["string", "null"]},
        "score": {"type": "string", "enum": ["chaud", "tiède", "froid", "hors_sujet"]},
        "resume": {"type": "string"},
        "justification": {"type": "string"},
        "action_recommandee": {"type": "string"},
    },
    "required": [
        "nom_probable", "budget", "delai", "type_de_bien", "ville",
        "score", "resume", "justification", "action_recommandee",
    ],
    "additionalProperties": False,
}

INSTRUCTIONS = """
Tu analyses des prospects immobiliers à partir d'emails ou messages bruts.

Réponds exclusivement en français standard. N'utilise aucun mot, expression
ou caractère provenant d'une autre langue ou d'un autre alphabet, même isolé.

Extrais uniquement les informations présentes dans le message.
Ne devine jamais une information manquante : utilise null.
Si un nom d'expéditeur est identifiable dans le message (signature, en-tête),
extrais-le dans nom_probable, sinon laisse null.

Règles de score :
- chaud : budget connu et projet dans moins de 3 mois ;
- tiède : projet réel, mais une information importante manque ;
- froid : demande vague, simple information ou projet lointain ;
- hors_sujet : le message n'est PAS une demande d'achat, de vente ou de location
  immobilière (spam, publicité, question pratique, plainte, message vide ou
  incompréhensible, autre langue). Ne force JAMAIS un message hors_sujet dans
  une autre catégorie.

Propose une action simple et concrète.
"""

CARACTERES_ILLEGAUX = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
ALPHABET_INATTENDU = re.compile(r"[\u0400-\u04FF\u4e00-\u9fff\u0600-\u06FF]")
ECHAPPEMENT_LITTERAL = re.compile(r"\\[nrtx]|[\\][0-9a-fA-F]{2}")

MOTS_CLES_SPAM = [
    "promo", "-50%", "-30%", "offre limitée", "cliquez ici",
    "profitez-en", "gratuit", "sans engagement", "1xbet", "casino",
]


def nettoyer(texte):
    if texte is None:
        return texte
    return CARACTERES_ILLEGAUX.sub("", str(texte))


def contient_probleme(*valeurs):
    for v in valeurs:
        if v and (ALPHABET_INATTENDU.search(str(v)) or ECHAPPEMENT_LITTERAL.search(str(v))):
            return True
    return False


def pre_filtre_local(message):
    if not message or not message.strip():
        return "Message vide"
    message_normalise = message.lower()
    if len(message.split()) <= 3 and not any(c.isdigit() for c in message):
        return "Message trop court pour être une demande exploitable"
    for mot_cle in MOTS_CLES_SPAM:
        if mot_cle in message_normalise:
            return f"Mots-clés publicitaires détectés ({mot_cle})"
    return None


def resultat_hors_sujet(raison):
    return {
        "nom_probable": None, "budget": None, "delai": None, "type_de_bien": None,
        "ville": None, "score": "hors_sujet", "resume": raison,
        "justification": f"Écarté automatiquement avant analyse : {raison}",
        "action_recommandee": "Aucune action requise",
    }


def analyser_message(client, message, index):
    raison_filtre = pre_filtre_local(message)
    if raison_filtre:
        return resultat_hors_sujet(raison_filtre), True  # True = filtré localement

    for tentative in range(3):
        try:
            response = client.responses.create(
                model="gpt-5.2",
                instructions=INSTRUCTIONS,
                input=f"Message :\n{message}",
                text={
                    "format": {
                        "type": "json_schema",
                        "name": "analyse_lead",
                        "strict": True,
                        "schema": SCHEMA_ANALYSE,
                    }
                },
                temperature=0.1,
                timeout=30,
            )
            analyse = json.loads(response.output_text)
            if contient_probleme(
                analyse["resume"], analyse["action_recommandee"], analyse["justification"]
            ):
                if tentative < 2:
                    continue
                else:
                    analyse["resume"] = nettoyer(analyse["resume"])
            return analyse, False
        except Exception as erreur:
            if tentative < 2:
                time.sleep(1.5 * (tentative + 1))
                continue
            return {
                "nom_probable": None, "budget": None, "delai": None, "type_de_bien": None,
                "ville": None, "score": "hors_sujet",
                "resume": f"Erreur de traitement : {erreur}",
                "justification": "Ce message n'a pas pu être analysé après plusieurs tentatives.",
                "action_recommandee": "À traiter manuellement",
            }, False


def afficher_badge(score):
    couleur = COULEURS_SCORE.get(score, "#EEEEEE")
    emoji = EMOJI_SCORE.get(score, "")
    return f"""<span style="background-color:{couleur}; padding:3px 10px;
    border-radius:12px; font-weight:600;">{emoji} {score}</span>"""


def generer_excel_stylise(df):
    """Reproduit la mise en forme du script original : en-tête colorée,
    lignes colorées par score, colonnes ajustées, tri chaud→froid, filtres."""
    import io
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    ordre_score = {"chaud": 0, "tiède": 1, "froid": 2, "hors_sujet": 3}
    df_trie = df.assign(_ordre=df["Score"].map(ordre_score)).sort_values("_ordre").drop(columns="_ordre")

    buffer = io.BytesIO()
    df_trie.to_excel(buffer, index=False)
    buffer.seek(0)

    classeur = load_workbook(buffer)
    feuille = classeur.active
    feuille.title = "Résultats"

    couleur_entete = PatternFill("solid", fgColor="1F4E78")
    for cellule in feuille[1]:
        cellule.font = Font(color="FFFFFF", bold=True)
        cellule.fill = couleur_entete
        cellule.alignment = Alignment(horizontal="center")

    largeurs = {"A": 20, "B": 12, "C": 45, "D": 35, "E": 18, "F": 18, "G": 20, "H": 18, "I": 45}
    for colonne, largeur in largeurs.items():
        feuille.column_dimensions[colonne].width = largeur

    for ligne in feuille.iter_rows(min_row=2):
        for cellule in ligne:
            cellule.alignment = Alignment(vertical="top", wrap_text=True)
        score = ligne[1].value
        if score == "chaud":
            ligne[1].fill = PatternFill("solid", fgColor="C6E0B4")
        elif score == "tiède":
            ligne[1].fill = PatternFill("solid", fgColor="FFE699")
        elif score == "froid":
            ligne[1].fill = PatternFill("solid", fgColor="F4CCCC")
        elif score == "hors_sujet":
            ligne[1].fill = PatternFill("solid", fgColor="D9D9D9")

    feuille.freeze_panes = "A2"
    feuille.auto_filter.ref = feuille.dimensions

    buffer_final = io.BytesIO()
    classeur.save(buffer_final)
    return buffer_final.getvalue()


st.title("🏠 Qualification de leads immobiliers")
st.caption("Collez vos demandes reçues, obtenez un tri par priorité en quelques secondes.")

with st.sidebar:
    st.header("Configuration")
    api_key = st.text_input(
        "Votre clé API OpenAI",
        type="password",
        help="Votre clé n'est jamais enregistrée : elle n'est utilisée que le temps de cette session, "
             "et disparaît dès que vous fermez ou rechargez la page.",
    )
    st.caption("Besoin d'une clé ? platform.openai.com/api-keys")

st.subheader("1. Collez vos messages")
st.caption('Séparez chaque email/message par une ligne contenant seulement : ---')

texte_brut = st.text_area(
    "Messages",
    height=250,
    placeholder="Bonjour, je cherche un T3 à Lyon, budget 320 000€, achat avant juin.\n"
                "---\n"
                "Bonjour, nous vendons notre maison, estimation souhaitée rapidement.\n"
                "---\n"
                "(un message par bloc, séparés par ---)",
    label_visibility="collapsed",
)

if texte_brut.strip():
    nb_detectes = len([m for m in texte_brut.split("---") if m.strip()])
    if nb_detectes == 1 and "---" not in texte_brut:
        st.warning(
            "⚠️ 1 seul message détecté. Si vous avez collé plusieurs emails, "
            "vérifiez qu'ils sont bien séparés par une ligne contenant uniquement : ---"
        )
    else:
        st.caption(f"✅ {nb_detectes} message(s) détecté(s), prêt(s) à analyser.")

lancer = st.button("Analyser", type="primary", disabled=not api_key or not texte_brut.strip())

if not api_key and texte_brut.strip():
    st.warning("Entrez votre clé API OpenAI dans la barre latérale pour lancer l'analyse.")

if lancer:
    messages = [m.strip() for m in texte_brut.split("---") if m.strip()]
    client = OpenAI(api_key=api_key)

    resultats = []
    nb_filtres_localement = 0
    barre = st.progress(0, text=f"0/{len(messages)} messages traités")

    for i, message in enumerate(messages, start=1):
        analyse, filtre_local = analyser_message(client, message, i)
        if filtre_local:
            nb_filtres_localement += 1

        resultats.append({
            "Lead": nettoyer(analyse.get("nom_probable")) or f"Lead {i}",
            "Score": nettoyer(analyse["score"]),
            "Résumé": nettoyer(analyse["resume"]),
            "Action recommandée": nettoyer(analyse["action_recommandee"]),
            "Budget": nettoyer(analyse["budget"]) or "Non communiqué",
            "Délai": nettoyer(analyse["delai"]) or "Non communiqué",
            "Type de bien": nettoyer(analyse["type_de_bien"]) or "Non communiqué",
            "Ville": nettoyer(analyse["ville"]) or "Non communiquée",
            "Justification": nettoyer(analyse["justification"]),
        })
        barre.progress(i / len(messages), text=f"{i}/{len(messages)} messages traités")

    barre.empty()
    st.session_state["resultats"] = pd.DataFrame(resultats)
    st.session_state["nb_filtres_localement"] = nb_filtres_localement

if "resultats" in st.session_state:
    df = st.session_state["resultats"]

    st.divider()
    st.subheader("2. Résultats")

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
        nb_locaux = st.session_state.get("nb_filtres_localement", 0)
        st.caption(
            f"🚫 {nb_hors_sujet} message(s) écarté(s) automatiquement (dont {nb_locaux} sans appel API) "
            "— non comptés ci-dessus."
        )

    with st.sidebar:
        st.divider()
        st.header("Filtres")
        scores_choisis = st.multiselect(
            "Score", options=["chaud", "tiède", "froid", "hors_sujet"],
            default=["chaud", "tiède", "froid"],
        )
        villes_disponibles = sorted(df["Ville"].dropna().unique().tolist())
        villes_choisies = st.multiselect("Ville", options=villes_disponibles, default=villes_disponibles)

    df_filtre = df[
        df["Score"].isin(scores_choisis) & (df["Ville"].isin(villes_choisies) | df["Ville"].isna())
    ]
    ordre_score = {"chaud": 0, "tiède": 1, "froid": 2, "hors_sujet": 3}
    df_filtre = df_filtre.assign(_ordre=df_filtre["Score"].map(ordre_score)).sort_values("_ordre")

    st.markdown(f"**{len(df_filtre)} lead(s) affiché(s)**")

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

    st.divider()
    st.download_button(
        "📥 Télécharger le rapport complet (.xlsx)",
        data=generer_excel_stylise(df),
        file_name="resultats_leads.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.divider()
st.caption("Les scores et recommandations sont produits par IA et méritent une vérification humaine sur les leads les plus prioritaires. Votre clé API n'est jamais stockée sur nos serveurs.")
