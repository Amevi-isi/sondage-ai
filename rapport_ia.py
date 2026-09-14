"""
rapport_ia.py
Génère un rapport PowerPoint complet à partir des deux formulaires KoBoToolbox.
- Formulaire 1 : Évaluation post-présentation (5 questions)
- Formulaire 2 : Évaluation initiale des connaissances IA (5 questions)
Un graphique (camembert ou barres) est créé pour chaque question.
"""

import os
import time
import argparse
import requests
from collections import Counter
from dotenv import load_dotenv
from datetime import datetime

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

load_dotenv()

# ─── CONFIGURATION ────────────────────────────────────────────────────────────
TOKEN         = os.getenv("KOBO_TOKEN")
BASE          = os.getenv("KOBO_KPI_BASE")
FORM_UID1     = os.getenv("FORM_UID1")   # Post-présentation
FORM_UID2     = os.getenv("FORM_UID2")   # Évaluation initiale
FICHIER_SORTIE = "rapport_seminaire_ia.pptx"

HEADERS = {"Authorization": f"Token {TOKEN}", "Accept": "application/json"}

# Nextcloud WebDAV
NC_URL     = os.getenv("NEXTCLOUD_URL", "").rstrip("/")
NC_USER    = os.getenv("NEXTCLOUD_USER")
NC_PASS    = os.getenv("NEXTCLOUD_PASS")
NC_DOSSIER = os.getenv("NEXTCLOUD_DOSSIER", "/SEMINAIRE").strip("/")


# ─── UPLOAD NEXTCLOUD ─────────────────────────────────────────────────────────

def envoyer_nextcloud(fichier_local):
    """
    Envoie le fichier PowerPoint sur Nextcloud via WebDAV (PUT).
    URL WebDAV : https://<host>/remote.php/dav/files/<user>/<dossier>/<fichier>
    """
    if not NC_URL or not NC_USER or not NC_PASS:
        print("⚠️  Variables NEXTCLOUD_* manquantes dans .env — envoi ignoré.")
        return

    nom_fichier = os.path.basename(fichier_local)
    webdav_url  = f"{NC_URL}/remote.php/dav/files/{NC_USER}/{NC_DOSSIER}/{nom_fichier}"

    with open(fichier_local, "rb") as f:
        r = requests.put(
            webdav_url,
            data=f,
            auth=(NC_USER, NC_PASS),
            headers={"Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation"},
        )

    if r.status_code in (200, 201, 204):
        print(f"☁️  Fichier envoyé sur Nextcloud → {NC_DOSSIER}/{nom_fichier}")
    else:
        print(f"❌ Erreur Nextcloud {r.status_code} : {r.text[:200]}")

# Couleurs thématiques
BLEU_TITRE = RGBColor(0x00, 0x45, 0x99)
BLEU_CLAIR = RGBColor(0x1F, 0x77, 0xB4)
GRIS_FOND  = RGBColor(0xF2, 0xF2, 0xF2)

# ─── DÉFINITION DES FORMULAIRES ───────────────────────────────────────────────
FORM1_CONFIG = {
    "uid": None,  # sera rempli depuis .env
    "titre": "Évaluation post-présentation IA",
    "questions": [
        {
            "champ": "_1_Apr_s_cette_pr_se_percevez_vous_l_IA_",
            "label": "1. Comment percevez-vous l'IA ?",
            "type": "pie",
            "choix": {
                "une_opportunit__strat_gique_imm_diate": "Opportunité stratégique immédiate",
                "une_opportunit__int_ressante___explorer": "Opportunité à explorer",
                "un_sujet_encore_secondaire": "Sujet encore secondaire",
                "je_reste_sceptique": "Je reste sceptique",
            },
        },
        {
            "champ": "_2_Pensez_vous_que_l_nistration_publique_",
            "label": "2. L'IA peut-elle améliorer l'administration publique ?",
            "type": "bar",
            "choix": {
                "oui_fortement": "Oui fortement",
                "oui_partiellement": "Oui partiellement",
                "peu": "Peu",
                "pas_du_tout": "Pas du tout",
            },
        },
        {
            "champ": "_4_Pensez_vous_que_c_rapidement_au_Togo_",
            "label": "4. Ces solutions peuvent-elles être déployées rapidement au Togo ?",
            "type": "pie",
            "choix": {
                "oui": "Oui",
                "oui_avec_accompagnement": "Oui avec accompagnement",
                "non": "Non",
            },
        },
        {
            "champ": "_5_Seriez_vous_favor_ans_votre_minist_re_",
            "label": "5. Favorable à des projets pilotes IA dans votre ministère ?",
            "type": "pie",
            "choix": {
                "oui_imm_diatement": "Oui immédiatement",
                "oui__apr_s__tude": "Oui, après étude",
                "non": "Non",
                "je_ne_sais_pas": "Je ne sais pas",
            },
        },
    ],
}

FORM2_CONFIG = {
    "uid": None,
    "titre": "Évaluation initiale — Connaissances IA",
    "questions": [
        {
            "champ": "_1_Avez_vous_d_j_ut_igence_artificielle_",
            "label": "1. Avez-vous déjà utilisé un outil d'IA ?",
            "type": "pie",
            "choix": {
                "oui_r_guli_rement": "Oui régulièrement",
                "oui_occasionnellement": "Oui occasionnellement",
                "non": "Non",
            },
        },
        {
            "champ": "_3_Dans_quels_domain_IA_peut_tre_utile_",
            "label": "3. Domaines où l'IA peut être utile (choix multiples)",
            "type": "bar",
            "multi": True,
            "choix": {
                "services_aux_citoyens": "Services aux citoyens",
                "sant": "Santé",
                "_ducation": "Éducation",
                "agriculture": "Agriculture",
                "gestion_administrative": "Gestion administrative",
                "aucun": "Aucun",
                "je_ne_sais_pas": "Je ne sais pas",
            },
        },
        {
            "champ": "_4_Quel_est_votre_ni_mpr_hension_de_l_IA_",
            "label": "4. Niveau de compréhension de l'IA",
            "type": "pie",
            "choix": {
                "bon": "Bon",
                "moyen": "Moyen",
                "faible": "Faible",
                "aucun": "Aucun",
            },
        },
        {
            "champ": "_5_Selon_vous_l_IA_istration_togolaise_",
            "label": "5. L'IA est-elle déjà utilisée dans l'administration togolaise ?",
            "type": "pie",
            "choix": {
                "oui": "Oui",
                "non": "Non",
                "je_ne_sais_pas": "Je ne sais pas",
            },
        },
    ],
}


# ─── RÉCUPÉRATION DES DONNÉES ─────────────────────────────────────────────────

def recuperer_soumissions(uid):
    """Récupère toutes les soumissions d'un formulaire (pagination automatique)."""
    url = f"{BASE}/{uid}/data/?format=json"
    soumissions = []
    while url:
        r = requests.get(url, headers=HEADERS)
        r.raise_for_status()
        data = r.json()
        soumissions.extend(data.get("results", []))
        url = data.get("next")
    return soumissions


def compter(soumissions, champ, multi=False):
    """
    Compte les occurrences de chaque valeur pour un champ.
    multi=True pour les select_multiple (valeurs séparées par espace).
    """
    compteur = Counter()
    for s in soumissions:
        val = s.get(champ, "")
        if not val:
            continue
        if multi:
            for v in val.split():
                compteur[v.strip()] += 1
        else:
            compteur[val.strip()] += 1
    return compteur


# ─── CONSTRUCTION DU POWERPOINT ───────────────────────────────────────────────

def ajouter_slide_titre(prs, titre_principal, sous_titre):
    """Slide de titre général."""
    layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(layout)

    slide.shapes.title.text = titre_principal
    slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = BLEU_TITRE
    slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(32)
    slide.shapes.title.text_frame.paragraphs[0].font.bold = True

    slide.placeholders[1].text = sous_titre
    slide.placeholders[1].text_frame.paragraphs[0].font.size = Pt(18)


def ajouter_slide_section(prs, titre_section, nb_repondants):
    """Slide de séparation entre les deux formulaires."""
    layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(layout)

    slide.shapes.title.text = titre_section
    slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = BLEU_TITRE
    slide.shapes.title.text_frame.paragraphs[0].font.bold = True

    slide.placeholders[1].text = f"{nb_repondants} répondant(s)"
    slide.placeholders[1].text_frame.paragraphs[0].font.size = Pt(20)


def _placer_graphique(slide, question_label, compteur, choix_labels, chart_type, x, y, w, h):
    """
    Ajoute un titre + un graphique à la position (x, y, w, h) indiquée sur une slide existante.
    """
    # Titre de la question
    txBox = slide.shapes.add_textbox(x, y, w, Inches(0.65))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = question_label
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = BLEU_TITRE

    # Résoudre les labels et valeurs
    categories = []
    valeurs = []
    for code, label in choix_labels.items():
        categories.append(label)
        valeurs.append(compteur.get(code, 0))

    # Pour le camembert, ne garder que les valeurs non nulles
    if chart_type == "pie":
        paires = [(c, v) for c, v in zip(categories, valeurs) if v > 0]
        if not paires:
            txBox2 = slide.shapes.add_textbox(x, y + Inches(0.7), w, Inches(0.5))
            txBox2.text_frame.text = "Aucune réponse"
            return
        categories, valeurs = zip(*paires)

    chart_data = ChartData()
    chart_data.categories = list(categories)
    chart_data.add_series("Réponses", list(valeurs))

    xl_type = XL_CHART_TYPE.PIE if chart_type == "pie" else XL_CHART_TYPE.COLUMN_CLUSTERED

    chart = slide.shapes.add_chart(
        xl_type,
        x, y + Inches(0.65),
        w, h - Inches(0.65),
        chart_data,
    ).chart

    chart.has_title = False
    chart.has_legend = True
    chart.legend.position = 2  # bas

    if chart_type == "pie":
        chart.plots[0].has_data_labels = True
        chart.plots[0].data_labels.show_percentage = True
        chart.plots[0].data_labels.show_value = False
    else:
        chart.plots[0].has_data_labels = True
        chart.plots[0].data_labels.show_value = True


def ajouter_slide_deux_graphiques(prs, q1, compteur1, q2, compteur2, titre_slide=""):
    """
    Crée une slide avec deux graphiques côte à côte.
    q1/q2 : dict de question (label, choix, type)
    Si q2 est None, un seul graphique est centré.
    titre_slide : nom affiché en haut de la slide.
    """
    layout = prs.slide_layouts[5]  # layout vide
    slide = prs.slides.add_slide(layout)

    # Supprimer tous les placeholders hérités du layout (évite "Click to add title")
    for ph in slide.placeholders:
        sp = ph._element
        sp.getparent().remove(sp)

    # ── Titre de la slide ──────────────────────────────────────────────────────
    if titre_slide:
        tb = slide.shapes.add_textbox(Inches(0.2), Inches(0.05), Inches(9.6), Inches(0.4))
        tf = tb.text_frame
        p  = tf.paragraphs[0]
        p.text = titre_slide
        p.font.size  = Pt(13)
        p.font.bold  = True
        p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        # Fond coloré simulé via la forme
        from pptx.oxml.ns import qn
        from lxml import etree
        spPr = tb._element.spPr
        solidFill = etree.SubElement(spPr, qn('a:solidFill'))
        srgbClr   = etree.SubElement(solidFill, qn('a:srgbClr'))
        srgbClr.set('val', '004599')

    # ── Dimensions des graphiques (décalés sous le titre) ─────────────────────
    marge   = Inches(0.2)
    largeur = Inches(4.7)
    top     = Inches(0.52) if titre_slide else Inches(0.2)
    hauteur = Inches(7.5) - top - Inches(0.1)

    _placer_graphique(slide, q1["label"], compteur1, q1["choix"], q1["type"],
                      marge, top, largeur, hauteur)

    if q2 is not None:
        _placer_graphique(slide, q2["label"], compteur2, q2["choix"], q2["type"],
                          marge + largeur + Inches(0.2), top, largeur, hauteur)
    return slide


def compter_soumissions(uid):
    """Retourne uniquement le nombre total de soumissions sans tout télécharger."""
    r = requests.get(f"{BASE}/{uid}/data/?format=json&limit=1", headers=HEADERS)
    r.raise_for_status()
    return r.json().get("count", 0)


# ─── MODE WATCH ───────────────────────────────────────────────────────────────

def watch(intervalle_sec=5):
    """
    Surveille les deux formulaires toutes les `intervalle_sec` secondes.
    Régénère le PowerPoint uniquement quand une nouvelle soumission est détectée.
    Arrêt propre avec Ctrl+C.
    """
    print(f"👁  Mode surveillance actif — vérification toutes les {intervalle_sec}s  (Ctrl+C pour arrêter)")
    compteurs = {FORM_UID1: -1, FORM_UID2: -1}

    while True:
        try:
            n1 = compter_soumissions(FORM_UID1)
            n2 = compter_soumissions(FORM_UID2)
            if n1 != compteurs[FORM_UID1] or n2 != compteurs[FORM_UID2]:
                nouvelle_soumission = compteurs[FORM_UID1] != -1  # pas le premier lancement
                if nouvelle_soumission:
                    print(f"🔔 Nouvelle soumission détectée  "
                          f"(F1: {compteurs[FORM_UID1]}→{n1}  F2: {compteurs[FORM_UID2]}→{n2})")
                compteurs[FORM_UID1] = n1
                compteurs[FORM_UID2] = n2
                main(envoyer=nouvelle_soumission)
            else:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] Aucun changement "
                      f"(F1: {n1}  F2: {n2})  — prochain check dans {intervalle_sec}s")
            time.sleep(intervalle_sec)
        except KeyboardInterrupt:
            print("\n⏹  Surveillance arrêtée.")
            break
        except Exception as e:
            print(f"⚠️  Erreur : {e}  — nouvelle tentative dans {intervalle_sec}s")
            time.sleep(intervalle_sec)


# ─── PIPELINE PRINCIPAL ───────────────────────────────────────────────────────

def main(envoyer=True):
    FORM1_CONFIG["uid"] = FORM_UID1
    FORM2_CONFIG["uid"] = FORM_UID2

    # — Récupération —
    print("⏳ Récupération formulaire 1 (post-présentation)…")
    s1 = recuperer_soumissions(FORM_UID1)
    print(f"   → {len(s1)} soumission(s)")

    print("⏳ Récupération formulaire 2 (évaluation initiale)…")
    s2 = recuperer_soumissions(FORM_UID2)
    print(f"   → {len(s2)} soumission(s)")

    # — Création du PowerPoint —
    prs = Presentation()
    prs.slide_width  = Inches(10)
    prs.slide_height = Inches(7.5)

    # Slide de titre général
    ajouter_slide_titre(
        prs,
        "Rapport — Intelligence Artificielle",
        f"Analyse des formulaires KoBoToolbox  •  {datetime.today().strftime('%d/%m/%Y')}",
    )

    # ── FORMULAIRE 2 : Évaluation initiale ────────────────────────────────────
    ajouter_slide_section(prs, FORM2_CONFIG["titre"], len(s2))
    questions2 = FORM2_CONFIG["questions"]
    nb_slides2 = (len(questions2) + 1) // 2
    for idx, i in enumerate(range(0, len(questions2), 2)):
        q1 = questions2[i]
        q2 = questions2[i + 1] if i + 1 < len(questions2) else None
        c1 = compter(s2, q1["champ"], multi=q1.get("multi", False))
        c2 = compter(s2, q2["champ"], multi=q2.get("multi", False)) if q2 else None
        titre = f"{FORM2_CONFIG['titre']}  ({idx + 1}/{nb_slides2})"
        ajouter_slide_deux_graphiques(prs, q1, c1, q2, c2, titre_slide=titre)

    # ── FORMULAIRE 1 : Post-présentation ──────────────────────────────────────
    ajouter_slide_section(prs, FORM1_CONFIG["titre"], len(s1))
    questions1 = FORM1_CONFIG["questions"]
    nb_slides1 = (len(questions1) + 1) // 2
    for idx, i in enumerate(range(0, len(questions1), 2)):
        q1 = questions1[i]
        q2 = questions1[i + 1] if i + 1 < len(questions1) else None
        c1 = compter(s1, q1["champ"], multi=q1.get("multi", False))
        c2 = compter(s1, q2["champ"], multi=q2.get("multi", False)) if q2 else None
        titre = f"{FORM1_CONFIG['titre']}  ({idx + 1}/{nb_slides1})"
        ajouter_slide_deux_graphiques(prs, q1, c1, q2, c2, titre_slide=titre)

    # — Sauvegarde —
    prs.save(FICHIER_SORTIE)
    nb_slides = len(prs.slides)
    print(f"✅ [{datetime.now().strftime('%H:%M:%S')}] {nb_slides} slides générées → {FICHIER_SORTIE}")

    # — Envoi sur Nextcloud uniquement si nouvelles soumissions —
    if envoyer:
        envoyer_nextcloud(FICHIER_SORTIE)
    else:
        print("ℹ️  Aucune nouvelle soumission — fichier Nextcloud conservé.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Rapport IA — KoBoToolbox → PowerPoint")
    parser.add_argument(
        "--watch",
        nargs="?",
        const=5,
        type=int,
        metavar="SECONDES",
        help="Mode surveillance : régénère le PowerPoint à chaque nouvelle soumission. "
             "Intervalle par défaut : 5s. Ex: --watch 120",
    )
    args = parser.parse_args()

    if args.watch is not None:
        watch(intervalle_sec=args.watch)
    else:
        main()
