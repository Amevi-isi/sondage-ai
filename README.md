# Rapport IA — KoBoToolbox → PowerPoint → Nextcloud

Génère automatiquement un rapport PowerPoint à partir des réponses de deux formulaires KoBoToolbox, puis l'envoie sur Nextcloud.

## Prérequis

Python 3.8+ requis. Installez les dépendances via :

```bash
pip install -r requirements.txt
```

Packages utilisés :

| Package | Version |
|---------|---------|
| python-pptx | 1.0.2 |
| requests | 2.32.5 |
| python-dotenv | 1.2.2 |

## Configuration

Créez un fichier `.env` à la racine du projet :

```env
# KoBoToolbox
KOBO_TOKEN=<votre_token_kpi>
KOBO_KPI_BASE=https://kf.form.gouv.tg/api/v2/assets
FORM_UID1=<uid_formulaire_post_presentation>
FORM_UID2=<uid_formulaire_evaluation_initiale>

# Nextcloud WebDAV
NEXTCLOUD_URL=https://cloud.digital.gouv.tg
NEXTCLOUD_USER=<votre_email>
NEXTCLOUD_PASS=<mot_de_passe_application>   # Paramètres → Sécurité → Mots de passe d'application
NEXTCLOUD_DOSSIER=/SEMINAIRE
```

> **Important** : `NEXTCLOUD_PASS` doit être un **mot de passe d'application** Nextcloud, pas votre mot de passe habituel.  
> Créez-en un dans : *Paramètres → Sécurité → Mots de passe d'application*.

## Utilisation

### Génération unique

```bash
python3 rapport_ia.py
```

Récupère les données, génère `rapport_ia.pptx` et l'envoie sur Nextcloud.

### Mode surveillance automatique

```bash
python3 rapport_ia.py --watch
```

Vérifie les formulaires toutes les **60 secondes**. Régénère et renvoie le PowerPoint sur Nextcloud uniquement si une nouvelle soumission est détectée. Arrêt avec **Ctrl+C**.

Intervalle personnalisé (ex : 2 minutes) :

```bash
python3 rapport_ia.py --watch 120
```

## Structure du rapport généré

| Slide | Contenu |
|-------|---------|
| 1 | Titre général + date |
| 2 | Section — Évaluation initiale (formulaire 2) |
| 3 | Q1 Usage IA + Q3 Domaines utiles |
| 4 | Q4 Niveau de compréhension + Q5 IA dans l'administration |
| 5 | Section — Post-présentation (formulaire 1) |
| 6 | Q1 Perception IA + Q2 Améliorer l'administration |
| 7 | Q4 Déploiement rapide + Q5 Projets pilotes |

## Fichiers à déployer

```
rapport_ia.py       ← script principal
requirements.txt    ← dépendances Python
.env                ← credentials (ne pas partager publiquement)
README.md           ← documentation
```

> ⚠️ Si vous partagez le projet (Git, clé USB), **excluez le `.env`** car il contient les tokens et mots de passe.

## Formulaires couverts

| Formulaire | Titre | Questions retenues |
|------------|-------|--------------------|
| FORM_UID2 | Évaluation initiale des connaissances IA | Q1, Q3, Q4, Q5 |
| FORM_UID1 | Évaluation post-présentation IA | Q1, Q2, Q4, Q5 |
