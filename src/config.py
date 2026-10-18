"""
Configuration générale de l'application.
Définit les chemins des répertoires, les constantes et les types de documents supportés.
"""

from pathlib import Path

# Chemins des répertoires de base
BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"

# Fichier pour la sauvegarde du compteur de numérotation
COMPTEUR_FILE = DATA_DIR / "compteur.json"

# Fichier exemple Excel et CSV
EXCEL_TEST_FILE = DATA_DIR / "donnees_test.xlsx"
CSV_TEST_FILE = DATA_DIR / "donnees_test.csv"

# Types de documents pris en charge
DOC_TYPES = {
    "attestation": {
        "label": "Attestation",
        "prefix": "ATT",
        "default_title": "ATTESTATION DE PARTICIPATION",
    },
    "convocation": {
        "label": "Convocation",
        "prefix": "CONV",
        "default_title": "CONVOCATION OFFICIELLE",
    },
    "recu": {
        "label": "Reçu",
        "prefix": "REC",
        "default_title": "REÇU DE PAIEMENT",
    },
    "bulletin": {
        "label": "Bulletin",
        "prefix": "BULL",
        "default_title": "BULLETIN D'ÉVALUATION",
    },
}

# Informations par défaut de l'émetteur
ORGANISATION_DEFAULT = {
    "nom": "CENTRE DE FORMATION & SERVICES",
    "adresse": "123 Avenue de la République, 75011 Paris",
    "telephone": "+33 1 23 45 67 89",
    "email": "contact@centre-services.fr",
    "siret": "123 456 789 00012",
}

# Création automatique des répertoires nécessaires
for directory in (DATA_DIR, TEMPLATES_DIR, OUTPUT_DIR):
    directory.mkdir(parents=True, exist_ok=True)
