"""
Module de gestion des modèles de documents personnalisables utilisant Jinja2.
Permet de charger, modifier, sauvegarder et rendre les textes des documents.
"""

import json
from pathlib import Path
from typing import Any, Dict

import jinja2
from src.config import TEMPLATES_DIR


MODELES_PAR_DEFAUT = {
    "attestation": {
        "titre": "ATTESTATION DE PARTICIPATION",
        "sous_titre": "Attestation officielle de présence et de suivi",
        "corps": (
            "Je soussigné(e), représentant de la direction, certifie par la présente que :\n\n"
            "{{ civilite }} {{ prenom }} {{ nom }}\n"
            "(Contact : {{ email if email else 'Non renseigné' }})\n\n"
            "a participé avec assiduité et succès à la session suivante :\n"
            "« {{ motif if motif else 'Formation professionnelle' }} ».\n\n"
            "Date de réalisation : {{ date }}\n"
            "Numéro de référence : {{ numero }}\n\n"
            "La présente attestation est délivrée à l'intéressé(e) pour servir et valoir ce que de droit."
        ),
        "signature_titre": "La Direction Générale",
        "pied_page_texte": "Document officiel émis avec numérotation unique vérifiable.",
    },
    "convocation": {
        "titre": "CONVOCATION OFFICIELLE",
        "sous_titre": "Convocation administrative",
        "corps": (
            "À l'attention de : {{ civilite }} {{ prenom }} {{ nom }}\n\n"
            "Vous êtes prié(e) de bien vouloir vous présenter le {{ date }} afin de prendre part à :\n"
            "« {{ motif if motif else 'Assemblée générale' }} ».\n\n"
            "Lieu de rendez-vous : Siège administratif / Salle de réunion principale.\n\n"
            "Nous vous remercions de vous munir de la présente convocation ainsi que d'une pièce d'identité en cours de validité.\n"
            "En cas d'empêchement, merci de nous prévenir au plus tôt à l'adresse contact habituelle."
        ),
        "signature_titre": "Le Secrétariat Général",
        "pied_page_texte": "Merci de confirmer votre présence dès réception de cette convocation.",
    },
    "recu": {
        "titre": "REÇU DE PAIEMENT",
        "sous_titre": "Justificatif de règlement financier",
        "corps": (
            "Reçu de : {{ civilite }} {{ prenom }} {{ nom }}\n"
            "Adresse électronique : {{ email if email else 'Non communiquée' }}\n\n"
            "La somme forfaitaire de : {{ montant if montant else '0.00' }} €\n"
            "Objet du versement : {{ motif if motif else 'Cotisation / Règlement' }}\n"
            "Date d'encaissement : {{ date }}\n\n"
            "Mode de règlement : Carte bancaire / Virement\n"
            "Numéro de reçu : {{ numero }}\n\n"
            "Le présent reçu est délivré pour servir de justificatif comptable et fiscal."
        ),
        "signature_titre": "Le Service Comptabilité",
        "pied_page_texte": "Reçu délivré conformément aux obligations comptables en vigueur.",
    },
    "bulletin": {
        "titre": "BULLETIN D'ÉVALUATION",
        "sous_titre": "Relevé de résultats et d'appréciations",
        "corps": (
            "Bénéficiaire : {{ civilite }} {{ prenom }} {{ nom }}\n"
            "Intitulé de l'évaluation : {{ motif if motif else 'Évaluation de fin de période' }}\n"
            "Date de l'évaluation : {{ date }}\n\n"
            "Résultats obtenus :\n"
            "- Note finale : {{ note if note else 'N/A' }} / 20\n\n"
            "Appréciation générale de l'équipe pédagogique :\n"
            "Excellent travail et investissement sérieux tout au long de la période.\n"
            "Les compétences requises ont été validées avec succès."
        ),
        "signature_titre": "Le Responsable Pédagogique",
        "pied_page_texte": "Relevé officiel certifié par le jury d'évaluation.",
    },
}


class GestionnaireModeles:
    """Gère le stockage, la modification et le rendu des modèles Jinja2."""

    def __init__(self, dossier_templates: Path = TEMPLATES_DIR):
        self.dossier_templates = Path(dossier_templates)
        self.env = jinja2.Environment(
            loader=jinja2.BaseLoader(),
            autoescape=False,
            trim_blocks=True,
            lstrip_blocks=True,
        )
        self._initialiser_fichiers_modeles()

    def _chemin_modele(self, doc_type: str) -> Path:
        return self.dossier_templates / f"{doc_type.lower()}.json"

    def _initialiser_fichiers_modeles(self) -> None:
        """Crée les fichiers de modèles par défaut s'ils n'existent pas encore."""
        self.dossier_templates.mkdir(parents=True, exist_ok=True)
        for doc_type, contenu in MODELES_PAR_DEFAUT.items():
            fichier = self._chemin_modele(doc_type)
            if not fichier.exists():
                self.sauvegarder_modele(doc_type, contenu)

    def charger_modele(self, doc_type: str) -> Dict[str, str]:
        """Charge le modèle demandé depuis son fichier JSON (ou modèle par défaut)."""
        doc_type_cle = doc_type.lower()
        fichier = self._chemin_modele(doc_type_cle)

        if fichier.exists():
            try:
                with open(fichier, "r", encoding="utf-8") as f:
                    donnees = json.load(f)
                    if isinstance(donnees, dict) and "corps" in donnees:
                        return donnees
            except (json.JSONDecodeError, OSError):
                pass

        return MODELES_PAR_DEFAUT.get(doc_type_cle, MODELES_PAR_DEFAUT["attestation"]).copy()

    def sauvegarder_modele(self, doc_type: str, donnees: Dict[str, str]) -> None:
        """Sauvegarde un modèle modifié dans son fichier JSON."""
        fichier = self._chemin_modele(doc_type)
        with open(fichier, "w", encoding="utf-8") as f:
            json.dump(donnees, f, indent=4, ensure_ascii=False)

    def reinitialiser_modele(self, doc_type: str) -> Dict[str, str]:
        """Rétablit le modèle d'origine pour un type de document."""
        doc_type_cle = doc_type.lower()
        modele_defaut = MODELES_PAR_DEFAUT.get(doc_type_cle, MODELES_PAR_DEFAUT["attestation"]).copy()
        self.sauvegarder_modele(doc_type_cle, modele_defaut)
        return modele_defaut

    def rendre_texte(self, template_str: str, contexte: Dict[str, Any]) -> str:
        """Rend une chaîne Jinja2 avec les variables fournies."""
        try:
            modele_j2 = self.env.from_string(template_str)
            return modele_j2.render(**contexte)
        except Exception as e:
            return f"[Erreur dans le modèle Jinja2 : {e}]\n\n{template_str}"

    def rendre_document(self, doc_type: str, contexte: Dict[str, Any]) -> Dict[str, str]:
        """
        Rend l'ensemble du document (titre, sous-titre, corps, etc.) avec Jinja2.
        """
        modele = self.charger_modele(doc_type)
        return {
            "titre": self.rendre_texte(modele.get("titre", "DOCUMENT"), contexte),
            "sous_titre": self.rendre_texte(modele.get("sous_titre", ""), contexte),
            "corps": self.rendre_texte(modele.get("corps", ""), contexte),
            "signature_titre": self.rendre_texte(modele.get("signature_titre", "Le Responsable"), contexte),
            "pied_page_texte": self.rendre_texte(modele.get("pied_page_texte", ""), contexte),
        }

    @staticmethod
    def obtenir_variables_disponibles() -> list:
        """Retourne la liste des variables Jinja2 utilisables dans les modèles."""
        return [
            ("{{ numero }}", "Numéro d'enregistrement automatique (ex: ATT-2026-0001)"),
            ("{{ date }}", "Date d'émission ou de l'événement (ex: 18/10/2026)"),
            ("{{ civilite }}", "Civilité du destinataire (M. ou Mme)"),
            ("{{ nom }}", "Nom de famille du destinataire"),
            ("{{ prenom }}", "Prénom du destinataire"),
            ("{{ email }}", "Adresse email de contact"),
            ("{{ motif }}", "Motif, intitulé de formation ou objet"),
            ("{{ montant }}", "Montant en euros (pour les reçus)"),
            ("{{ note }}", "Note obtenue sur 20 (pour les bulletins)"),
            ("{{ organisation_nom }}", "Nom de l'organisme émetteur"),
            ("{{ organisation_adresse }}", "Adresse de l'organisme"),
            ("{{ organisation_tel }}", "Téléphone de contact"),
        ]
