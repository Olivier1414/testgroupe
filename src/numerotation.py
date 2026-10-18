"""
Module de gestion de la numérotation automatique des documents.
Gère un compteur persistant dans un fichier JSON pour garantir l'unicité
des numéros de référence générés pour chaque type de document.
"""

import json
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional

from src.config import COMPTEUR_FILE, DOC_TYPES


class GestionnaireNumerotation:
    """Gère l'incrémentation et la persistance des numéros de documents."""

    def __init__(self, fichier_compteur: Path = COMPTEUR_FILE):
        self.fichier_compteur = Path(fichier_compteur)
        self.compteurs: Dict[str, int] = self._charger_compteurs()

    def _charger_compteurs(self) -> Dict[str, int]:
        """Charge les compteurs depuis le fichier JSON, ou initialise des valeurs par défaut."""
        valeurs_par_defaut = {"global": 1}
        for doc_type in DOC_TYPES.keys():
            valeurs_par_defaut[doc_type] = 1

        if self.fichier_compteur.exists():
            try:
                with open(self.fichier_compteur, "r", encoding="utf-8") as f:
                    donnees = json.load(f)
                    if isinstance(donnees, dict):
                        # Fusion avec les valeurs par défaut au cas où de nouveaux types apparaissent
                        for k, v in donnees.items():
                            if isinstance(v, int):
                                valeurs_par_defaut[k] = v
            except (json.JSONDecodeError, OSError):
                pass

        return valeurs_par_defaut

    def _sauvegarder_compteurs(self) -> None:
        """Sauvegarde les compteurs actuels dans le fichier JSON."""
        self.fichier_compteur.parent.mkdir(parents=True, exist_ok=True)
        try:
            with open(self.fichier_compteur, "w", encoding="utf-8") as f:
                json.dump(self.compteurs, f, indent=4, ensure_ascii=False)
        except OSError as e:
            print(f"[Avertissement] Impossible d'enregistrer le compteur : {e}")

    def formater_numero(self, doc_type: str, seq: int, annee: Optional[int] = None) -> str:
        """Formate un numéro séquentiel selon la convention PREFIX-ANNEE-NUMERO."""
        if annee is None:
            annee = datetime.now().year

        doc_info = DOC_TYPES.get(doc_type.lower())
        prefix = doc_info["prefix"] if doc_info else "DOC"
        return f"{prefix}-{annee}-{seq:04d}"

    def obtenir_prochain_numero(self, doc_type: str = "attestation", incrementer: bool = True) -> str:
        """
        Retourne le prochain numéro de document.
        Si incrementer=True, incrémente et sauvegarde le compteur.
        """
        doc_type_cle = doc_type.lower()
        if doc_type_cle not in self.compteurs:
            self.compteurs[doc_type_cle] = 1

        seq_actuel = self.compteurs[doc_type_cle]
        numero_formate = self.formater_numero(doc_type_cle, seq_actuel)

        if incrementer:
            self.compteurs[doc_type_cle] = seq_actuel + 1
            self.compteurs["global"] = self.compteurs.get("global", 1) + 1
            self._sauvegarder_compteurs()

        return numero_formate

    def previsualiser_prochain_numero(self, doc_type: str = "attestation") -> str:
        """Affiche le prochain numéro sans incrémenter le compteur."""
        return self.obtenir_prochain_numero(doc_type, incrementer=False)

    def reinitialiser_compteur(self, doc_type: str = "all", valeur: int = 1) -> None:
        """Réinitialise le compteur d'un type particulier ou de tous les types."""
        if doc_type == "all":
            for k in self.compteurs.keys():
                self.compteurs[k] = valeur
        else:
            self.compteurs[doc_type.lower()] = valeur
        self._sauvegarder_compteurs()

    def definir_valeur_compteur(self, doc_type: str, valeur: int) -> None:
        """Définit manuellement la valeur suivante du compteur."""
        if valeur < 1:
            valeur = 1
        self.compteurs[doc_type.lower()] = valeur
        self._sauvegarder_compteurs()
