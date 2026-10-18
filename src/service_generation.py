"""
Service orchestrateur de la génération de documents.
Coordonne la numérotation automatique, le rendu des modèles Jinja2,
et la génération des formats PDF (ReportLab) et Word (python-docx).
"""

from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

from src.config import DOC_TYPES, OUTPUT_DIR
from src.docx_generator import GenerateurDOCX
from src.excel_manager import GestionnaireExcel
from src.numerotation import GestionnaireNumerotation
from src.pdf_generator import GenerateurPDF
from src.template_manager import GestionnaireModeles


class ServiceGenerateur:
    """Coordonne la génération unitaire et par lot de documents."""

    def __init__(self, dossier_sortie: Path = OUTPUT_DIR):
        self.dossier_sortie = Path(dossier_sortie)
        self.dossier_sortie.mkdir(parents=True, exist_ok=True)

        self.numerotation = GestionnaireNumerotation()
        self.modeles = GestionnaireModeles()
        self.generateur_pdf = GenerateurPDF(dossier_sortie=self.dossier_sortie)
        self.generateur_docx = GenerateurDOCX(dossier_sortie=self.dossier_sortie)

    def generer_document_unique(
        self,
        donnees: Dict[str, Any],
        format_sortie: str = "both",
        numero_personnalise: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Génère un document pour une personne.
        format_sortie : 'pdf', 'docx' (ou 'word'), ou 'both'.
        Retourne un dictionnaire contenant les chemins des fichiers créés.
        """
        doc_type = donnees.get("type_document", "attestation").lower()
        if doc_type not in DOC_TYPES:
            doc_type = "attestation"

        # 1. Attribution du numéro automatique
        if numero_personnalise:
            numero = numero_personnalise
        else:
            numero = self.numerotation.obtenir_prochain_numero(doc_type=doc_type)

        # 2. Préparation du contexte pour Jinja2
        contexte = dict(donnees)
        contexte["numero"] = numero
        contexte["doc_type"] = doc_type

        # 3. Rendu du modèle Jinja2
        modele_rendu = self.modeles.rendre_document(doc_type=doc_type, contexte=contexte)

        resultat = {
            "numero": numero,
            "destinataire": f"{contexte.get('prenom', '')} {contexte.get('nom', '')}".strip(),
            "doc_type": doc_type,
            "fichiers": {},
        }

        nom_base = f"{numero}_{contexte.get('nom', 'doc')}_{contexte.get('prenom', '')}".strip("_")
        nom_base = "".join(c for c in nom_base if c.isalnum() or c in ("-", "_")).strip()

        # 4. Génération PDF
        if format_sortie.lower() in ("pdf", "both", "tous"):
            nom_pdf = f"{nom_base}.pdf"
            chemin_pdf = self.generateur_pdf.generer(
                modele_rendu=modele_rendu,
                contexte=contexte,
                nom_fichier=nom_pdf,
            )
            resultat["fichiers"]["pdf"] = chemin_pdf

        # 5. Génération Word (.docx)
        if format_sortie.lower() in ("docx", "word", "both", "tous"):
            nom_docx = f"{nom_base}.docx"
            chemin_docx = self.generateur_docx.generer(
                modele_rendu=modele_rendu,
                contexte=contexte,
                nom_fichier=nom_docx,
            )
            resultat["fichiers"]["docx"] = chemin_docx

        return resultat

    def generer_par_lot(
        self,
        chemin_fichier: Path,
        format_sortie: str = "both",
        callback_progression: Optional[Callable[[int, int, str], None]] = None,
    ) -> List[Dict[str, Any]]:
        """
        Lit un fichier Excel ou CSV et génère les documents correspondants pour chaque ligne.
        callback_progression(courant, total, message) : fonction optionnelle pour suivre l'avancement.
        """
        lignes = GestionnaireExcel.lire_fichier(chemin_fichier)
        total = len(lignes)
        resultats = []

        for index, ligne in enumerate(lignes, start=1):
            nom_complet = f"{ligne.get('prenom', '')} {ligne.get('nom', '')}".strip()
            if callback_progression:
                callback_progression(index, total, f"Génération pour : {nom_complet}...")

            res = self.generer_document_unique(
                donnees=ligne,
                format_sortie=format_sortie,
            )
            resultats.append(res)

        if callback_progression:
            callback_progression(total, total, f"Terminé ! {total} document(s) généré(s).")

        return resultats
