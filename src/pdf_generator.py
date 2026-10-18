"""
Générateur de documents PDF utilisant ReportLab.
Produit des documents soignés et professionnels avec en-tête,
numérotation automatique, corps personnalisé par Jinja2 et zone de signature.
"""

from datetime import datetime
from pathlib import Path
from typing import Any, Dict

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm, mm
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    HRFlowable,
    KeepTogether,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
from src.config import ORGANISATION_DEFAULT, OUTPUT_DIR


class NumeroteurDePage(canvas.Canvas):
    """Canvas personnalisé pour ajouter un pied de page avec numérotation 'Page X sur Y'."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.pages = []

    def showPage(self):
        self.pages.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self.pages)
        for page in self.pages:
            self.__dict__.update(page)
            self.dessiner_pied_de_page(num_pages)
            super().showPage()
        super().save()

    def dessiner_pied_de_page(self, total_pages: int):
        self.saveState()
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#718096"))

        # Ligne de séparation au dessus du pied de page
        self.setStrokeColor(colors.HexColor("#E2E8F0"))
        self.setLineWidth(0.8)
        self.line(20 * mm, 18 * mm, 190 * mm, 18 * mm)

        # Texte légal à gauche
        texte_gauche = "Document généré automatiquement - Valide sans rature."
        self.drawString(20 * mm, 13 * mm, texte_gauche)

        # Numérotation à droite
        texte_droite = f"Page {self._pageNumber} / {total_pages}"
        self.drawRightString(190 * mm, 13 * mm, texte_droite)

        self.restoreState()


class GenerateurPDF:
    """Génère des fichiers PDF à partir d'un contexte de données et d'un modèle Jinja2."""

    def __init__(self, dossier_sortie: Path = OUTPUT_DIR):
        self.dossier_sortie = Path(dossier_sortie)
        self.dossier_sortie.mkdir(parents=True, exist_ok=True)
        self.styles = self._creer_styles()

    def _creer_styles(self) -> Dict[str, ParagraphStyle]:
        """Définit la feuille de styles typographiques du document."""
        styles_base = getSampleStyleSheet()

        # Couleurs du thème
        couleur_primaire = colors.HexColor("#1A365D")   # Bleu marine profond
        couleur_secondaire = colors.HexColor("#2B6CB0") # Bleu roi
        couleur_texte = colors.HexColor("#2D3748")      # Gris anthracite lisible
        couleur_muette = colors.HexColor("#718096")     # Gris doux

        styles = {
            "OrgNom": ParagraphStyle(
                "OrgNom",
                parent=styles_base["Normal"],
                fontName="Helvetica-Bold",
                fontSize=11,
                leading=14,
                textColor=couleur_primaire,
            ),
            "OrgDetails": ParagraphStyle(
                "OrgDetails",
                parent=styles_base["Normal"],
                fontName="Helvetica",
                fontSize=8.5,
                leading=11,
                textColor=couleur_muette,
            ),
            "RefBadge": ParagraphStyle(
                "RefBadge",
                parent=styles_base["Normal"],
                fontName="Helvetica-Bold",
                fontSize=9.5,
                leading=13,
                alignment=2,  # Aligné à droite
                textColor=couleur_primaire,
            ),
            "RefSousTitre": ParagraphStyle(
                "RefSousTitre",
                parent=styles_base["Normal"],
                fontName="Helvetica",
                fontSize=8.5,
                leading=11,
                alignment=2,
                textColor=couleur_muette,
            ),
            "TitreDocument": ParagraphStyle(
                "TitreDocument",
                parent=styles_base["Title"],
                fontName="Helvetica-Bold",
                fontSize=18,
                leading=22,
                alignment=1,  # Centré
                textColor=couleur_primaire,
                spaceAfter=4,
            ),
            "SousTitre": ParagraphStyle(
                "SousTitre",
                parent=styles_base["Normal"],
                fontName="Helvetica-Oblique",
                fontSize=10.5,
                leading=14,
                alignment=1,
                textColor=couleur_secondaire,
                spaceAfter=15,
            ),
            "Corps": ParagraphStyle(
                "Corps",
                parent=styles_base["Normal"],
                fontName="Helvetica",
                fontSize=10.5,
                leading=16,
                textColor=couleur_texte,
                spaceAfter=10,
            ),
            "TableEntete": ParagraphStyle(
                "TableEntete",
                parent=styles_base["Normal"],
                fontName="Helvetica-Bold",
                fontSize=9.5,
                leading=12,
                textColor=colors.white,
            ),
            "TableCell": ParagraphStyle(
                "TableCell",
                parent=styles_base["Normal"],
                fontName="Helvetica",
                fontSize=9,
                leading=13,
                textColor=couleur_texte,
            ),
            "TableCellBold": ParagraphStyle(
                "TableCellBold",
                parent=styles_base["Normal"],
                fontName="Helvetica-Bold",
                fontSize=9,
                leading=13,
                textColor=couleur_primaire,
            ),
            "SignatureTitre": ParagraphStyle(
                "SignatureTitre",
                parent=styles_base["Normal"],
                fontName="Helvetica-Bold",
                fontSize=10,
                leading=13,
                alignment=1,
                textColor=couleur_primaire,
            ),
            "SignatureMention": ParagraphStyle(
                "SignatureMention",
                parent=styles_base["Normal"],
                fontName="Helvetica-Oblique",
                fontSize=8.5,
                leading=11,
                alignment=1,
                textColor=couleur_muette,
            ),
        }
        return styles

    def generer(
        self,
        modele_rendu: Dict[str, str],
        contexte: Dict[str, Any],
        nom_fichier: str = None,
    ) -> Path:
        """
        Produit le document PDF sur disque.
        Retourne le chemin absolu du fichier généré.
        """
        numero = contexte.get("numero", "DOC-0001")
        if not nom_fichier:
            nom_propre = f"{contexte.get('nom', 'document')}_{contexte.get('prenom', '')}".strip("_")
            nom_fichier = f"{numero}_{nom_propre}.pdf"

        chemin_pdf = self.dossier_sortie / nom_fichier

        # Création du document
        doc = SimpleDocTemplate(
            str(chemin_pdf),
            pagesize=A4,
            leftMargin=20 * mm,
            rightMargin=20 * mm,
            topMargin=18 * mm,
            bottomMargin=24 * mm,
        )

        histoire = []

        # 1. En-tête (Organisme à gauche, Référence & Date à droite)
        histoire.extend(self._construire_en_tete(contexte))
        histoire.append(Spacer(1, 4 * mm))

        # Ligne décorative
        histoire.append(
            HRFlowable(
                width="100%",
                thickness=1.5,
                color=colors.HexColor("#1A365D"),
                spaceAfter=15,
                spaceBefore=0,
            )
        )

        # 2. Titre et sous-titre du document
        titre = modele_rendu.get("titre", "DOCUMENT OFFICIEL")
        sous_titre = modele_rendu.get("sous_titre", "")
        histoire.append(Paragraph(titre, self.styles["TitreDocument"]))
        if sous_titre:
            histoire.append(Paragraph(sous_titre, self.styles["SousTitre"]))
        else:
            histoire.append(Spacer(1, 4 * mm))

        # 3. Encadré d'identification du destinataire
        histoire.append(self._construire_cadre_destinataire(contexte))
        histoire.append(Spacer(1, 6 * mm))

        # 4. Corps principal (rendu Jinja2)
        corps_texte = modele_rendu.get("corps", "")
        for paragraphe in corps_texte.split("\n\n"):
            para_propre = paragraphe.strip().replace("\n", "<br/>")
            if para_propre:
                histoire.append(Paragraph(para_propre, self.styles["Corps"]))
                histoire.append(Spacer(1, 2 * mm))

        histoire.append(Spacer(1, 4 * mm))

        # 5. Tableau récapitulatif détaillé
        tableau_details = self._construire_tableau_details(contexte)
        if tableau_details:
            histoire.append(tableau_details)
            histoire.append(Spacer(1, 8 * mm))

        # 6. Bloc de signature
        signature_titre = modele_rendu.get("signature_titre", "La Direction")
        histoire.append(self._construire_bloc_signature(contexte, signature_titre))

        # Génération du fichier avec numéroteur de page
        doc.build(histoire, canvasmaker=NumeroteurDePage)
        return chemin_pdf

    def _construire_en_tete(self, contexte: Dict[str, Any]) -> list:
        """Construit le bloc supérieur avec les coordonnées de l'organisme et la référence."""
        org = contexte.get("organisation", ORGANISATION_DEFAULT)
        texte_org = (
            f"<b>{org.get('nom', 'ORGANISME')}</b><br/>"
            f"{org.get('adresse', '')}<br/>"
            f"Tél : {org.get('telephone', '')} | Email : {org.get('email', '')}"
        )

        numero = contexte.get("numero", "DOC-0001")
        date_emission = contexte.get("date", datetime.now().strftime("%d/%m/%Y"))
        texte_ref = (
            f"<b>RÉFÉRENCE : {numero}</b><br/>"
            f"Date d'émission : {date_emission}<br/>"
            f"Statut : Document certifié"
        )

        tableau_en_tete = Table(
            [
                [
                    Paragraph(texte_org, self.styles["OrgDetails"]),
                    Paragraph(texte_ref, self.styles["RefBadge"]),
                ]
            ],
            colWidths=[105 * mm, 65 * mm],
        )
        tableau_en_tete.setStyle(
            TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("ALIGN", (1, 0), (1, 0), "RIGHT"),
                    ("TOPPADDING", (0, 0), (-1, -1), 0),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                ]
            )
        )
        return [tableau_en_tete]

    def _construire_cadre_destinataire(self, contexte: Dict[str, Any]) -> Table:
        """Crée un encadré stylisé pour récapituler les données de la personne concernée."""
        civilite = contexte.get("civilite", "")
        nom = contexte.get("nom", "").upper()
        prenom = contexte.get("prenom", "")
        email = contexte.get("email", "Non renseigné")
        date_ev = contexte.get("date", "")

        lignes = [
            [
                Paragraph("<b>Destinataire concerné :</b>", self.styles["TableCellBold"]),
                Paragraph(f"{civilite} {prenom} {nom}", self.styles["TableCell"]),
                Paragraph("<b>Date concernée :</b>", self.styles["TableCellBold"]),
                Paragraph(str(date_ev), self.styles["TableCell"]),
            ],
            [
                Paragraph("<b>Contact :</b>", self.styles["TableCellBold"]),
                Paragraph(str(email), self.styles["TableCell"]),
                Paragraph("<b>Réf. Dossier :</b>", self.styles["TableCellBold"]),
                Paragraph(str(contexte.get("numero", "")), self.styles["TableCell"]),
            ],
        ]

        cadre = Table(lignes, colWidths=[40 * mm, 50 * mm, 35 * mm, 45 * mm])
        cadre.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F7FAFC")),
                    ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#CBD5E0")),
                    ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                    ("LEFTPADDING", (0, 0), (-1, -1), 8),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ]
            )
        )
        return cadre

    def _construire_tableau_details(self, contexte: Dict[str, Any]) -> Table:
        """Génère un tableau de synthèse pour les éléments quantitatifs (montant, motif, note)."""
        donnees = [
            [
                Paragraph("<b>ÉLÉMENT</b>", self.styles["TableEntete"]),
                Paragraph("<b>DÉSIGNATION / VALEUR</b>", self.styles["TableEntete"]),
            ]
        ]

        if contexte.get("motif"):
            donnees.append(
                [
                    Paragraph("<b>Objet / Motif :</b>", self.styles["TableCellBold"]),
                    Paragraph(str(contexte["motif"]), self.styles["TableCell"]),
                ]
            )

        if contexte.get("montant") and str(contexte["montant"]) not in ("0", "0.00", ""):
            donnees.append(
                [
                    Paragraph("<b>Montant total :</b>", self.styles["TableCellBold"]),
                    Paragraph(f"<b>{contexte['montant']} €</b> (TTC)", self.styles["TableCell"]),
                ]
            )

        if contexte.get("note"):
            donnees.append(
                [
                    Paragraph("<b>Résultat d'évaluation :</b>", self.styles["TableCellBold"]),
                    Paragraph(f"<b>{contexte['note']} / 20</b>", self.styles["TableCell"]),
                ]
            )

        donnees.append(
            [
                Paragraph("<b>Statut de validation :</b>", self.styles["TableCellBold"]),
                Paragraph("Enregistré et certifié conforme", self.styles["TableCell"]),
            ]
        )

        tableau = Table(donnees, colWidths=[55 * mm, 115 * mm])
        tableau.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A365D")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
                    ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#CBD5E0")),
                    ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                    ("LEFTPADDING", (0, 0), (-1, -1), 8),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ]
            )
        )
        return tableau

    def _construire_bloc_signature(self, contexte: Dict[str, Any], signature_titre: str) -> KeepTogether:
        """Construit le bloc de signature situé en bas à droite du document."""
        date_emission = contexte.get("date", datetime.now().strftime("%d/%m/%Y"))

        lignes = [
            [
                "",
                Paragraph(f"Fait le {date_emission}", self.styles["SignatureMention"]),
            ],
            [
                "",
                Paragraph(f"<b>{signature_titre}</b>", self.styles["SignatureTitre"]),
            ],
            [
                "",
                Spacer(1, 15 * mm),  # Emplacement pour la signature manuscrite
            ],
            [
                "",
                Paragraph("[ Cachet et Signature autorisée ]", self.styles["SignatureMention"]),
            ],
        ]

        table_sig = Table(lignes, colWidths=[90 * mm, 80 * mm])
        table_sig.setStyle(
            TableStyle(
                [
                    ("ALIGN", (1, 0), (1, -1), "CENTER"),
                    ("LINEBEFORE", (1, 1), (1, -1), 0.5, colors.HexColor("#CBD5E0")),
                    ("BACKGROUND", (1, 0), (1, -1), colors.HexColor("#F8FAFC")),
                    ("BOX", (1, 0), (1, -1), 0.5, colors.HexColor("#CBD5E0")),
                    ("TOPPADDING", (0, 0), (-1, -1), 3),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ]
            )
        )
        return KeepTogether([table_sig])
