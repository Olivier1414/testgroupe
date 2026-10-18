"""
Interface en ligne de commande (CLI) pour le générateur de documents.
Permet de générer des documents sans interface graphique (idéal pour les scripts et tests rapides).

Exemples d'utilisation :
    python cli.py --type attestation --nom Dupont --prenom Jean --motif "Formation Python"
    python cli.py --type recu --nom Martin --prenom Sophie --montant 120.00
    python cli.py --excel data/donnees_test.xlsx --format both
"""

import argparse
import sys
from datetime import datetime
from pathlib import Path

# Ajout du dossier racine au chemin d'importation
RACINE = Path(__file__).resolve().parent
if str(RACINE) not in sys.path:
    sys.path.insert(0, str(RACINE))

from src.config import DOC_TYPES, OUTPUT_DIR
from src.service_generation import ServiceGenerateur


def main():
    parser = argparse.ArgumentParser(
        description="Générateur de documents automatisés (PDF & Word)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        "--type",
        choices=list(DOC_TYPES.keys()),
        default="attestation",
        help="Type de document (attestation, convocation, recu, bulletin)",
    )
    parser.add_argument("--nom", default="Dupont", help="Nom du bénéficiaire")
    parser.add_argument("--prenom", default="Jean", help="Prénom du bénéficiaire")
    parser.add_argument("--civilite", default="M.", help="Civilité (M., Mme)")
    parser.add_argument("--email", default="", help="Adresse email")
    parser.add_argument("--date", default=datetime.now().strftime("%d/%m/%Y"), help="Date (JJ/MM/AAAA)")
    parser.add_argument("--motif", default="Session de formation", help="Objet ou motif du document")
    parser.add_argument("--montant", default="0.00", help="Montant en euros (pour les reçus)")
    parser.add_argument("--note", default="15", help="Note obtenue sur 20 (pour les bulletins)")
    parser.add_argument(
        "--format",
        dest="format_sortie",
        choices=["pdf", "docx", "both"],
        default="both",
        help="Format de sortie (pdf, docx, both)",
    )
    parser.add_argument(
        "--excel",
        dest="chemin_excel",
        help="Chemin vers un fichier Excel (.xlsx) ou CSV pour générer les documents par lot",
    )

    args = parser.parse_args()
    service = ServiceGenerateur()

    print("=" * 65)
    print("       GÉNÉRATEUR DE DOCUMENTS AUTOMATISÉS (PDF & WORD)")
    print("=" * 65)

    if args.chemin_excel:
        # Mode par lot
        chemin = Path(args.chemin_excel)
        if not chemin.exists():
            print(f"[ERREUR] Le fichier spécifié est introuvable : {chemin}")
            sys.exit(1)

        print(f"-> Traitement par lot depuis le fichier : {chemin}")
        print(f"-> Format(s) demandé(s) : {args.format_sortie.upper()}")

        def progression(courant, total, message):
            print(f"[{courant}/{total}] {message}")

        resultats = service.generer_par_lot(
            chemin_fichier=chemin,
            format_sortie=args.format_sortie,
            callback_progression=progression,
        )

        print("-" * 65)
        print(f"✓ Génération terminée avec succès : {len(resultats)} document(s) généré(s).")
        print(f"Dossier de sortie : {OUTPUT_DIR}")
    else:
        # Mode unitaire
        donnees = {
            "type_document": args.type,
            "civilite": args.civilite,
            "nom": args.nom,
            "prenom": args.prenom,
            "email": args.email,
            "date": args.date,
            "motif": args.motif,
            "montant": args.montant,
            "note": args.note,
        }

        print(f"-> Type de document : {DOC_TYPES[args.type]['label']}")
        print(f"-> Bénéficiaire     : {args.civilite} {args.prenom} {args.nom}")
        print(f"-> Format           : {args.format_sortie.upper()}")

        res = service.generer_document_unique(
            donnees=donnees,
            format_sortie=args.format_sortie,
        )

        print("-" * 65)
        print(f"✓ Document généré avec succès !")
        print(f"-> Référence automatique : {res['numero']}")
        for fmt, chemin in res["fichiers"].items():
            print(f"-> Fichier {fmt.upper()} : {chemin}")

    print("=" * 65)


if __name__ == "__main__":
    main()
