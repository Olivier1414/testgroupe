"""
Script de test et de vérification automatisé.
Valide l'ensemble des fonctionnalités demandées par le cahier des charges :
1. Numérotation automatique et incrémentation persistante
2. Personnalisation et rendu des modèles avec Jinja2
3. Génération des fichiers PDF avec ReportLab
4. Génération des fichiers Word (.docx) avec python-docx
5. Importation et traitement des fichiers Excel / CSV

Exécution :
    python test_generation.py
"""

import sys
import unittest
from pathlib import Path

# Ajout du dossier racine au chemin d'importation
RACINE = Path(__file__).resolve().parent
if str(RACINE) not in sys.path:
    sys.path.insert(0, str(RACINE))

from src.config import CSV_TEST_FILE, DOC_TYPES, EXCEL_TEST_FILE, OUTPUT_DIR
from src.docx_generator import GenerateurDOCX
from src.excel_manager import GestionnaireExcel
from src.numerotation import GestionnaireNumerotation
from src.pdf_generator import GenerateurPDF
from src.service_generation import ServiceGenerateur
from src.template_manager import GestionnaireModeles


class TestGenerateurDocuments(unittest.TestCase):
    """Suite de tests unitaires et d'intégration."""

    @classmethod
    def setUpClass(cls):
        cls.dossier_test_output = OUTPUT_DIR / "tests_temp"
        cls.dossier_test_output.mkdir(parents=True, exist_ok=True)
        cls.service = ServiceGenerateur(dossier_sortie=cls.dossier_test_output)

    def test_01_numerotation_automatique(self):
        """Vérifie que la numérotation s'incrémente correctement et respecte le format."""
        num = GestionnaireNumerotation()
        num.reinitialiser_compteur("attestation", valeur=1)

        prochain = num.previsualiser_prochain_numero("attestation")
        self.assertTrue(prochain.startswith("ATT-"), f"Format incorrect : {prochain}")
        self.assertTrue(prochain.endswith("-0001"), f"Le premier numéro doit finir par 0001 : {prochain}")

        genere_1 = num.obtenir_prochain_numero("attestation", incrementer=True)
        self.assertEqual(genere_1, prochain)

        genere_2 = num.obtenir_prochain_numero("attestation", incrementer=True)
        self.assertTrue(genere_2.endswith("-0002"), f"L'incrémentation a échoué : {genere_2}")

    def test_02_rendu_modele_jinja2(self):
        """Vérifie que le moteur Jinja2 remplace bien les balises par les variables."""
        modeles = GestionnaireModeles()
        contexte = {
            "civilite": "Mme",
            "nom": "Curie",
            "prenom": "Marie",
            "email": "marie.curie@sciences.fr",
            "date": "18/10/2026",
            "motif": "Recherche en physique",
            "montant": "500.00",
            "note": "20",
            "numero": "ATT-2026-0099",
        }

        rendu = modeles.rendre_document("attestation", contexte)
        self.assertIn("Marie CURIE", rendu["corps"].upper())
        self.assertIn("Recherche en physique", rendu["corps"])
        self.assertIn("ATT-2026-0099", rendu["corps"])

    def test_03_generation_pdf_reportlab(self):
        """Vérifie la création effective de fichiers PDF avec ReportLab pour chaque type."""
        generateur = GenerateurPDF(dossier_sortie=self.dossier_test_output)
        modeles = GestionnaireModeles()

        for doc_type in DOC_TYPES.keys():
            contexte = {
                "civilite": "M.",
                "nom": f"Test_{doc_type}",
                "prenom": "Utilisateur",
                "email": "test@test.fr",
                "date": "18/10/2026",
                "motif": f"Test automatique {doc_type}",
                "montant": "100.00",
                "note": "16",
                "numero": f"TEST-{doc_type.upper()}-0001",
            }
            modele_rendu = modeles.rendre_document(doc_type, contexte)
            fichier_pdf = generateur.generer(modele_rendu, contexte)

            self.assertTrue(fichier_pdf.exists(), f"Le fichier PDF n'a pas été créé : {fichier_pdf}")
            self.assertGreater(fichier_pdf.stat().st_size, 1000, f"Le fichier PDF semble vide : {fichier_pdf}")

    def test_04_generation_word_docx(self):
        """Vérifie la création effective de fichiers Word (.docx) pour chaque type."""
        generateur = GenerateurDOCX(dossier_sortie=self.dossier_test_output)
        modeles = GestionnaireModeles()

        for doc_type in DOC_TYPES.keys():
            contexte = {
                "civilite": "Mme",
                "nom": f"Test_{doc_type}",
                "prenom": "Utilisatrice",
                "email": "test@test.fr",
                "date": "18/10/2026",
                "motif": f"Test Word {doc_type}",
                "montant": "75.00",
                "note": "15",
                "numero": f"TEST-DOCX-{doc_type.upper()}-0001",
            }
            modele_rendu = modeles.rendre_document(doc_type, contexte)
            fichier_docx = generateur.generer(modele_rendu, contexte)

            self.assertTrue(fichier_docx.exists(), f"Le fichier DOCX n'a pas été créé : {fichier_docx}")
            self.assertGreater(fichier_docx.stat().st_size, 1000, f"Le fichier DOCX semble vide : {fichier_docx}")

    def test_05_lecture_donnees_et_lot(self):
        """Vérifie la lecture des fichiers de test (Excel ou CSV) et la génération par lot."""
        fichier_a_tester = EXCEL_TEST_FILE if EXCEL_TEST_FILE.exists() else CSV_TEST_FILE
        if not fichier_a_tester.exists():
            GestionnaireExcel.generer_fichier_test(CSV_TEST_FILE)
            fichier_a_tester = CSV_TEST_FILE

        donnees = GestionnaireExcel.lire_fichier(fichier_a_tester)
        self.assertGreater(len(donnees), 0, "Aucune donnée lue depuis le fichier de test.")

        resultats = self.service.generer_par_lot(fichier_a_tester, format_sortie="both")
        self.assertEqual(len(resultats), len(donnees))

        for res in resultats:
            self.assertIn("pdf", res["fichiers"])
            self.assertIn("docx", res["fichiers"])
            self.assertTrue(res["fichiers"]["pdf"].exists())
            self.assertTrue(res["fichiers"]["docx"].exists())

    def test_06_interface_graphique_tkinter(self):
        """Vérifie que l'interface Tkinter s'initialise sans erreur de widget, de style ou de police."""
        try:
            import tkinter as tk
            from app import ApplicationGenerateur
            app = ApplicationGenerateur()
            app.update_idletasks()
            app.destroy()
        except tk.TclError as e:
            # En environnement sans affichage graphique (serveur sans écran), on ignore l'absence d'écran
            msg = str(e).lower()
            if "no display" in msg or "display" in msg:
                self.skipTest(f"Pas de serveur d'affichage disponible : {e}")
            else:
                self.fail(f"Erreur Tkinter détectée : {e}")


def main():
    print("=" * 65)
    print("   LANCEMENT DE LA SUITE DE TESTS DU GÉNÉRATEUR DE DOCUMENTS")
    print("=" * 65)

    suite = unittest.TestLoader().loadTestsFromTestCase(TestGenerateurDocuments)
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)

    if result.wasSuccessful():
        print("\n" + "=" * 65)
        print(" [SUCCÈS] Tous les tests ont été validés avec succès !")
        print(" - Numérotation automatique : OK")
        print(" - Modèles Jinja2 personnalisables : OK")
        print(" - Générateur PDF (ReportLab) : OK")
        print(" - Générateur Word (python-docx) : OK")
        print(" - Importation et publipostage : OK")
        print(" - Interface graphique Tkinter : OK")
        print("=" * 65)
        return 0
    else:
        print("\n[ÉCHEC] Des erreurs ont été détectées dans les tests.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
