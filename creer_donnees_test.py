"""
Script utilitaire pour générer les fichiers de test (Excel .xlsx et CSV).
Peut être exécuté directement :
    python creer_donnees_test.py
"""

import csv
import io
import sys
import zipfile
from datetime import datetime
from pathlib import Path

# Ajout du dossier racine au chemin d'importation
RACINE = Path(__file__).resolve().parent
if str(RACINE) not in sys.path:
    sys.path.insert(0, str(RACINE))

from src.config import CSV_TEST_FILE, EXCEL_TEST_FILE
from src.excel_manager import COLONNES_ATTENDUES, GestionnaireExcel


def creer_xlsx_autonome(chemin_destination: Path, colonnes: list, donnees: list) -> None:
    """
    Crée un fichier .xlsx 100% valide en utilisant uniquement la bibliothèque standard (zipfile).
    Fonctionne même sans openpyxl !
    """
    chemin_destination = Path(chemin_destination)
    chemin_destination.parent.mkdir(parents=True, exist_ok=True)

    def col_letter(n):
        res = ""
        while n > 0:
            n, rem = divmod(n - 1, 26)
            res = chr(65 + rem) + res
        return res

    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"""

    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

    wb_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

    workbook_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Documents" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""

    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>"""

    # Construction des lignes de la feuille
    sheet_rows = []
    # Ligne 1 : En-têtes
    r1_cells = []
    for c_idx, col in enumerate(colonnes, start=1):
        ref = f"{col_letter(c_idx)}1"
        val = str(col).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        r1_cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{val}</t></is></c>')
    sheet_rows.append(f'<row r="1">{"".join(r1_cells)}</row>')

    # Lignes 2+ : Données
    for r_idx, d in enumerate(donnees, start=2):
        row_cells = []
        for c_idx, col in enumerate(colonnes, start=1):
            ref = f"{col_letter(c_idx)}{r_idx}"
            val = str(d.get(col, "")).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            row_cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{val}</t></is></c>')
        sheet_rows.append(f'<row r="{r_idx}">{"".join(row_cells)}</row>')

    sheet_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    {"".join(sheet_rows)}
  </sheetData>
</worksheet>"""

    with zipfile.ZipFile(chemin_destination, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/styles.xml", styles_xml)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)


def main():
    print("=" * 60)
    print("   Génération des fichiers de démonstration")
    print("=" * 60)

    donnees = GestionnaireExcel.creer_donnees_exemples()

    # 1. Génération du fichier Excel (.xlsx)
    try:
        try:
            import openpyxl
            chemin_excel = GestionnaireExcel.generer_fichier_test(EXCEL_TEST_FILE)
            print(f"[OK] Fichier Excel généré via openpyxl : {chemin_excel}")
        except ImportError:
            creer_xlsx_autonome(EXCEL_TEST_FILE, COLONNES_ATTENDUES, donnees)
            print(f"[OK] Fichier Excel généré en mode autonome : {EXCEL_TEST_FILE}")
    except Exception as e:
        print(f"[Avertissement] Génération Excel : {e}")

    # 2. Génération du fichier CSV
    try:
        chemin_csv = GestionnaireExcel.generer_fichier_test(CSV_TEST_FILE)
        print(f"[OK] Fichier CSV généré   : {chemin_csv}")
    except Exception as e:
        print(f"[Avertissement] Génération CSV   : {e}")

    print("\nLes fichiers de test sont prêts dans le dossier 'data/' !")


if __name__ == "__main__":
    main()
