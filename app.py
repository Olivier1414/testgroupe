"""
Application Graphique (Tkinter) - Générateur de documents automatisés.
Permet de générer des attestations, convocations, reçus et bulletins en PDF et Word,
à partir d'un formulaire manuel ou d'un fichier Excel.
"""

import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Ajout du dossier racine au chemin d'importation
RACINE = Path(__file__).resolve().parent
if str(RACINE) not in sys.path:
    sys.path.insert(0, str(RACINE))

from src.config import CSV_TEST_FILE, DOC_TYPES, EXCEL_TEST_FILE, OUTPUT_DIR
from src.excel_manager import GestionnaireExcel
from src.numerotation import GestionnaireNumerotation
from src.service_generation import ServiceGenerateur
from src.template_manager import GestionnaireModeles


def ouvrir_dossier_ou_fichier(chemin: Path) -> None:
    """Ouvre un dossier ou un fichier dans le gestionnaire de fichiers du système."""
    chemin_str = str(Path(chemin).resolve())
    try:
        if sys.platform.startswith("win"):
            os.startfile(chemin_str)
        elif sys.platform.startswith("darwin"):
            subprocess.run(["open", chemin_str], check=False)
        else:
            subprocess.run(["xdg-open", chemin_str], check=False)
    except Exception as e:
        messagebox.showwarning("Ouverture", f"Impossible d'ouvrir l'élément : {e}")


class ApplicationGenerateur(tk.Tk):
    """Fenêtre principale de l'application Tkinter."""

    def __init__(self):
        super().__init__()

        self.title("Générateur de Documents Automatisés - PDF & Word")
        self.geometry("980x740")
        self.minsize(880, 640)

        # Services métier
        self.service = ServiceGenerateur()
        self.numerotation = self.service.numerotation
        self.modeles = self.service.modeles

        # Données de l'onglet Excel
        self.fichier_excel_actuel: Path = None
        self.donnees_excel_chargees = []

        # Création des données de test si absentes
        self._initialiser_donnees_demonstration()

        # Configuration du style visuel
        self._configurer_styles()

        # Construction de l'interface
        self._construire_interface()

        # Rafraîchissement initial
        self._actualiser_apercu_numero()

    def _initialiser_donnees_demonstration(self) -> None:
        """S'assure que les fichiers de test sont présents."""
        try:
            if not EXCEL_TEST_FILE.exists() and not CSV_TEST_FILE.exists():
                from creer_donnees_test import main as generer_tests
                generer_tests()
        except Exception:
            pass

    def _configurer_styles(self) -> None:
        """Applique une palette moderne et élégante aux widgets ttk."""
        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self.couleur_primaire = "#1A365D"    # Bleu marine
        self.couleur_secondaire = "#2B6CB0"  # Bleu moyen
        self.couleur_fond = "#F7FAFC"        # Blanc cassé
        self.couleur_texte = "#2D3748"       # Gris foncé

        self.configure(bg="#EDF2F7")

        self.style.configure(".", font=("Segoe UI", 9), foreground=self.couleur_texte)
        self.style.configure("TNotebook", background="#EDF2F7", tabmargins=[2, 5, 2, 0])
        self.style.configure(
            "TNotebook.Tab",
            font=("Segoe UI", 9, "bold"),
            padding=[14, 6],
            background="#CBD5E0",
            foreground="#2D3748",
        )
        self.style.map(
            "TNotebook.Tab",
            background=[("selected", "#1A365D")],
            foreground=[("selected", "#FFFFFF")],
        )

        self.style.configure(
            "Action.TButton",
            font=("Segoe UI", 10, "bold"),
            background=self.couleur_primaire,
            foreground="#FFFFFF",
            padding=[12, 6],
        )
        self.style.map(
            "Action.TButton",
            background=[("active", self.couleur_secondaire), ("pressed", "#152E4D")],
        )

        self.style.configure(
            "Secondaire.TButton",
            font=("Segoe UI", 9),
            background="#E2E8F0",
            padding=[10, 4],
        )

        self.style.configure("Header.TLabel", font=("Segoe UI", 16, "bold"), foreground=self.couleur_primaire)
        self.style.configure("SubHeader.TLabel", font=("Segoe UI", 9), foreground="#718096")
        self.style.configure("Badge.TLabel", font=("Segoe UI", 9, "bold"), foreground=self.couleur_secondaire)
        self.style.configure("TLabelframe.Label", font=("Segoe UI", 9, "bold"), foreground=self.couleur_primaire)

    def _construire_interface(self) -> None:
        """Construit l'en-tête, les onglets et le pied de page."""
        # 1. Bannière supérieure
        banniere = tk.Frame(self, bg="#FFFFFF", height=70, bd=1, relief="solid")
        banniere.pack(side=tk.TOP, fill=tk.X)

        frame_titres = tk.Frame(banniere, bg="#FFFFFF")
        frame_titres.pack(side=tk.LEFT, padx=18, pady=10)

        lbl_titre = ttk.Label(frame_titres, text="Générateur de Documents Automatisés", style="Header.TLabel", background="#FFFFFF")
        lbl_titre.pack(anchor="w")

        lbl_desc = ttk.Label(
            frame_titres,
            text="Attestations, Convocations, Reçus & Bulletins en formats PDF et Word (docx)",
            style="SubHeader.TLabel",
            background="#FFFFFF",
        )
        lbl_desc.pack(anchor="w")

        btn_ouvrir_dossier = ttk.Button(
            banniere,
            text="📁 Ouvrir le dossier de sortie",
            style="Secondaire.TButton",
            command=lambda: ouvrir_dossier_ou_fichier(OUTPUT_DIR),
        )
        btn_ouvrir_dossier.pack(side=tk.RIGHT, padx=18, pady=16)

        # 2. Système d'onglets (Notebook)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=12, pady=10)

        # Création des 4 onglets
        self.tab_formulaire = ttk.Frame(self.notebook, padding=12)
        self.tab_excel = ttk.Frame(self.notebook, padding=12)
        self.tab_modeles = ttk.Frame(self.notebook, padding=12)
        self.tab_parametres = ttk.Frame(self.notebook, padding=12)

        self.notebook.add(self.tab_formulaire, text="  📄 Saisie directe (Formulaire)  ")
        self.notebook.add(self.tab_excel, text="  📊 Import Excel / CSV (Par lot)  ")
        self.notebook.add(self.tab_modeles, text="  ✏️ Modèles personnalisables (Jinja2)  ")
        self.notebook.add(self.tab_parametres, text="  ⚙️ Numérotation & Options  ")

        # Remplissage de chaque onglet
        self._construire_onglet_formulaire()
        self._construire_onglet_excel()
        self._construire_onglet_modeles()
        self._construire_onglet_parametres()

        # 3. Barre de statut inférieure
        barre_statut = tk.Frame(self, bg="#E2E8F0", height=26)
        barre_statut.pack(side=tk.BOTTOM, fill=tk.X)

        self.lbl_statut = tk.Label(
            barre_statut,
            text="Prêt. Choisissez un type de document pour commencer.",
            bg="#E2E8F0",
            fg="#4A5568",
            font=("Segoe UI", 8),
            anchor="w",
        )
        self.lbl_statut.pack(side=tk.LEFT, padx=12, pady=3)

    # =========================================================================
    # ONGLET 1 : FORMULAIRE INDIVIDUEL
    # =========================================================================
    def _construire_onglet_formulaire(self) -> None:
        cadre_principal = ttk.Frame(self.tab_formulaire)
        cadre_principal.pack(fill=tk.BOTH, expand=True)

        # Cadre supérieur : Choix du type de document et aperçu du numéro
        cadre_type = ttk.LabelFrame(cadre_principal, text=" Type de document & Numérotation automatique ", padding=10)
        cadre_type.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(cadre_type, text="Document à générer :", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", padx=6, pady=4)

        self.var_doc_type = tk.StringVar(value="attestation")
        choix_types = [
            ("Attestation", "attestation"),
            ("Convocation", "convocation"),
            ("Reçu de paiement", "recu"),
            ("Bulletin de notes", "bulletin"),
        ]
        self.combo_type = ttk.Combobox(
            cadre_type,
            values=[nom for nom, cle in choix_types],
            state="readonly",
            width=26,
            font=("Segoe UI", 9),
        )
        self.combo_type.current(0)
        self.combo_type.grid(row=0, column=1, sticky="w", padx=6, pady=4)
        self.combo_type.bind("<<ComboboxSelected>>", self._au_changement_type_document)

        self.lbl_apercu_num = ttk.Label(
            cadre_type,
            text="Prochain numéro attribué : ATT-2026-0001",
            style="Badge.TLabel",
        )
        self.lbl_apercu_num.grid(row=0, column=2, sticky="e", padx=20, pady=4)
        cadre_type.grid_columnconfigure(2, weight=1)

        # Cadre central : Champs de formulaire
        cadre_champs = ttk.LabelFrame(cadre_principal, text=" Informations du bénéficiaire / destinataire ", padding=12)
        cadre_champs.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Variables de saisie
        self.var_civilite = tk.StringVar(value="M.")
        self.var_nom = tk.StringVar(value="Dupont")
        self.var_prenom = tk.StringVar(value="Jean")
        self.var_email = tk.StringVar(value="jean.dupont@email.fr")
        self.var_date = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        self.var_motif = tk.StringVar(value="Formation Python & Bureautique")
        self.var_montant = tk.StringVar(value="150.00")
        self.var_note = tk.StringVar(value="17.5")

        # Grille de formulaire
        # Ligne 0 : Civilité, Nom, Prénom
        ttk.Label(cadre_champs, text="Civilité :").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        combo_civ = ttk.Combobox(cadre_champs, textvariable=self.var_civilite, values=["M.", "Mme", "Mlle"], width=8, state="readonly")
        combo_civ.grid(row=0, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(cadre_champs, text="Nom de famille :").grid(row=0, column=2, sticky="w", padx=6, pady=6)
        ent_nom = ttk.Entry(cadre_champs, textvariable=self.var_nom, width=22)
        ent_nom.grid(row=0, column=3, sticky="ew", padx=6, pady=6)

        ttk.Label(cadre_champs, text="Prénom :").grid(row=0, column=4, sticky="w", padx=6, pady=6)
        ent_prenom = ttk.Entry(cadre_champs, textvariable=self.var_prenom, width=22)
        ent_prenom.grid(row=0, column=5, sticky="ew", padx=6, pady=6)

        # Ligne 1 : Email & Date
        ttk.Label(cadre_champs, text="Adresse email :").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        ent_email = ttk.Entry(cadre_champs, textvariable=self.var_email, width=28)
        ent_email.grid(row=1, column=1, columnspan=2, sticky="ew", padx=6, pady=6)

        ttk.Label(cadre_champs, text="Date de référence :").grid(row=1, column=3, sticky="w", padx=6, pady=6)
        ent_date = ttk.Entry(cadre_champs, textvariable=self.var_date, width=16)
        ent_date.grid(row=1, column=4, sticky="w", padx=6, pady=6)

        # Ligne 2 : Motif / Intitulé
        ttk.Label(cadre_champs, text="Objet / Motif :").grid(row=2, column=0, sticky="w", padx=6, pady=6)
        ent_motif = ttk.Entry(cadre_champs, textvariable=self.var_motif)
        ent_motif.grid(row=2, column=1, columnspan=5, sticky="ew", padx=6, pady=6)

        # Ligne 3 : Montant (€) et Note (/20)
        ttk.Label(cadre_champs, text="Montant (€) (si reçu) :").grid(row=3, column=0, sticky="w", padx=6, pady=6)
        ent_montant = ttk.Entry(cadre_champs, textvariable=self.var_montant, width=14)
        ent_montant.grid(row=3, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(cadre_champs, text="Note (/20) (si bulletin) :").grid(row=3, column=2, sticky="w", padx=6, pady=6)
        ent_note = ttk.Entry(cadre_champs, textvariable=self.var_note, width=14)
        ent_note.grid(row=3, column=3, sticky="w", padx=6, pady=6)

        cadre_champs.grid_columnconfigure(3, weight=1)
        cadre_champs.grid_columnconfigure(5, weight=1)

        # Cadre inférieur : Choix du format & Bouton de génération
        cadre_actions = ttk.Frame(cadre_principal)
        cadre_actions.pack(fill=tk.X, pady=(0, 6))

        # Checkboxes format
        cadre_formats = ttk.LabelFrame(cadre_actions, text=" Format(s) de sortie demandé(s) ", padding=8)
        cadre_formats.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        self.var_fmt_pdf = tk.BooleanVar(value=True)
        self.var_fmt_docx = tk.BooleanVar(value=True)

        chk_pdf = ttk.Checkbutton(cadre_formats, text="📄 Fichier PDF (ReportLab)", variable=self.var_fmt_pdf)
        chk_pdf.pack(side=tk.LEFT, padx=12)

        chk_docx = ttk.Checkbutton(cadre_formats, text="📝 Fichier Word (python-docx)", variable=self.var_fmt_docx)
        chk_docx.pack(side=tk.LEFT, padx=12)

        # Bouton principal
        btn_generer = ttk.Button(
            cadre_actions,
            text="⚡ Générer le document",
            style="Action.TButton",
            command=self._generer_document_formulaire,
        )
        btn_generer.pack(side=tk.RIGHT, padx=6, pady=8)

        # Zone d'affichage des résultats récents
        cadre_resultat = ttk.LabelFrame(cadre_principal, text=" Résultat de la dernière génération ", padding=8)
        cadre_resultat.pack(fill=tk.X)

        self.lbl_resultat_info = ttk.Label(
            cadre_resultat,
            text="Aucun document généré pour le moment.",
            foreground="#718096",
        )
        self.lbl_resultat_info.pack(side=tk.LEFT, padx=8, pady=4)

        self.btn_ouvrir_dernier = ttk.Button(
            cadre_resultat,
            text="Ouvrir le fichier",
            style="Secondaire.TButton",
            state=tk.DISABLED,
            command=self._ouvrir_dernier_fichier_genere,
        )
        self.btn_ouvrir_dernier.pack(side=tk.RIGHT, padx=8)

        self.dernier_fichier_cree: Path = None

    def _recuperer_type_selectionne(self) -> str:
        """Traduit le libellé sélectionné de la combobox en clé interne."""
        mapping = {
            "Attestation": "attestation",
            "Convocation": "convocation",
            "Reçu de paiement": "recu",
            "Bulletin de notes": "bulletin",
        }
        return mapping.get(self.combo_type.get(), "attestation")

    def _au_changement_type_document(self, event=None) -> None:
        """Met à jour l'aperçu du numéro quand l'utilisateur change de type."""
        self._actualiser_apercu_numero()

    def _actualiser_apercu_numero(self) -> None:
        """Affiche le prochain numéro prévisionnel."""
        cle = self._recuperer_type_selectionne()
        prochain = self.numerotation.previsualiser_prochain_numero(cle)
        self.lbl_apercu_num.config(text=f"Prochain numéro attribué : {prochain}")

    def _generer_document_formulaire(self) -> None:
        """Génère le document selon les informations du formulaire."""
        nom = self.var_nom.get().strip()
        prenom = self.var_prenom.get().strip()

        if not nom:
            messagebox.showerror("Champ requis", "Veuillez renseigner au moins le nom du bénéficiaire.")
            return

        fmt_pdf = self.var_fmt_pdf.get()
        fmt_docx = self.var_fmt_docx.get()

        if not fmt_pdf and not fmt_docx:
            messagebox.showwarning("Format requis", "Veuillez cocher au moins un format de sortie (PDF ou Word).")
            return

        format_sortie = "both" if (fmt_pdf and fmt_docx) else ("pdf" if fmt_pdf else "docx")
        doc_type = self._recuperer_type_selectionne()

        donnees = {
            "type_document": doc_type,
            "civilite": self.var_civilite.get(),
            "nom": nom,
            "prenom": prenom,
            "email": self.var_email.get().strip(),
            "date": self.var_date.get().strip(),
            "motif": self.var_motif.get().strip(),
            "montant": self.var_montant.get().strip(),
            "note": self.var_note.get().strip(),
        }

        try:
            res = self.service.generer_document_unique(donnees=donnees, format_sortie=format_sortie)
            fichiers = res.get("fichiers", {})
            liste_noms = [Path(p).name for p in fichiers.values()]
            numero = res.get("numero", "")

            # Mémorise le fichier principal à ouvrir
            self.dernier_fichier_cree = next(iter(fichiers.values())) if fichiers else None
            self.btn_ouvrir_dernier.config(state=tk.NORMAL)

            msg = f"✓ Succès ! Référence {numero} générée ({', '.join(liste_noms)})."
            self.lbl_resultat_info.config(text=msg, foreground="#22543D")
            self.lbl_statut.config(text=f"Document {numero} généré avec succès dans le dossier 'output/'.")

            # Actualise le prochain numéro
            self._actualiser_apercu_numero()
            self._actualiser_tableau_compteurs()

            messagebox.showinfo(
                "Document généré avec succès",
                f"Le document [{numero}] a été généré :\n\n"
                + "\n".join(f"• {f}" for f in liste_noms)
                + f"\n\nEmplacement : {OUTPUT_DIR}",
            )
        except Exception as e:
            messagebox.showerror("Erreur de génération", f"Une erreur est survenue :\n{e}")

    def _ouvrir_dernier_fichier_genere(self) -> None:
        """Ouvre le dernier fichier produit."""
        if self.dernier_fichier_cree and Path(self.dernier_fichier_cree).exists():
            ouvrir_dossier_ou_fichier(self.dernier_fichier_cree)
        else:
            ouvrir_dossier_ou_fichier(OUTPUT_DIR)

    # =========================================================================
    # ONGLET 2 : IMPORTATION EXCEL (PAR LOT)
    # =========================================================================
    def _construire_onglet_excel(self) -> None:
        cadre_principal = ttk.Frame(self.tab_excel)
        cadre_principal.pack(fill=tk.BOTH, expand=True)

        # 1. Sélection de fichier
        cadre_fichier = ttk.LabelFrame(cadre_principal, text=" Sélection du fichier de données ", padding=10)
        cadre_fichier.pack(fill=tk.X, pady=(0, 10))

        self.lbl_chemin_excel = ttk.Label(
            cadre_fichier,
            text="Aucun fichier sélectionné. Cliquez sur 'Parcourir...' ou utilisez l'exemple fourni.",
            foreground="#4A5568",
        )
        self.lbl_chemin_excel.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)

        btn_exemple = ttk.Button(
            cadre_fichier,
            text="Charger le fichier test",
            style="Secondaire.TButton",
            command=self._charger_fichier_exemple,
        )
        btn_exemple.pack(side=tk.RIGHT, padx=6)

        btn_parcourir = ttk.Button(
            cadre_fichier,
            text="📂 Parcourir (.xlsx / .csv)...",
            style="Secondaire.TButton",
            command=self._choisir_fichier_excel,
        )
        btn_parcourir.pack(side=tk.RIGHT, padx=6)

        # 2. Aperçu tabulaire des données
        cadre_table = ttk.LabelFrame(cadre_principal, text=" Aperçu des enregistrements détectés ", padding=8)
        cadre_table.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        colonnes = ("type", "civilite", "nom", "prenom", "email", "date", "motif", "montant", "note")
        self.arbre_excel = ttk.Treeview(cadre_table, columns=colonnes, show="headings", height=8)

        titres_colonnes = {
            "type": "Type Document",
            "civilite": "Civilité",
            "nom": "Nom",
            "prenom": "Prénom",
            "email": "Email",
            "date": "Date",
            "motif": "Motif / Objet",
            "montant": "Montant (€)",
            "note": "Note",
        }
        largeurs = {
            "type": 105, "civilite": 60, "nom": 100, "prenom": 100,
            "email": 140, "date": 85, "motif": 170, "montant": 75, "note": 60,
        }

        for col, titre in titres_colonnes.items():
            self.arbre_excel.heading(col, text=titre)
            self.arbre_excel.column(col, width=largeurs.get(col, 100), anchor="w")

        barre_defil_y = ttk.Scrollbar(cadre_table, orient=tk.VERTICAL, command=self.arbre_excel.yview)
        barre_defil_x = ttk.Scrollbar(cadre_table, orient=tk.HORIZONTAL, command=self.arbre_excel.xview)
        self.arbre_excel.configure(yscrollcommand=barre_defil_y.set, xscrollcommand=barre_defil_x.set)

        self.arbre_excel.grid(row=0, column=0, sticky="nsew")
        barre_defil_y.grid(row=0, column=1, sticky="ns")
        barre_defil_x.grid(row=1, column=0, sticky="ew")

        cadre_table.grid_rowconfigure(0, weight=1)
        cadre_table.grid_columnconfigure(0, weight=1)

        # 3. Barre de progression et actions par lot
        cadre_lancement = ttk.Frame(cadre_principal)
        cadre_lancement.pack(fill=tk.X, pady=(0, 6))

        # Choix de format pour le lot
        self.var_lot_pdf = tk.BooleanVar(value=True)
        self.var_lot_docx = tk.BooleanVar(value=True)

        cadre_options_lot = ttk.LabelFrame(cadre_lancement, text=" Formats de sortie ", padding=6)
        cadre_options_lot.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Checkbutton(cadre_options_lot, text="PDF", variable=self.var_lot_pdf).pack(side=tk.LEFT, padx=6)
        ttk.Checkbutton(cadre_options_lot, text="Word (.docx)", variable=self.var_lot_docx).pack(side=tk.LEFT, padx=6)

        self.btn_lancer_lot = ttk.Button(
            cadre_lancement,
            text="⚡ Lancer la génération par lot",
            style="Action.TButton",
            state=tk.DISABLED,
            command=self._lancer_generation_lot,
        )
        self.btn_lancer_lot.pack(side=tk.RIGHT, padx=6)

        # Barre de progression
        self.barre_progression = ttk.Progressbar(cadre_principal, orient=tk.HORIZONTAL, mode="determinate")
        self.barre_progression.pack(fill=tk.X, pady=(0, 6))

        # Journal / Log de génération
        cadre_journal = ttk.LabelFrame(cadre_principal, text=" Journal d'exécution ", padding=6)
        cadre_journal.pack(fill=tk.BOTH, expand=False, pady=(0, 4))

        self.texte_journal = tk.Text(cadre_journal, height=5, font=("Consolas", 8), bg="#FFFFFF", fg="#2D3748")
        self.texte_journal.pack(fill=tk.BOTH, expand=True)

    def _choisir_fichier_excel(self) -> None:
        """Ouvre une boîte de dialogue pour choisir un fichier Excel ou CSV."""
        chemin = filedialog.askopenfilename(
            title="Choisir un fichier de données",
            filetypes=[
                ("Fichiers Excel ou CSV", "*.xlsx;*.csv"),
                ("Fichiers Excel (*.xlsx)", "*.xlsx"),
                ("Fichiers CSV (*.csv)", "*.csv"),
                ("Tous les fichiers", "*.*"),
            ],
        )
        if chemin:
            self._charger_fichier(Path(chemin))

    def _charger_fichier_exemple(self) -> None:
        """Charge automatiquement le fichier de test du projet."""
        if EXCEL_TEST_FILE.exists():
            self._charger_fichier(EXCEL_TEST_FILE)
        elif CSV_TEST_FILE.exists():
            self._charger_fichier(CSV_TEST_FILE)
        else:
            try:
                from creer_donnees_test import main as creer_tests
                creer_tests()
                if EXCEL_TEST_FILE.exists():
                    self._charger_fichier(EXCEL_TEST_FILE)
                else:
                    self._charger_fichier(CSV_TEST_FILE)
            except Exception as e:
                messagebox.showerror("Fichier de test", f"Impossible de générer le fichier de test : {e}")

    def _charger_fichier(self, chemin: Path) -> None:
        """Lit et affiche les données du fichier dans le tableau."""
        try:
            donnees = GestionnaireExcel.lire_fichier(chemin)
            if not donnees:
                messagebox.showwarning("Fichier vide", "Aucune donnée n'a été trouvée dans le fichier sélectionné.")
                return

            self.fichier_excel_actuel = chemin
            self.donnees_excel_chargees = donnees
            self.lbl_chemin_excel.config(text=f"Fichier : {chemin.name} ({len(donnees)} ligne(s))")

            # Vider la table
            for row in self.arbre_excel.get_children():
                self.arbre_excel.delete(row)

            # Remplir la table
            for d in donnees:
                self.arbre_excel.insert(
                    "",
                    tk.END,
                    values=(
                        d.get("type_document", "attestation"),
                        d.get("civilite", ""),
                        d.get("nom", ""),
                        d.get("prenom", ""),
                        d.get("email", ""),
                        d.get("date", ""),
                        d.get("motif", ""),
                        d.get("montant", ""),
                        d.get("note", ""),
                    ),
                )

            self.btn_lancer_lot.config(state=tk.NORMAL)
            self._ajouter_log(f"Fichier chargé avec succès : {chemin.name} ({len(donnees)} entrées)")
            self.lbl_statut.config(text=f"{len(donnees)} ligne(s) prêtes pour la génération par lot.")
        except Exception as e:
            messagebox.showerror("Erreur de lecture", f"Impossible de lire le fichier :\n{e}")

    def _ajouter_log(self, texte: str) -> None:
        """Ajoute un message horodaté dans le journal d'exécution."""
        horodatage = datetime.now().strftime("%H:%M:%S")
        self.texte_journal.insert(tk.END, f"[{horodatage}] {texte}\n")
        self.texte_journal.see(tk.END)

    def _lancer_generation_lot(self) -> None:
        """Exécute la génération de masse pour l'ensemble des enregistrements."""
        if not self.donnees_excel_chargees:
            messagebox.showwarning("Aucune donnée", "Veuillez d'abord charger un fichier de données.")
            return

        fmt_pdf = self.var_lot_pdf.get()
        fmt_docx = self.var_lot_docx.get()

        if not fmt_pdf and not fmt_docx:
            messagebox.showwarning("Format requis", "Veuillez cocher au moins un format de sortie (PDF ou Word).")
            return

        format_sortie = "both" if (fmt_pdf and fmt_docx) else ("pdf" if fmt_pdf else "docx")
        total = len(self.donnees_excel_chargees)

        self.barre_progression["maximum"] = total
        self.barre_progression["value"] = 0
        self._ajouter_log(f"--- Démarrage de la génération par lot ({total} documents) ---")

        succes = 0
        erreurs = 0

        for idx, ligne in enumerate(self.donnees_excel_chargees, start=1):
            nom_complet = f"{ligne.get('prenom', '')} {ligne.get('nom', '')}".strip()
            try:
                res = self.service.generer_document_unique(donnees=ligne, format_sortie=format_sortie)
                numero = res.get("numero", "")
                fichiers = [Path(p).name for p in res.get("fichiers", {}).values()]
                self._ajouter_log(f"[{idx}/{total}] Succès : {numero} -> {nom_complet} ({', '.join(fichiers)})")
                succes += 1
            except Exception as e:
                self._ajouter_log(f"[{idx}/{total}] Erreur pour {nom_complet} : {e}")
                erreurs += 1

            self.barre_progression["value"] = idx
            self.update_idletasks()

        self._ajouter_log(f"--- Fin du traitement : {succes} réussi(s), {erreurs} échec(s) ---")
        self._actualiser_apercu_numero()
        self._actualiser_tableau_compteurs()

        messagebox.showinfo(
            "Génération par lot terminée",
            f"Traitement achevé !\n\n• Documents réussis : {succes}\n• Échecs : {erreurs}\n\n"
            f"Fichiers enregistrés dans :\n{OUTPUT_DIR}",
        )

    # =========================================================================
    # ONGLET 3 : MODÈLES PERSONNALISABLES (JINJA2)
    # =========================================================================
    def _construire_onglet_modeles(self) -> None:
        cadre_principal = ttk.Frame(self.tab_modeles)
        cadre_principal.pack(fill=tk.BOTH, expand=True)

        # 1. Sélecteur de modèle
        cadre_haut = ttk.Frame(cadre_principal)
        cadre_haut.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(cadre_haut, text="Choisir le modèle à personnaliser :", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=6)

        self.combo_modele = ttk.Combobox(
            cadre_haut,
            values=["Attestation", "Convocation", "Reçu", "Bulletin"],
            state="readonly",
            width=20,
        )
        self.combo_modele.current(0)
        self.combo_modele.pack(side=tk.LEFT, padx=6)
        self.combo_modele.bind("<<ComboboxSelected>>", self._charger_modele_dans_editeur)

        btn_reinit = ttk.Button(
            cadre_haut,
            text="🔄 Rétablir le modèle d'origine",
            style="Secondaire.TButton",
            command=self._reinitialiser_modele_courant,
        )
        btn_reinit.pack(side=tk.RIGHT, padx=6)

        btn_sauver = ttk.Button(
            cadre_haut,
            text="💾 Enregistrer les modifications",
            style="Action.TButton",
            command=self._sauvegarder_modele_courant,
        )
        btn_sauver.pack(side=tk.RIGHT, padx=6)

        # 2. Zone d'édition et panneau d'aide Jinja2
        panneau_scinde = ttk.PanedWindow(cadre_principal, orient=tk.HORIZONTAL)
        panneau_scinde.pack(fill=tk.BOTH, expand=True)

        # Côté gauche : Champs du modèle
        cadre_editeur = ttk.Frame(panneau_scinde, padding=6)
        panneau_scinde.add(cadre_editeur, weight=3)

        # Titre du document
        ttk.Label(cadre_editeur, text="Titre principal du document :", font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(0, 2))
        self.ent_modele_titre = ttk.Entry(cadre_editeur, font=("Segoe UI", 10, "bold"))
        self.ent_modele_titre.pack(fill=tk.X, pady=(0, 6))

        # Sous-titre
        ttk.Label(cadre_editeur, text="Sous-titre explicatif :", font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(0, 2))
        self.ent_modele_soustitre = ttk.Entry(cadre_editeur)
        self.ent_modele_soustitre.pack(fill=tk.X, pady=(0, 6))

        # Corps de texte Jinja2
        ttk.Label(cadre_editeur, text="Corps du document (texte avec balises Jinja2 {{ variable }}) :", font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(0, 2))
        self.texte_modele_corps = tk.Text(cadre_editeur, wrap=tk.WORD, font=("Consolas", 10), bg="#FFFFFF", fg="#2D3748")
        self.texte_modele_corps.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        # Signature
        ttk.Label(cadre_editeur, text="Titre du signataire :", font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(0, 2))
        self.ent_modele_signature = ttk.Entry(cadre_editeur)
        self.ent_modele_signature.pack(fill=tk.X)

        # Côté droit : Guide des balises Jinja2 disponibles
        cadre_aide = ttk.LabelFrame(panneau_scinde, text=" Variables Jinja2 disponibles ", padding=8)
        panneau_scinde.add(cadre_aide, weight=2)

        ttk.Label(
            cadre_aide,
            text="Double-cliquez sur une variable pour l'insérer dans le texte :",
            wraplength=250,
            font=("Segoe UI", 8, "italic"),
            foreground="#4A5568",
        ).pack(anchor="w", pady=(0, 6))

        self.liste_variables = tk.Listbox(
            cadre_aide,
            font=("Consolas", 9),
            bg="#F7FAFC",
            selectmode=tk.SINGLE,
            bd=1,
            relief="solid",
        )
        self.liste_variables.pack(fill=tk.BOTH, expand=True, pady=(0, 6))
        self.liste_variables.bind("<Double-Button-1>", self._inserer_variable_au_curseur)

        for balise, description in GestionnaireModeles.obtenir_variables_disponibles():
            self.liste_variables.insert(tk.END, f"{balise}  -  {description}")

        btn_inserer = ttk.Button(
            cadre_aide,
            text="➕ Insérer la variable sélectionnée",
            style="Secondaire.TButton",
            command=self._inserer_variable_au_curseur,
        )
        btn_inserer.pack(fill=tk.X)

        # Chargement initial du modèle affiché
        self._charger_modele_dans_editeur()

    def _recuperer_cle_modele_actuel(self) -> str:
        mapping = {
            "Attestation": "attestation",
            "Convocation": "convocation",
            "Reçu": "recu",
            "Bulletin": "bulletin",
        }
        return mapping.get(self.combo_modele.get(), "attestation")

    def _charger_modele_dans_editeur(self, event=None) -> None:
        """Charge le modèle sélectionné dans les champs de saisie."""
        cle = self._recuperer_cle_modele_actuel()
        modele = self.modeles.charger_modele(cle)

        self.ent_modele_titre.delete(0, tk.END)
        self.ent_modele_titre.insert(0, modele.get("titre", ""))

        self.ent_modele_soustitre.delete(0, tk.END)
        self.ent_modele_soustitre.insert(0, modele.get("sous_titre", ""))

        self.ent_modele_signature.delete(0, tk.END)
        self.ent_modele_signature.insert(0, modele.get("signature_titre", ""))

        self.texte_modele_corps.delete("1.0", tk.END)
        self.texte_modele_corps.insert("1.0", modele.get("corps", ""))

    def _inserer_variable_au_curseur(self, event=None) -> None:
        """Insère la balise Jinja2 sélectionnée là où se trouve le curseur."""
        selection = self.liste_variables.curselection()
        if not selection:
            return
        element = self.liste_variables.get(selection[0])
        # Extrait la balise avant les tirets
        balise = element.split("  -  ")[0].strip()
        self.texte_modele_corps.insert(tk.INSERT, balise)
        self.texte_modele_corps.focus_set()

    def _sauvegarder_modele_courant(self) -> None:
        """Sauvegarde les modifications apportées au modèle."""
        cle = self._recuperer_cle_modele_actuel()
        donnees = {
            "titre": self.ent_modele_titre.get().strip(),
            "sous_titre": self.ent_modele_soustitre.get().strip(),
            "corps": self.texte_modele_corps.get("1.0", tk.END).strip(),
            "signature_titre": self.ent_modele_signature.get().strip(),
            "pied_page_texte": "Document officiel avec numérotation unique vérifiable.",
        }

        try:
            self.modeles.sauvegarder_modele(cle, donnees)
            messagebox.showinfo("Modèle enregistré", f"Le modèle '{cle.upper()}' a été mis à jour avec succès !")
        except Exception as e:
            messagebox.showerror("Erreur de sauvegarde", f"Impossible d'enregistrer le modèle : {e}")

    def _reinitialiser_modele_courant(self) -> None:
        """Rétablit le modèle original."""
        cle = self._recuperer_cle_modele_actuel()
        if messagebox.askyesno("Confirmation", f"Voulez-vous rétablir le modèle '{cle.upper()}' par défaut ?"):
            self.modeles.reinitialiser_modele(cle)
            self._charger_modele_dans_editeur()
            messagebox.showinfo("Rétabli", f"Le modèle '{cle.upper()}' a été réinitialisé.")

    # =========================================================================
    # ONGLET 4 : NUMÉROTATION & PARAMÈTRES
    # =========================================================================
    def _construire_onglet_parametres(self) -> None:
        cadre_principal = ttk.Frame(self.tab_parametres)
        cadre_principal.pack(fill=tk.BOTH, expand=True)

        # 1. Gestion des compteurs de numérotation
        cadre_compteurs = ttk.LabelFrame(cadre_principal, text=" État actuel de la numérotation automatique ", padding=12)
        cadre_compteurs.pack(fill=tk.X, pady=(0, 12))

        ttk.Label(
            cadre_compteurs,
            text=(
                "Chaque document reçoit un identifiant unique sous la forme :\n"
                "  • Attestations : ATT-AAAA-XXXX (ex: ATT-2026-0001)\n"
                "  • Convocations : CONV-AAAA-XXXX (ex: CONV-2026-0001)\n"
                "  • Reçus        : REC-AAAA-XXXX (ex: REC-2026-0001)\n"
                "  • Bulletins    : BULL-AAAA-XXXX (ex: BULL-2026-0001)\n\n"
                "Les compteurs sont conservés de manière permanente dans 'data/compteur.json'."
            ),
            font=("Segoe UI", 9),
            foreground="#2D3748",
        ).pack(anchor="w", pady=(0, 10))

        # Tableau des valeurs actuelles
        colonnes_cpt = ("type", "prefixe", "valeur", "prochain")
        self.arbre_compteurs = ttk.Treeview(cadre_compteurs, columns=colonnes_cpt, show="headings", height=5)
        self.arbre_compteurs.heading("type", text="Type de document")
        self.arbre_compteurs.heading("prefixe", text="Préfixe")
        self.arbre_compteurs.heading("valeur", text="Séquence en cours")
        self.arbre_compteurs.heading("prochain", text="Prochain numéro généré")

        self.arbre_compteurs.column("type", width=160)
        self.arbre_compteurs.column("prefixe", width=90)
        self.arbre_compteurs.column("valeur", width=130)
        self.arbre_compteurs.column("prochain", width=180)
        self.arbre_compteurs.pack(fill=tk.X, pady=(0, 8))

        cadre_btn_cpt = ttk.Frame(cadre_compteurs)
        cadre_btn_cpt.pack(fill=tk.X)

        btn_reinit_tous = ttk.Button(
            cadre_btn_cpt,
            text="🔄 Remettre tous les compteurs à 1",
            style="Secondaire.TButton",
            command=self._reinitialiser_tous_les_compteurs,
        )
        btn_reinit_tous.pack(side=tk.LEFT, padx=6)

        # 2. Dossiers de travail
        cadre_dossiers = ttk.LabelFrame(cadre_principal, text=" Emplacement des fichiers générés ", padding=12)
        cadre_dossiers.pack(fill=tk.X, pady=(0, 12))

        ttk.Label(cadre_dossiers, text=f"Dossier de sortie : {OUTPUT_DIR}", font=("Consolas", 9)).pack(anchor="w", pady=(0, 8))

        btn_ouvrir_rep = ttk.Button(
            cadre_dossiers,
            text="📁 Ouvrir l'explorateur de fichiers dans ce dossier",
            style="Action.TButton",
            command=lambda: ouvrir_dossier_ou_fichier(OUTPUT_DIR),
        )
        btn_ouvrir_rep.pack(anchor="w")

        # 3. Technologies utilisées (mention du cahier des charges)
        cadre_techno = ttk.LabelFrame(cadre_principal, text=" À propos de cette application ", padding=12)
        cadre_techno.pack(fill=tk.X)

        ttk.Label(
            cadre_techno,
            text=(
                "Projet développé conformément au cahier des charges :\n"
                "• Python 3 comme langage principal\n"
                "• ReportLab pour la production des fichiers PDF avec mise en page A4\n"
                "• python-docx pour la génération des fichiers Word (.docx)\n"
                "• Jinja2 pour l'interprétation et la personnalisation des modèles textuels\n"
                "• openpyxl & module csv pour l'exploitation et la lecture de fichiers Excel\n"
                "• Tkinter & ttk pour l'interface utilisateur graphique"
            ),
            font=("Segoe UI", 9),
            foreground="#4A5568",
        ).pack(anchor="w")

        self._actualiser_tableau_compteurs()

    def _actualiser_tableau_compteurs(self) -> None:
        """Met à jour les valeurs affichées dans le tableau des compteurs."""
        for item in self.arbre_compteurs.get_children():
            self.arbre_compteurs.delete(item)

        for cle, info in DOC_TYPES.items():
            valeur = self.numerotation.compteurs.get(cle, 1)
            prochain = self.numerotation.previsualiser_prochain_numero(cle)
            self.arbre_compteurs.insert(
                "",
                tk.END,
                values=(info["label"], info["prefix"], str(valeur), prochain),
            )

    def _reinitialiser_tous_les_compteurs(self) -> None:
        """Réinitialise tous les compteurs à 1."""
        if messagebox.askyesno("Confirmation", "Êtes-vous sûr de vouloir remettre tous les compteurs à 1 ?"):
            self.numerotation.reinitialiser_compteur(doc_type="all", valeur=1)
            self._actualiser_tableau_compteurs()
            self._actualiser_apercu_numero()
            messagebox.showinfo("Compteurs réinitialisés", "Tous les compteurs ont été remis à 1.")


def lancer_application():
    """Point d'entrée pour démarrer l'application Tkinter."""
    app = ApplicationGenerateur()
    app.mainloop()


if __name__ == "__main__":
    lancer_application()
