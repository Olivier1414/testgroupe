# 📄 Générateur de Documents Automatisés (PDF & Word)

Bienvenue dans le projet **Générateur de Documents Automatisés** !

Cette application en Python a été conçue pour produire facilement des **attestations**, **convocations**, **reçus de paiement** et **bulletins d'évaluation** aux formats **PDF** et **Word (.docx)**, soit individuellement via un formulaire, soit par lot à partir d'un fichier **Excel**.

L'application a été pensée pour être **simple, fonctionnelle et très facile à comprendre**, tout en respectant l'intégralité du cahier des charges.

---

## 🎯 Fonctionnalités principales

1. **Double format de sortie :**
   - 📄 **PDF professionnels** avec mise en page soignée, en-tête d'organisme, numérotation et zone de signature (générés avec **ReportLab**).
   - 📝 **Documents Word (.docx)** entièrement modifiables, stylisés avec tableaux et bordures (générés avec **python-docx**).

2. **Deux modes de saisie des données :**
   - **Formulaire manuel :** Saisissez directement les informations d'un bénéficiaire et générez son document en un clic.
   - **Importation Excel / CSV (par lot) :** Chargez un tableau d'élèves, d'adhérents ou de clients pour générer tous leurs documents d'un coup (publipostage).

3. **Modèles personnalisables (Jinja2) :**
   - Modifiez les textes, formules et mentions légales directement depuis l'interface grâce au moteur **Jinja2** (ex : `{{ prenom }}`, `{{ nom }}`, `{{ date }}`, `{{ motif }}`).
   - Possibilité de rétablir les modèles d'origine à tout moment.

4. **Numérotation automatique et persistante :**
   - Attribution d'un numéro d'enregistrement unique par document (ex : `ATT-2026-0001`, `REC-2026-0002`).
   - Sauvegarde persistante des compteurs dans un fichier `compteur.json`.

5. **Interface graphique conviviale (Tkinter) :**
   - Aucune compétence technique requise : des onglets clairs, des boutons d'ouverture directe des dossiers et une prévisualisation des données.

---

## 📂 Organisation du projet

```text
pjvc/
├── app.py                     # Point d'entrée principal (Interface Graphique Tkinter)
├── cli.py                     # Interface en ligne de commande (alternative sans interface)
├── test_generation.py         # Suite de tests pour valider le bon fonctionnement
├── creer_donnees_test.py      # Script de création des données de démonstration
├── requirements.txt           # Liste des bibliothèques nécessaires
├── README.md                  # Guide complet d'installation et d'utilisation
│
├── data/                      # Données de travail
│   ├── compteur.json          # Fichier de suivi de la numérotation automatique
│   ├── donnees_test.csv       # Données d'exemple au format CSV
│   └── donnees_test.xlsx      # Données d'exemple au format Excel
│
├── templates/                 # Modèles de documents éditables (Jinja2)
│   ├── attestation.json       # Modèle d'attestation de présence / formation
│   ├── convocation.json       # Modèle de convocation officielle
│   ├── recu.json              # Modèle de reçu de paiement
│   └── bulletin.json          # Modèle de bulletin d'évaluation
│
├── output/                    # Dossier où sont enregistrés vos documents générés
│
└── src/                       # Code source modulaire et documenté
    ├── __init__.py
    ├── config.py              # Chemins, types de documents et constantes
    ├── numerotation.py        # Gestion de la numérotation automatique
    ├── template_manager.py    # Chargement et rendu des modèles Jinja2
    ├── excel_manager.py       # Lecture des fichiers Excel (.xlsx) et CSV
    ├── pdf_generator.py       # Moteur de génération PDF (ReportLab)
    ├── docx_generator.py      # Moteur de génération Word (python-docx)
    └── service_generation.py  # Orchestrateur central
```

---

## 🛠️ Installation étape par étape

### 1. Prérequis
- Avoir installé **Python 3.8 ou supérieur** sur votre ordinateur (téléchargeable gratuitement sur [python.org](https://www.python.org/)).
- Lors de l'installation de Python sur Windows, pensez à bien cocher la case **"Add Python to PATH"**.

### 2. Ouvrir un terminal dans le dossier du projet
Ouvrez votre terminal (Invite de commandes, PowerShell ou Terminal VS Code) dans le dossier du projet :
```bash
cd chemin/vers/pjvc
```

### 3. (Optionnel mais recommandé) Créer un environnement virtuel
```bash
python -m venv venv
```
Activez l'environnement :
- **Sous Windows (PowerShell) :**
  ```powershell
  .\venv\Scripts\Activate.ps1
  ```
- **Sous Windows (Invite de commandes cmd) :**
  ```cmd
  venv\Scripts\activate.bat
  ```

### 4. Installer les dépendances
Installez les bibliothèques requises en une seule commande :
```bash
pip install -r requirements.txt
```

Les bibliothèques installées sont :
- `reportlab` : pour la création des PDF
- `python-docx` : pour la création des fichiers Word
- `jinja2` : pour les variables dynamiques dans les textes
- `openpyxl` : pour lire et écrire les fichiers Excel `.xlsx`

---

## 🚀 Comment lancer et utiliser l'application ?

### Méthode 1 : L'interface graphique (recommandée)

Lancez l'application avec la commande :
```bash
python app.py
```

La fenêtre s'ouvre avec 4 onglets :

#### 1. Onglet « Saisie directe (Formulaire) » :
1. Choisissez le type de document dans la liste déroulante (*Attestation*, *Convocation*, *Reçu*, *Bulletin*). Le prochain numéro automatique est affiché en haut à droite.
2. Remplissez les informations du bénéficiaire (Nom, Prénom, Date, Motif, etc.).
3. Cochez les formats souhaités : **PDF**, **Word** ou les deux.
4. Cliquez sur **« ⚡ Générer le document »**.
5. Un message de confirmation s'affiche et vous pouvez cliquer sur **« Ouvrir le fichier »** pour le consulter immédiatement !

#### 2. Onglet « Import Excel / CSV (Par lot) » :
1. Cliquez sur **« 📂 Parcourir... »** pour sélectionner votre fichier Excel ou sur **« Charger le fichier test »** pour utiliser l'exemple inclus.
2. Le tableau affiche immédiatement les personnes et leurs informations.
3. Cochez les formats désirés (**PDF** / **Word**).
4. Cliquez sur **« ⚡ Lancer la génération par lot »**.
5. Une barre de progression et un journal affichent chaque document créé.
6. Cliquez sur le bouton en haut à droite **« 📁 Ouvrir le dossier de sortie »** pour retrouver tous les fichiers dans le dossier `output/`.

#### 3. Onglet « Modèles personnalisables (Jinja2) » :
1. Choisissez le document dont vous souhaitez modifier le texte (ex : *Attestation*).
2. Personnalisez le titre, le sous-titre ou le corps du texte.
3. Pour insérer une variable dynamique, double-cliquez sur l'une des balises disponibles dans le volet de droite (ex : `{{ nom }}`, `{{ date }}`, `{{ motif }}`).
4. Cliquez sur **« 💾 Enregistrer les modifications »**. Tous les prochains documents générés utiliseront votre nouvelle formulation !
5. Si besoin, le bouton **« 🔄 Rétablir le modèle d'origine »** remet le texte par défaut.

#### 4. Onglet « Numérotation & Options » :
- Consultez le compteur actuel pour chaque type de document.
- Réinitialisez les compteurs à 1 si vous souhaitez repartir de zéro.
- Accédez directement au dossier de destination des fichiers.

---

### Méthode 2 : En ligne de commande (CLI)

Si vous préférez exécuter des commandes ou automatiser des tâches sans ouvrir l'interface graphique :

- **Générer une attestation individuelle :**
  ```bash
  python cli.py --type attestation --nom Dupont --prenom Jean --motif "Formation Bureautique" --format both
  ```

- **Générer un reçu de paiement :**
  ```bash
  python cli.py --type recu --nom Bernard --prenom Thomas --montant 150.00 --format pdf
  ```

- **Générer tous les documents d'un fichier Excel en lot :**
  ```bash
  python cli.py --excel data/donnees_test.xlsx --format both
  ```

---

## 🧪 Données de test et vérification

### Données de démonstration fournies
Le dossier `data/` contient déjà des exemples prêts à l'emploi :
- `data/donnees_test.xlsx` : fichier Excel d'exemple contenant plusieurs personnes avec des attestations, convocations, reçus et bulletins.
- `data/donnees_test.csv` : le même exemple au format CSV délimité par des points-virgules.

Pour régénérer les fichiers de test à tout moment :
```bash
python creer_donnees_test.py
```

### Lancer la suite de tests automatisés
Pour vous assurer que toutes les fonctionnalités sont opérationnelles :
```bash
python test_generation.py
```
Ce script teste :
- L'incrémentation de la numérotation automatique.
- Le moteur Jinja2 et le remplacement des variables.
- La génération effective des fichiers PDF (ReportLab).
- La génération effective des fichiers Word (python-docx).
- La lecture du fichier Excel et la génération par lot.

---

## 💡 Structure d'un fichier Excel pour vos propres données

Si vous créez votre propre fichier Excel, créez une feuille avec les colonnes suivantes sur la première ligne :

| type_document | civilite | nom | prenom | email | date | motif | montant | note |
|---|---|---|---|---|---|---|---|---|
| attestation | Mme | Martin | Sophie | s.martin@email.fr | 15/10/2026 | Formation Python | 0.00 | 18 |
| recu | M. | Bernard | Thomas | t.bernard@email.fr | 16/10/2026 | Adhésion annuelle | 120.00 | |
| convocation | Mme | Dubois | Claire | c.dubois@email.fr | 25/11/2026 | Assemblée générale | 0.00 | |
| bulletin | M. | Lefebvre | Lucas | l.lefebvre@email.fr | 18/10/2026 | Évaluation semestre | 0.00 | 17 |

*Note : Les colonnes non pertinentes pour un document (ex : le montant pour une attestation) peuvent rester vides.*

---

## ❓ En cas de problème fréquent

- **Erreur `ModuleNotFoundError: No module named 'reportlab'` (ou `docx`, `jinja2`, `openpyxl`) :**
  Vérifiez que vous avez bien exécuté `pip install -r requirements.txt`.
- **Les accents ne s'affichent pas correctement dans un fichier CSV :**
  Enregistrez votre CSV au format UTF-8 ou utilisez directement un fichier Excel `.xlsx`.
- **Où sont mes fichiers générés ?**
  Tous les fichiers PDF et Word sont enregistrés dans le sous-dossier `output/`. Vous pouvez cliquer sur le bouton **« 📁 Ouvrir le dossier de sortie »** en haut à droite de l'application.

---

Bonne utilisation et bonnes vacances d'apprentissage ! 🎓
"# testgroupe" 
