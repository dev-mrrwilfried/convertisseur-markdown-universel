# Convertisseur Universel de Fichiers vers Markdown

**Version :** Python 3.7+  
**Compatibilité :** Windows, Linux, macOS  
**Fonctionnalités :** Multithread, OCR, Extraction sémantique, Exploration web

---

## 📋 Table des matières

1. [Introduction](#introduction)
2. [Installation](#installation)
3. [Formats supportés](#formats-supportés)
4. [Utilisation de base](#utilisation-de-base)
5. [Options avancées](#options-avancées)
6. [Modes spéciaux](#modes-spéciaux)
7. [Exemples d'utilisation](#exemples-dutilisation)
8. [Dépannage](#dépannage)

---

## 🚀 Introduction

Le **Convertisseur Universel** est un outil Python puissant qui permet de convertir automatiquement des dizaines de formats de fichiers vers le format Markdown. Il intègre des fonctionnalités avancées comme l'OCR pour les images scannées, l'extraction sémantique pour les PDF, et même l'exploration complète de sites web.

### ✨ Caractéristiques principales

- **30+ formats supportés** (PDF, DOCX, XLSX, PPTX, images, HTML, etc.)
- **Traitement multithread** avec barre de progression en temps réel
- **OCR intégré** pour extraire le texte des images et PDF scannés
- **Extraction sémantique** pour une meilleure structure des documents
- **Exploration de sites web** avec filtrage intelligent des contenus
- **Interface simple** en ligne de commande
- **Gestion automatique des erreurs** et des chemins longs

---

## 📦 Installation

### Prérequis
- Python 3.7 ou supérieur
- pip (gestionnaire de paquets Python)

### Installation automatique des dépendances

Utilisez le script fourni pour installer automatiquement toutes les dépendances :

```bash
# Windows
InstallDependance.bat

# Linux/macOS
python convertisseur.py --install-deps
```

### Installation manuelle

Si vous préférez installer manuellement :

```bash
pip install python-docx beautifulsoup4 lxml PyPDF2 pdf2image pytesseract Pillow opencv-python numpy openpyxl xlrd python-pptx striprtf odfpy ebooklib pyyaml pdfminer.six pdfplumber requests
```

### Dépendance système (pour OCR)

Pour utiliser l'OCR, installez Tesseract :
- **Windows** : [Tesseract GitHub Releases](https://github.com/tesseract-ocr/tesseract)
- **Linux** : `sudo apt-get install tesseract-ocr tesseract-ocr-fra`
- **macOS** : `brew install tesseract tesseract-lang`

---

## 📄 Formats supportés

| Catégorie | Extensions | Fonctionnalités spéciales |
|-----------|------------|---------------------------|
| **Documents** | `.docx`, `.doc`, `.odt` | Structure, tableaux |
| **Tableurs** | `.xlsx`, `.xls`, `.ods`, `.csv`, `.tsv` | Feuilles multiples, formats |
| **Présentations** | `.pptx`, `.ppt`, `.odp` | Notes, images (option) |
| **PDF** | `.pdf` | OCR, extraction sémantique |
| **Images** | `.png`, `.jpg`, `.jpeg`, `.bmp`, `.tiff`, `.gif`, `.webp` | OCR automatique |
| **Web** | `.html`, `.htm`, `.xml`, `.xhtml` | Nettoyage du contenu |
| **E-books** | `.epub` | Structure complète |
| **Données** | `.json`, `.yaml`, `.yml` | Formatage préservé |
| **Autres** | `.txt`, `.md`, `.rtf`, `.log`, `.ini`, `.cfg` | Contenu brut |
| **Web en ligne** | `http://`, `https://` | Exploration de sites |

---

## 🎯 Utilisation de base

### Conversion d'un fichier unique

```bash
python convertisseur.py document.pdf
```

**Résultat :** Crée `document.md` dans le même dossier

### Spécifier le fichier de sortie

```bash
python convertisseur.py document.pdf -o sortie.md
```

### Conversion d'un dossier complet

```bash
python convertisseur.py -d /chemin/vers/dossier -o dossier_markdown
```

### Conversion récursive

```bash
python convertisseur.py -d /chemin/vers/dossier -r -o sortie_recursive
```

### Conversion de plusieurs fichiers

```bash
python convertisseur.py file1.pdf file2.docx file3.xlsx --batch -o dossier_sortie
```

---

## ⚙️ Options avancées

### Options de performance

| Option | Description | Exemple |
|--------|-------------|---------|
| `--threads N` | Nombre de threads à utiliser | `--threads 4` |
| `-v, --verbose` | Affichage détaillé des opérations | `-v` |
| `--no-progress` | Désactive la barre de progression | `--no-progress` |

### Options OCR

| Option | Description | Défaut |
|--------|-------------|--------|
| `--ocr-only` | Force l'OCR sur toutes les images | Désactivé |
| `--ocr-lang` | Langues pour l'OCR | `fra+eng` |
| `--no-compress` | Préserve la qualité des images | Désactivé |

**Langues OCR disponibles :** `fra` (français), `eng` (anglais), `spa` (espagnol), `deu` (allemand), etc.

### Options PDF avancées

```bash
# Mode extraction sémantique (meilleure structure)
python convertisseur.py document.pdf --semantic

# Combinaison OCR + sémantique
python convertisseur.py document.pdf --semantic --ocr-only
```

### Options PPTX avancées

```bash
# Extraire les images des présentations
python convertisseur.py presentation.pptx --pptx-images

# Inclure les notes de présentation
python convertisseur.py presentation.pptx --pptx-notes

# Combinaison complète
python convertisseur.py presentation.pptx --pptx-images --pptx-notes
```

---

## 🌐 Modes spéciaux

### 1. Exploration de sites web

Convertit un site web complet en archive Markdown :

```bash
# Exploration basique
python convertisseur.py https://example.com --website

# Exploration approfondie
python convertisseur.py https://example.com --website --depth 3 --max-pages 100 -o archive_site
```

**Paramètres :**
- `--depth N` : Profondeur d'exploration (défaut: 2)
- `--max-pages N` : Nombre maximum de pages (défaut: 50)

**Fonctionnalités intelligentes :**
- Filtrage automatique des commentaires, publicités, et contenus parasites
- Structure hiérarchique préservée
- Génération d'un index de navigation
- Gestion des domaines et sous-domaines

### 2. Mode batch intelligent

```bash
# Traitement de masse avec progression
python convertisseur.py *.pdf *.docx --batch -o resultats --threads 8 -v
```

### 3. Conversion d'URLs individuelles

```bash
# Page web unique
python convertisseur.py https://example.com/article.html -o article.md
```

---

## 📝 Exemples d'utilisation

### Exemple 1 : Rapport d'activité complet

```bash
# Convertir tous les documents d'un projet
python convertisseur.py -d ./documents_projet -r -o markdown_projet --threads 6 -v
```

### Exemple 2 : Archive de site de documentation

```bash
# Télécharger toute la documentation d'un projet
python convertisseur.py https://docs.example.com --website --depth 4 --max-pages 200 -o doc_archive
```

### Exemple 3 : Traitement de PDF scannés

```bash
# PDF avec beaucoup d'images et texte scanné
python convertisseur.py document_scanne.pdf --semantic --ocr-only --ocr-lang fra+eng -v
```

### Exemple 4 : Présentation complète avec médias

```bash
# PPTX avec extraction complète
python convertisseur.py presentation.pptx --pptx-images --pptx-notes -o presentation_complete.md
```

### Exemple 5 : Traitement de masse optimisé

```bash
# 100+ fichiers avec performance maximale
python convertisseur.py -d ./gros_dossier -r --threads 12 --no-compress -o resultats_rapides
```

---

## 📊 Indicateurs de progression

Pendant la conversion, vous verrez des informations en temps réel :

```
🚀 Convertisseur Document - 14:30:25
💻 CPU: 8 cœurs disponibles
────────────────────────────────────────────────────
📁 Analyse: 45 fichier(s) trouvé(s)
📦 Taille totale: 125.8 MB
🧵 Threads: 7
📍 Sortie: ./markdown_output
────────────────────────────────────────────────────
[████████████████████████████▓▓] 85.5% (38/45) | ✅32 ❌6 | 2.3m | ETA: 0.4m | rapport.pdf
```

**Légende :**
- **Barre verte** : Progression globale
- **✅** : Fichiers convertis avec succès
- **❌** : Fichiers en erreur
- **ETA** : Temps estimé restant

---

## 🔧 Dépannage

### Erreurs courantes

#### 1. **"Module not found"**
```bash
# Solution : Installer les dépendances
python convertisseur.py --install-deps
```

#### 2. **"Tesseract not found" (OCR)**
- **Windows** : Vérifiez que Tesseract est dans le PATH
- **Linux** : `sudo apt-get install tesseract-ocr`
- **macOS** : `brew install tesseract`

#### 3. **"Permission denied"**
```bash
# Utilisez des permissions appropriées
sudo python convertisseur.py document.pdf  # Linux/macOS
# Ou exécutez en tant qu'administrateur sur Windows
```

#### 4. **"Path too long" (Windows)**
Le script gère automatiquement les chemins longs en utilisant des hash pour les noms de fichiers.

#### 5. **Mémoire insuffisante**
```bash
# Réduire le nombre de threads
python convertisseur.py document.pdf --threads 2
```

### Conseils d'optimisation

1. **Performance maximale :**
   ```bash
   --threads 8 --no-compress
   ```

2. **Qualité maximale :**
   ```bash
   --semantic --ocr-only --pptx-images --pptx-notes -v
   ```

3. **Pour de très gros fichiers :**
   ```bash
   --threads 1 --verbose  # Évite la surcharge mémoire
   ```

### Diagnostic

Utilisez le mode verbose pour diagnostiquer les problèmes :

```bash
python convertisseur.py document.pdf -v
```

---

## 📋 Structure des fichiers de sortie

### Pour un fichier unique
```
document.pdf → document.md
```

### Pour un dossier
```
documents/
├── doc1.pdf → markdown_output/doc1.md
├── images/pic.jpg → markdown_output/images/pic.md
└── excel/data.xlsx → markdown_output/excel/data.md
```

### Pour un site web
```
archive_site/
├── INDEX.md                    # Index de navigation
├── example.com/
│   ├── index.md               # Page d'accueil
│   ├── about/
│   │   └── index.md          # Page À propos
│   └── documentation/
│       ├── guide.md          # Guide utilisateur
│       └── api.md            # Documentation API
└── short/                     # Fichiers avec chemins trop longs (hash)
    └── a1b2c3d4e5f6.md
```

---

## 🤝 Support et contribution

### Signaler un problème

Si vous rencontrez un bug :

1. Activez le mode verbose : `-v`
2. Notez le message d'erreur exact
3. Indiquez votre système d'exploitation
4. Précisez le type de fichier problématique

### Demandes de fonctionnalités

Le convertisseur est conçu pour être extensible. Les formats peuvent être ajoutés facilement dans la classe `UniversalFileConverter`.

---

## 📜 Licence et crédits

**Intégrations :**
- **Markitdown** : Fonctionnalités PPTX avancées
- **mcp-pdf-reader** : Extraction sémantique PDF
- **Tesseract** : Moteur OCR
- **BeautifulSoup** : Parsing HTML/XML
- **Pillow** : Traitement d'images

---
