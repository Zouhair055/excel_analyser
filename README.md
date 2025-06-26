# 📊 Analyseur Excel - Remplissage Automatique

## 🎯 Description

L'Analyseur Excel est une application web qui permet d'uploader des fichiers Excel et d'appliquer automatiquement des règles de remplissage pour les colonnes vides. L'application détecte les colonnes, applique des règles prédéfinies et permet de télécharger le fichier traité.

## ✨ Fonctionnalités

### 🔍 Détection Automatique
- Analyse automatique des colonnes Excel
- Identification des colonnes vides
- Affichage des informations du fichier (nombre de lignes, colonnes)

### 🤖 Règles de Traitement
L'application applique automatiquement ces règles :

1. **Règle ADVICEPRO** : Si "ADVICEPRO" est détecté dans la colonne 'Description'
   - `Nature` → "G- Suppliers"
   - `Descrip` → "ADVICEPRO"
   - `Vessel` → "N/A"
   - `Service` → "OHD"

2. **Extraction de Références** : Extraction automatique depuis 'Description'
   - Pattern : `AE\d+` (ex: AE1602600010153)
   - Pattern : `OFFICE \d+ \w+` (ex: OFFICE 123 PARIS)

3. **Classification USD** : Si 'Bank account' contient "USD"
   - `Nature` → "Import"

### 🎨 Interface Utilisateur
- Design moderne et responsive
- Zone de drag & drop pour les fichiers
- Barre de progression en temps réel
- Affichage détaillé des résultats
- Téléchargement direct du fichier traité

## 🚀 Installation et Utilisation

### Prérequis
- Python 3.11+
- pip

### Installation
```bash
# Cloner le projet
cd excel-analyzer

# Activer l'environnement virtuel
source venv/bin/activate

# Installer les dépendances
pip install -r requirements.txt
```

### Lancement
```bash
# Démarrer le serveur
python src/main.py

# L'application sera accessible sur http://localhost:5001
```

### Utilisation
1. Ouvrir http://localhost:5001 dans votre navigateur
2. Glisser-déposer votre fichier Excel ou cliquer pour parcourir
3. Attendre le traitement automatique
4. Télécharger le fichier traité

## 🏗️ Architecture Technique

### Backend (Flask)
- **Framework** : Flask avec CORS
- **Traitement** : pandas + openpyxl
- **API REST** : 
  - `POST /api/excel/upload` - Upload et traitement
  - `GET /api/excel/download/<filename>` - Téléchargement
  - `GET /api/excel/columns/<filename>` - Informations colonnes

### Frontend
- **Technologies** : HTML5, CSS3, JavaScript vanilla
- **Design** : Responsive, animations CSS
- **Interactions** : Drag & drop, AJAX

### Structure du Projet
```
excel-analyzer/
├── src/
│   ├── main.py              # Point d'entrée Flask
│   ├── routes/
│   │   ├── excel.py         # API Excel
│   │   └── user.py          # API utilisateur (template)
│   ├── models/              # Modèles de données
│   ├── static/              # Frontend
│   │   ├── index.html       # Interface principale
│   │   ├── style.css        # Styles CSS
│   │   └── script.js        # Logique JavaScript
│   └── database/            # Base de données SQLite
├── venv/                    # Environnement virtuel
├── requirements.txt         # Dépendances Python
├── test_data.xlsx          # Fichier de test
└── README.md               # Documentation
```

## 🧪 Tests

### Tests Automatiques
```bash
# Test de l'API
python test_api.py

# Vérification des résultats
python verify_results.py
```

### Tests Manuels
1. Interface web : http://localhost:5001
2. Upload de fichiers .xlsx et .xls
3. Vérification des règles appliquées
4. Téléchargement des fichiers traités

## 📈 Résultats de Test

✅ **Tests Réussis** :
- Upload de fichier Excel
- Application des 3 règles de traitement
- Extraction correcte des références
- Téléchargement du fichier traité
- Interface utilisateur responsive

## 🔧 Personnalisation

### Ajouter de Nouvelles Règles
Modifier le fichier `src/routes/excel.py`, fonction `apply_rules()` :

```python
def apply_rules(df):
    # Vos nouvelles règles ici
    if 'Nouvelle_Colonne' in df.columns:
        # Logique de traitement
        pass
    
    return df
```

### Modifier l'Interface
- **HTML** : `src/static/index.html`
- **CSS** : `src/static/style.css`
- **JavaScript** : `src/static/script.js`

## 🚀 Déploiement

### Option 1 : Déploiement Local
```bash
python src/main.py
```

### Option 2 : Déploiement Production
L'application est prête pour le déploiement avec des services comme :
- Heroku
- AWS
- Google Cloud
- DigitalOcean

## 📊 Formats Supportés
- **.xlsx** (Excel 2007+)
- **.xls** (Excel 97-2003)

## 🔒 Sécurité
- Validation des types de fichiers
- Limitation de taille (10MB max)
- Nettoyage automatique des fichiers temporaires
- CORS configuré pour les requêtes cross-origin

## 🤝 Contribution
Le projet est structuré pour faciliter l'ajout de nouvelles fonctionnalités :
- Nouvelles règles de traitement
- Support de nouveaux formats
- Améliorations de l'interface
- Optimisations de performance

## 📝 Licence
Développé avec ❤️ pour l'automatisation du traitement Excel.

---

**Note** : Ce projet démontre une approche simple et efficace pour le traitement automatique de fichiers Excel sans nécessiter de modèles d'IA complexes. Les règles sont facilement configurables et extensibles.

