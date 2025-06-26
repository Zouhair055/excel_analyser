# 📊 Analyseur Excel - Remplissage Automatique

## 🎯 Description

L'Analyseur Excel est une application web qui permet d'uploader des fichiers Excel et d'appliquer automatiquement des règles de remplissage pour les colonnes vides. L'application détecte les colonnes, applique des règles prédéfinies et permet de télécharger le fichier traité.

## ✨ Fonctionnalités

### 🔍 Détection Automatique
- Analyse automatique des colonnes Excel
- Identification des colonnes vides
- Affichage des informations du fichier (nombre de lignes, colonnes)

### 🤖 Règles de Traitement
L'application applique automatiquement les régles

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
Développé par Zouhair.

---


