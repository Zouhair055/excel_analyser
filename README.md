# 📊 Analyseur Excel - Remplissage Automatique

## 🎯 Description

Application web simplifiée pour l'upload et le traitement automatique de fichiers Excel avec remplissage intelligent des colonnes vides.

## ✨ Fonctionnalités

### 🔍 Traitement Automatique
- Upload de fichiers Excel (.xlsx, .xls)
- Nettoyage et standardisation des colonnes
- Application de règles de remplissage intelligentes
- Formatage professionnel avec formatage conditionnel

### 🎨 Interface Utilisateur
- Zone de drag & drop
- Barre de progression
- Téléchargement direct du fichier traité
- Statistiques de traitement

## 🚀 Installation et Utilisation

### Prérequis
- Python 3.11+
- pip

### Installation
```bash
cd excel-analyzer
source venv/bin/activate  # ou venv\Scripts\activate sur Windows
pip install -r requirements.txt
```

### Lancement
```bash
python src/main.py
# Accessible sur http://localhost:5001
```

## 🏗️ Architecture Simplifiée

### Backend (Flask)
- **Framework** : Flask avec CORS
- **Traitement** : pandas + openpyxl
- **API REST** : 
  - `POST /api/excel/upload` - Upload et traitement
  - `GET /api/excel/download/<filename>` - Téléchargement
  - `GET /api/excel/health` - État du service

### Structure du Projet
```
excel-analyzer/
├── src/
│   ├── main.py              # Point d'entrée Flask (simplifié)
│   ├── routes/
│   │   ├── excel_clean.py   # API Excel (optimisé ~400 lignes)
│   │   └── user.py          # API utilisateur (complété)
│   ├── models/              # Modèles de données
│   └── static/              # Frontend
├── uploads/                 # Fichiers uploadés
├── processed/               # Fichiers traités
└── rules_*.json            # Fichiers de règles
```

## 📊 Améliorations Apportées

- **Réduction drastique du code** : de 1800 à ~400 lignes
- **Suppression des fonctions redondantes**
- **Simplification du système de règles**
- **Conservation des fonctionnalités essentielles**
- **Amélioration de la lisibilité**

## 🔒 Sécurité
- Protection des colonnes financières
- Validation des types de fichiers
- Gestion d'erreurs robuste

---
Développé par Zouhair - Version optimisée
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


