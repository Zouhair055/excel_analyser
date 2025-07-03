#!/usr/bin/env python3
"""
Script de nettoyage du projet Excel Analyzer
Supprime tous les fichiers inutiles et garde seulement le système à règles
"""

import os
import shutil
import glob
from datetime import datetime

def clean_project():
    """Nettoie le projet en supprimant tous les fichiers inutiles"""
    
    print("🧹 NETTOYAGE DU PROJET EXCEL ANALYZER")
    print("=" * 50)
    
    # Fichiers et dossiers à supprimer
    files_to_remove = [
        # Scripts ML obsolètes
        "model_auto_remplissage/model.py",
        "model_auto_remplissage/model_entrainement.py",
        "model_auto_remplissage/Prédicteur_ML.py",
        "model_auto_remplissage/Integration.py",
        "model_auto_remplissage/interface.py",
        "model_auto_remplissage/script.py",
        "model_auto_remplissage/smart_integration.py",
        "model_auto_remplissage/smart_rule_extractor.py",
        "model_auto_remplissage/train_simple.py",
        "model_auto_remplissage/training_pipeline.py",
        
        # Tests obsolètes
        "test_smart_rules.py",
        "training_advanced/train_with_progress.py",
        
        # Dossier script_clone complet (ancien comparateur)
        "script_clone/",
        
        # Fichiers de cache Python
        "model_auto_remplissage/__pycache__/",
        "src/__pycache__/",
        
        # Modèles ML (gardés dans models/ pour référence historique)
        # "model_auto_remplissage/models/*.joblib",
        # "model_auto_remplissage/models/metadata_*.json",
    ]
    
    # Fichiers à garder absolument
    essential_files = [
        "src/main.py",
        "src/routes/excel.py",
        "src/routes/user.py", 
        "src/models/user.py",
        "src/static/index.html",
        "src/static/style.css",
        "src/static/script.js",
        "model_auto_remplissage/train_hybrid_system.py",  # Pour générer de nouvelles règles
        "rules_corrected_*.json",  # Fichiers de règles
        "requirements.txt",
        "README.md",
        "RESUME_EXECUTIF.md",
        "test_api.py",
        "test_rules_detection.py",
        "test_web_app.py",
        "verify_results.py",
        "clean_project.py"  # Ce script
    ]
    
    print("📋 Fichiers/dossiers à supprimer :")
    removed_count = 0
    
    for item in files_to_remove:
        full_path = os.path.abspath(item)
        
        if os.path.exists(full_path):
            try:
                if os.path.isdir(full_path):
                    shutil.rmtree(full_path)
                    print(f"  🗂️  Dossier supprimé: {item}")
                else:
                    os.remove(full_path)
                    print(f"  📄 Fichier supprimé: {item}")
                removed_count += 1
            except Exception as e:
                print(f"  ❌ Erreur lors de la suppression de {item}: {e}")
        else:
            print(f"  ⏭️  Inexistant: {item}")
    
    # Nettoyage des fichiers de cache supplémentaires
    print("\n🧹 Nettoyage des fichiers de cache...")
    cache_patterns = [
        "**/__pycache__/",
        "**/*.pyc",
        "**/*.pyo",
        "**/Thumbs.db",
        "**/.DS_Store"
    ]
    
    for pattern in cache_patterns:
        for file_path in glob.glob(pattern, recursive=True):
            try:
                if os.path.isdir(file_path):
                    shutil.rmtree(file_path)
                else:
                    os.remove(file_path)
                print(f"  🗑️  Cache supprimé: {file_path}")
                removed_count += 1
            except Exception as e:
                print(f"  ❌ Erreur cache: {e}")
    
    print(f"\n✅ Nettoyage terminé: {removed_count} éléments supprimés")
    
    # Vérifier les fichiers essentiels
    print("\n🔍 Vérification des fichiers essentiels...")
    missing_files = []
    
    for pattern in essential_files:
        if "*" in pattern:
            # Pattern avec wildcard
            matches = glob.glob(pattern)
            if not matches:
                missing_files.append(pattern)
        else:
            # Fichier spécifique
            if not os.path.exists(pattern):
                missing_files.append(pattern)
    
    if missing_files:
        print("⚠️  Fichiers essentiels manquants:")
        for file in missing_files:
            print(f"   - {file}")
    else:
        print("✅ Tous les fichiers essentiels sont présents")
    
    return removed_count

def display_final_structure():
    """Affiche la structure finale du projet"""
    
    print("\n" + "=" * 50)
    print("📁 STRUCTURE FINALE DU PROJET")
    print("=" * 50)
    
    # Structure attendue
    expected_structure = {
        "src/": [
            "main.py",
            "routes/excel.py",
            "routes/user.py", 
            "models/user.py",
            "static/index.html",
            "static/style.css",
            "static/script.js",
            "database/app.db"
        ],
        "model_auto_remplissage/": [
            "train_hybrid_system.py",
            "training_data.xlsx"
        ],
        "uploads/": ["(fichiers de test)"],
        "processed/": ["(fichiers traités)"],
        "": [
            "rules_corrected_*.json",
            "requirements.txt",
            "README.md",
            "RESUME_EXECUTIF.md"
        ]
    }
    
    for directory, files in expected_structure.items():
        dir_path = directory if directory else "."
        print(f"\n📂 {dir_path if directory else 'Racine'}")
        
        for file in files:
            if "*" in file:
                # Pattern
                matches = glob.glob(os.path.join(dir_path, file))
                if matches:
                    for match in matches:
                        print(f"   ✅ {os.path.basename(match)}")
                else:
                    print(f"   ❌ {file} (non trouvé)")
            else:
                file_path = os.path.join(dir_path, file)
                if os.path.exists(file_path):
                    print(f"   ✅ {file}")
                else:
                    print(f"   ⚠️  {file}")

if __name__ == "__main__":
    print("🚀 DÉMARRAGE DU NETTOYAGE")
    print()
    
    # Confirmation
    response = input("⚠️  Voulez-vous vraiment nettoyer le projet ? (y/N): ").strip().lower()
    
    if response in ['y', 'yes', 'oui']:
        removed = clean_project()
        display_final_structure()
        
        print("\n" + "=" * 50)
        print("🎉 NETTOYAGE TERMINÉ")
        print(f"📊 {removed} éléments supprimés")
        print("✨ Le projet est maintenant optimisé pour le système à règles uniquement")
        print("\n💡 Prochaines étapes:")
        print("   1. Corriger la détection des règles dans src/routes/excel.py")
        print("   2. Supprimer les imports TensorFlow/ML")
        print("   3. Tester l'application")
        
    else:
        print("❌ Nettoyage annulé")
