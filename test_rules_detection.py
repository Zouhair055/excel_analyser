#!/usr/bin/env python3
"""
Script de test pour vérifier la détection des fichiers de règles
"""

import sys
import os
import glob
import json

def test_rules_detection():
    """Test la détection des fichiers de règles"""
    
    print("🧪 Test de détection des fichiers de règles")
    print("=" * 50)
    
    # Patterns de recherche
    patterns = [
        "rules_corrected_*.json",
        "rules_only_*.json", 
        "intelligent_rules_*.json",
        "smart_rules_*.json"
    ]
    
    # Répertoires de recherche
    project_root = os.path.abspath(".")
    search_dirs = [
        project_root,
        os.path.join(project_root, "model_auto_remplissage"),
        os.path.join(project_root, "model_auto_remplissage", "models"),
        os.path.join(project_root, "src"),
        os.path.join(project_root, "src", "routes")
    ]
    
    print(f"📁 Répertoire racine du projet: {project_root}")
    print()
    
    rule_files = []
    
    # Recherche par patterns
    print("🔍 Recherche par patterns:")
    for directory in search_dirs:
        abs_dir = os.path.abspath(directory)
        print(f"   📁 Dans: {abs_dir}")
        
        if os.path.exists(abs_dir):
            for pattern in patterns:
                search_path = os.path.join(abs_dir, pattern)
                found_files = glob.glob(search_path)
                if found_files:
                    print(f"      ✅ Pattern '{pattern}': {found_files}")
                    rule_files.extend(found_files)
                else:
                    print(f"      🔍 Pattern '{pattern}': aucun fichier")
        else:
            print(f"      ⚠️ Répertoire inexistant")
    
    print()
    
    # Recherche récursive
    if not rule_files:
        print("🔍 Recherche récursive:")
        for root, dirs, files in os.walk(project_root):
            for file in files:
                if ((file.startswith("rules_corrected_") or 
                     file.startswith("intelligent_rules_") or 
                     file.startswith("smart_rules_") or
                     file.startswith("rules_only_")) and file.endswith(".json")):
                    full_path = os.path.join(root, file)
                    rule_files.append(full_path)
                    print(f"   ✅ Trouvé: {full_path}")
    
    print()
    print("📋 RÉSULTATS:")
    print(f"   🎯 {len(rule_files)} fichier(s) de règles trouvé(s)")
    
    if rule_files:
        # Trier par date de modification
        rule_files_sorted = sorted(rule_files, key=os.path.getctime, reverse=True)
        latest_file = rule_files_sorted[0]
        
        print(f"   📅 Fichier le plus récent: {latest_file}")
        
        # Test de chargement
        try:
            with open(latest_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            rules = data.get('rules', [])
            print(f"   ✅ {len(rules)} règles chargées avec succès")
            
            if rules:
                # Analyser les types de règles
                rule_types = {}
                for rule in rules:
                    rule_type = rule.get('rule_type', 'unknown')
                    rule_types[rule_type] = rule_types.get(rule_type, 0) + 1
                
                print("   📊 Types de règles:")
                for rule_type, count in rule_types.items():
                    print(f"      - {rule_type}: {count}")
                
                # Montrer un exemple
                first_rule = rules[0]
                pattern = first_rule.get('pattern', '')
                fixed_cols = first_rule.get('fixed_columns', {})
                print(f"   🎯 Exemple de règle:")
                print(f"      Pattern: '{pattern[:50]}...'")
                print(f"      Colonnes fixes: {list(fixed_cols.keys())}")
                
                return True
            else:
                print("   ⚠️ Aucune règle dans le fichier")
                return False
                
        except Exception as e:
            print(f"   ❌ Erreur lors du chargement: {e}")
            return False
    else:
        print("   ❌ Aucun fichier de règles détecté")
        return False

def test_simple_rules_predictor():
    """Test de la classe SimpleRulesPredictor"""
    
    print("\n" + "=" * 50)
    print("🧪 Test de SimpleRulesPredictor")
    print("=" * 50)
    
    # Ajouter le chemin src au PYTHONPATH
    sys.path.insert(0, os.path.join(os.path.abspath("."), "src"))
    
    try:
        from routes.excel import SimpleRulesPredictor
        
        predictor = SimpleRulesPredictor()
        success = predictor.load_rules()
        
        if success:
            print("✅ SimpleRulesPredictor fonctionne correctement")
            return True
        else:
            print("❌ SimpleRulesPredictor n'a pas pu charger les règles")
            return False
            
    except ImportError as e:
        print(f"❌ Impossible d'importer SimpleRulesPredictor: {e}")
        return False
    except Exception as e:
        print(f"❌ Erreur avec SimpleRulesPredictor: {e}")
        return False

if __name__ == "__main__":
    print("🚀 Test du système de détection des règles")
    print()
    
    # Test 1: Détection des fichiers
    success1 = test_rules_detection()
    
    # Test 2: Classe SimpleRulesPredictor
    success2 = test_simple_rules_predictor()
    
    print("\n" + "=" * 50)
    print("📊 RÉSUMÉ FINAL:")
    print(f"   Détection fichiers: {'✅' if success1 else '❌'}")
    print(f"   SimpleRulesPredictor: {'✅' if success2 else '❌'}")
    
    if success1 and success2:
        print("🎉 Système prêt à fonctionner!")
    else:
        print("⚠️ Corrections nécessaires")
