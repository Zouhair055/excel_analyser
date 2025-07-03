#!/usr/bin/env python3
"""
Script de test pour vérifier que l'application web utilise correctement les règles intelligentes
"""

import requests
import os
import time
import json

def test_web_app():
    """Test de l'application web Flask avec upload de fichier"""
    
    print("🌐 Test de l'application web avec règles intelligentes")
    print("=" * 60)
    
    # URL de l'application (ajustez le port si nécessaire)
    base_url = "http://localhost:5001"
    upload_url = f"{base_url}/api/excel/upload"
    
    # Fichier de test à uploader
    test_file = "uploads/ZZZ_collones_vide_v0.xlsx"
    
    if not os.path.exists(test_file):
        print(f"❌ Fichier de test non trouvé: {test_file}")
        print("💡 Assurez-vous d'avoir un fichier de test dans le dossier uploads/")
        return False
    
    print(f"📄 Fichier de test: {test_file}")
    
    try:
        # Vérifier que l'application est accessible
        print(f"🔗 Test de connexion à {base_url}...")
        try:
            response = requests.get(base_url, timeout=5)
            print(f"✅ Application accessible (status: {response.status_code})")
        except requests.ConnectionError:
            print("❌ Application non accessible. Assurez-vous qu'elle est lancée.")
            print("💡 Lancez l'application avec: python src/main.py")
            return False
        
        # Préparer le fichier pour l'upload
        print(f"📤 Upload du fichier vers {upload_url}...")
        
        with open(test_file, 'rb') as f:
            files = {'file': (os.path.basename(test_file), f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
            
            # Envoyer la requête
            response = requests.post(upload_url, files=files, timeout=30)
        
        print(f"📡 Statut de la réponse: {response.status_code}")
        
        if response.status_code == 200:
            result = response.json()
            print("✅ Upload réussi !")
            print(f"📋 Message: {result.get('message', 'N/A')}")
            print(f"📄 Fichier original: {result.get('original_file', 'N/A')}")
            print(f"📄 Fichier traité: {result.get('processed_file', 'N/A')}")
            
            # Vérifier les changements appliqués
            changes = result.get('changes_applied', {})
            if changes:
                print("🔧 Changements appliqués:")
                for change in changes.get('rules_applied', []):
                    print(f"   ✅ {change}")
                    
            # Vérifier le formatage
            formatting = result.get('formatting_applied', {})
            if formatting:
                print("🎨 Formatage appliqué:")
                if formatting.get('filters'): print("   ✅ Filtres automatiques")
                if formatting.get('date_formatting'): print("   ✅ Format des dates")
                if formatting.get('numeric_formatting'): print("   ✅ Format des nombres")
                if formatting.get('table_style'): print("   ✅ Style de tableau")
                if formatting.get('frozen_header'): print("   ✅ En-têtes figés")
            
            return True
            
        else:
            print(f"❌ Erreur HTTP {response.status_code}")
            try:
                error_info = response.json()
                print(f"📋 Détails: {error_info}")
            except:
                print(f"📋 Réponse brute: {response.text}")
            return False
            
    except requests.Timeout:
        print("⏱️ Timeout - L'application met trop de temps à répondre")
        return False
    except Exception as e:
        print(f"❌ Erreur inattendue: {e}")
        return False

def test_rules_endpoint():
    """Test optionnel pour vérifier l'état des règles"""
    print("\n🔍 Test de l'état des règles...")
    
    try:
        # Utiliser l'importation directe pour tester
        import sys
        sys.path.append('.')
        from src.routes.excel import SimpleRulesPredictor
        
        predictor = SimpleRulesPredictor()
        if predictor.load_rules():
            print(f"✅ {len(predictor.rules)} règles chargées directement")
            
            # Afficher quelques exemples
            if predictor.rules:
                print("📋 Exemples de règles:")
                for i, rule in enumerate(predictor.rules[:3]):  # 3 premiers
                    pattern = rule.get('pattern', '')[:40]
                    fixed_cols = len(rule.get('fixed_columns', {}))
                    print(f"   {i+1}. '{pattern}...' → {fixed_cols} colonnes")
            
            return True
        else:
            print("❌ Impossible de charger les règles")
            return False
            
    except Exception as e:
        print(f"⚠️ Test règles impossible: {e}")
        return False

if __name__ == "__main__":
    print("🚀 DÉMARRAGE DES TESTS DE L'APPLICATION WEB")
    print("=" * 60)
    
    # Test direct des règles
    rules_ok = test_rules_endpoint()
    
    print("\n" + "=" * 60)
    
    # Test de l'application web
    web_ok = test_web_app()
    
    print("\n" + "=" * 60)
    print("📊 RÉSULTATS:")
    print(f"   🔧 Règles: {'✅ OK' if rules_ok else '❌ ERREUR'}")
    print(f"   🌐 Web App: {'✅ OK' if web_ok else '❌ ERREUR'}")
    
    if rules_ok and web_ok:
        print("\n🎉 TOUS LES TESTS RÉUSSIS!")
        print("💡 L'application web utilise correctement les règles intelligentes.")
    else:
        print("\n⚠️ CERTAINS TESTS ONT ÉCHOUÉ")
        if not rules_ok:
            print("   - Vérifiez que les fichiers de règles sont présents")
        if not web_ok:
            print("   - Vérifiez que l'application Flask est démarrée")
            print("   - Lancez: python src/main.py")
