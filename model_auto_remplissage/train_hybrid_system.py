"""
🎯 ENTRAÎNEMENT RÈGLES INTELLIGENTES À NIVEAUX
Approche adaptative : Description → Bank+CCY → Contexte bancaire
"""
import os
import pandas as pd
import json
from datetime import datetime
from collections import defaultdict, Counter
import re

def load_training_data():
    """Charge les données d'entraînement"""
    possible_paths = [
        "training_data_2024.xlsx",
        "model_auto_remplissage/training_data_2024.xlsx",
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            print(f"📊 Chargement: {path}")
            df = pd.read_excel(path, skiprows=7)
            df = df.dropna(how='all')
            print(f"✅ {len(df)} lignes chargées")
            return df

    print("❌ Fichier training_data_2024.xlsx non trouvé")
    return None

def extract_level1_rules(df):
    """🎯 NIVEAU 1: Règles PARFAITES (100% confiance)"""
    print("🔍 NIVEAU 1: Extraction des règles PARFAITES...")
    
    rules = []
    pattern_groups = defaultdict(list)
    
    # Grouper par patterns de Description
    for idx, row in df.iterrows():
        desc = str(row.get('Description', '')).lower().strip()
        if len(desc) > 3:  # 🔥 SEUIL TRÈS BAS - Plus de règles
            signature = {}
            for col in ['Nature', 'Descrip', 'Vessel', 'Service', 'Reference']:
                if col in df.columns:
                    value = str(row.get(col, '')).strip()
                    if value and value != 'nan' and value != '':
                        signature[col] = value
                    else:
                        signature[col] = ''
            
            pattern_groups[desc].append(signature)
    
    # Analyser chaque pattern avec CONFIANCE PARFAITE
    for desc_pattern, signatures in pattern_groups.items():
        if len(signatures) >= 2:  # 🔥 MINIMUM 2 occurrences seulement
            rule = {
                'pattern': desc_pattern,
                'support': len(signatures),
                'rule_type': 'level1_perfect',
                'level': 1,
                'fixed_columns': {},
                'global_confidence': 0,
                'signature_type': 'description_only'
            }
            
            confidences = []
            
            for col in ['Nature', 'Descrip', 'Vessel', 'Service', 'Reference']:
                values = [sig.get(col, '') for sig in signatures]
                non_empty = [v for v in values if v]
                
                if non_empty:
                    value_counts = Counter(non_empty)
                    most_common, count = value_counts.most_common(1)[0]
                    confidence = count / len(signatures)
                    
                    # 🎯 SEUIL PARFAIT : 100% SEULEMENT
                    if confidence == 1.0:  # 100% parfait
                        rule['fixed_columns'][col] = most_common
                        confidences.append(confidence)
            
            # Valider la règle - Au moins 1 colonne parfaite
            if confidences and len(rule['fixed_columns']) >= 1:
                rule['global_confidence'] = 1.0  # Toujours parfait
                rules.append(rule)
    
    print(f"✅ NIVEAU 1: {len(rules)} règles PARFAITES extraites")
    return rules

def extract_level2_rules(df, level1_patterns):
    """🎯 NIVEAU 2: Règles Bank+CCY+Desc PARFAITES"""
    print("🔍 NIVEAU 2: Extraction des règles Bank+CCY+Desc PARFAITES...")
    
    rules = []
    pattern_groups = defaultdict(list)
    covered_descriptions = {rule['pattern'] for rule in level1_patterns}
    
    # Créer des signatures composites
    for idx, row in df.iterrows():
        bank_account = str(row.get('Bank account', '')).strip()
        ccy = str(row.get('CCY', '')).strip()
        description = str(row.get('Description', '')).lower().strip()
        
        # Ignorer si déjà couvert par NIVEAU 1
        if description in covered_descriptions:
            continue
        
        # Créer signature composite avec TOUS les éléments
        if (bank_account and bank_account != 'nan' and 
            ccy and ccy != 'nan' and 
            description and len(description) > 3):
            
            composite_signature = f"{bank_account}|{ccy}|{description}"
            
            signature_record = {}
            for col in ['Nature', 'Descrip', 'Vessel', 'Service', 'Reference']:
                if col in df.columns:
                    value = str(row.get(col, '')).strip()
                    if value and value != 'nan' and value != '':
                        signature_record[col] = value
                    else:
                        signature_record[col] = ''
            
            pattern_groups[composite_signature].append(signature_record)
    
    # Analyser patterns composites avec CONFIANCE PARFAITE
    for composite_pattern, signatures in pattern_groups.items():
        if len(signatures) >= 2:  # 🔥 MINIMUM 2 occurrences
            rule = {
                'pattern': composite_pattern,
                'support': len(signatures),
                'rule_type': 'level2_perfect',
                'level': 2,
                'fixed_columns': {},
                'global_confidence': 0,
                'signature_type': 'bank_ccy_description'
            }
            
            confidences = []
            
            for col in ['Nature', 'Descrip', 'Vessel', 'Service', 'Reference']:
                values = [sig.get(col, '') for sig in signatures]
                non_empty = [v for v in values if v]
                
                if non_empty:
                    value_counts = Counter(non_empty)
                    most_common, count = value_counts.most_common(1)[0]
                    confidence = count / len(signatures)
                    
                    # 🎯 SEUIL PARFAIT : 100% SEULEMENT
                    if confidence == 1.0:
                        rule['fixed_columns'][col] = most_common
                        confidences.append(confidence)
            
            # Valider la règle - Au moins 1 colonne parfaite
            if confidences and len(rule['fixed_columns']) >= 1:
                rule['global_confidence'] = 1.0
                rules.append(rule)
    
    print(f"✅ NIVEAU 2: {len(rules)} règles Bank+CCY+Desc PARFAITES extraites")
    return rules

def extract_level3_rules(df, level1_patterns, level2_patterns):
    """🎯 NIVEAU 3: Règles Bank+CCY PARFAITES"""
    print("🔍 NIVEAU 3: Extraction des règles Bank+CCY PARFAITES...")
    
    rules = []
    pattern_groups = defaultdict(list)
    
    # Patterns déjà couverts
    covered_descriptions = {rule['pattern'] for rule in level1_patterns}
    covered_composites = {rule['pattern'] for rule in level2_patterns}
    
    # Grouper par Bank+CCY
    for idx, row in df.iterrows():
        bank_account = str(row.get('Bank account', '')).strip()
        ccy = str(row.get('CCY', '')).strip()
        description = str(row.get('Description', '')).lower().strip()
        
        # Vérifier si pas déjà couvert
        already_covered = False
        if description in covered_descriptions:
            already_covered = True
        
        composite_check = f"{bank_account}|{ccy}|{description}"
        if composite_check in covered_composites:
            already_covered = True
        
        if already_covered:
            continue
        
        # Créer signature Bank+CCY
        if (bank_account and bank_account != 'nan' and 
            ccy and ccy != 'nan'):
            
            bank_ccy_signature = f"{bank_account}|{ccy}"
            
            signature_record = {}
            for col in ['Nature', 'Descrip', 'Vessel', 'Service']:
                if col in df.columns:
                    value = str(row.get(col, '')).strip()
                    if value and value != 'nan' and value != '':
                        signature_record[col] = value
                    else:
                        signature_record[col] = ''
            
            pattern_groups[bank_ccy_signature].append(signature_record)
    
    # Analyser patterns Bank+CCY avec CONFIANCE PARFAITE
    for bank_ccy_pattern, signatures in pattern_groups.items():
        if len(signatures) >= 2:  # 🔥 MINIMUM 2 occurrences
            rule = {
                'pattern': bank_ccy_pattern,
                'support': len(signatures),
                'rule_type': 'level3_perfect',
                'level': 3,
                'fixed_columns': {},
                'global_confidence': 0,
                'signature_type': 'bank_ccy_only'
            }
            
            confidences = []
            
            for col in ['Nature', 'Descrip', 'Vessel', 'Service']:
                values = [sig.get(col, '') for sig in signatures]
                non_empty = [v for v in values if v]
                
                if non_empty:
                    value_counts = Counter(non_empty)
                    most_common, count = value_counts.most_common(1)[0]
                    confidence = count / len(signatures)
                    
                    # 🎯 SEUIL PARFAIT : 100% SEULEMENT
                    if confidence == 1.0:
                        rule['fixed_columns'][col] = most_common
                        confidences.append(confidence)
            
            # Valider la règle - Au moins 1 colonne parfaite
            if confidences and len(rule['fixed_columns']) >= 1:
                rule['global_confidence'] = 1.0
                rules.append(rule)
    
    print(f"✅ NIVEAU 3: {len(rules)} règles Bank+CCY PARFAITES extraites")
    return rules

def combine_all_levels(level1_rules, level2_rules, level3_rules):
    """🎯 COMBINER TOUS LES NIVEAUX avec priorité"""
    print("🔗 Combinaison des règles multi-niveaux...")
    
    all_rules = []
    
    # NIVEAU 1 : Priorité maximale
    for rule in level1_rules:
        rule['priority'] = 1
        all_rules.append(rule)
    
    # NIVEAU 2 : Priorité moyenne
    for rule in level2_rules:
        rule['priority'] = 2
        all_rules.append(rule)
    
    # NIVEAU 3 : Priorité faible
    for rule in level3_rules:
        rule['priority'] = 3
        all_rules.append(rule)
    
    # Trier par priorité puis confiance
    all_rules.sort(key=lambda r: (r['priority'], r['global_confidence']), reverse=False)
    
    print(f"🏆 TOTAL: {len(all_rules)} règles multi-niveaux combinées")
    print(f"   📊 Niveau 1 (Description): {len(level1_rules)} règles")
    print(f"   📊 Niveau 2 (Bank+CCY+Desc): {len(level2_rules)} règles")
    print(f"   📊 Niveau 3 (Bank+CCY): {len(level3_rules)} règles")
    
    return all_rules

def save_multilevel_rules(rules, df):
    """Sauvegarde les règles multi-niveaux avec analyse"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"rules_multilevel_{timestamp}.json"
    
    rules_data = {
        'extraction_time': datetime.now().isoformat(),
        'total_rules': len(rules),
        'system_type': 'rules_multilevel',
        'extraction_method': 'adaptive_levels_description_bank_ccy',
        'levels': {
            'level1': len([r for r in rules if r['level'] == 1]),
            'level2': len([r for r in rules if r['level'] == 2]),
            'level3': len([r for r in rules if r['level'] == 3])
        },
        'rules': rules
    }
    
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(rules_data, f, indent=2, ensure_ascii=False)
        
        print(f"\n💾 Règles multi-niveaux sauvegardées: {filename}")
        return filename
        
    except Exception as e:
        print(f"❌ Erreur sauvegarde: {e}")
        return None

def main():
    """Fonction principale MULTI-NIVEAUX"""
    print("🚀 ENTRAÎNEMENT RÈGLES MULTI-NIVEAUX")
    print("=" * 60)
    print("Approche: Description → Bank+CCY+Desc → Bank+CCY")
    print()
    
    # 1. Charger données
    df = load_training_data()
    if df is None:
        return
    
    # Vérifier colonnes nécessaires
    required_cols = ['Bank account', 'CCY', 'Description']
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        print(f"❌ Colonnes manquantes: {missing_cols}")
        return
    
    # 2. Extraire règles par niveaux
    level1_rules = extract_level1_rules(df)
    level2_rules = extract_level2_rules(df, level1_rules)
    level3_rules = extract_level3_rules(df, level1_rules, level2_rules)
    
    # 3. Combiner tous les niveaux
    all_rules = combine_all_levels(level1_rules, level2_rules, level3_rules)
    
    # 4. Sauvegarder
    rules_file = save_multilevel_rules(all_rules, df)
    
    # 5. Résumé détaillé
    print("\n" + "=" * 60)
    print("🎉 EXTRACTION MULTI-NIVEAUX TERMINÉE!")
    print(f"✅ {len(all_rules)} règles totales extraites")
    print(f"💾 Fichier: {rules_file}")
    
    # 6. Analyse par niveau
    print("\n📊 ANALYSE PAR NIVEAU:")
    for level in [1, 2, 3]:
        level_rules = [r for r in all_rules if r['level'] == level]
        if level_rules:
            avg_confidence = sum(r['global_confidence'] for r in level_rules) / len(level_rules)
            avg_support = sum(r['support'] for r in level_rules) / len(level_rules)
            print(f"   🎯 Niveau {level}: {len(level_rules)} règles")
            print(f"      Confiance moyenne: {avg_confidence:.3f}")
            print(f"      Support moyen: {avg_support:.1f}")
    
    # 7. Exemples par niveau
    print("\n📋 EXEMPLES PAR NIVEAU:")
    for level in [1, 2, 3]:
        level_rules = [r for r in all_rules if r['level'] == level]
        if level_rules:
            rule = level_rules[0]
            pattern_preview = rule['pattern'][:50] + "..." if len(rule['pattern']) > 50 else rule['pattern']
            print(f"   📌 Niveau {level}: '{pattern_preview}'")
            print(f"      Remplit: {list(rule['fixed_columns'].keys())}")
    
    print("\n🚀 Système multi-niveaux prêt!")
    print("Test: python src/main.py")

if __name__ == "__main__":
    main()