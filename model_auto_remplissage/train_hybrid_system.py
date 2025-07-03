"""
🎯 ENTRAÎNEMENT RÈGLES INTELLIGENTES CORRIGÉ
Prend en compte TOUTES les valeurs, même les constantes
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
        "training_data.xlsx",
        "model_auto_remplissage/training_data.xlsx",
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            print(f"📊 Chargement: {path}")
            df = pd.read_excel(path, skiprows=7)
            df = df.dropna(how='all')
            print(f"✅ {len(df)} lignes chargées")
            return df
    
    print("❌ Fichier training_data.xlsx non trouvé")
    return None

def extract_intelligent_rules_corrected(df):
    """Extrait les règles intelligentes CORRIGÉES"""
    print("🔍 Extraction des règles intelligentes corrigées...")
    
    rules = []
    pattern_groups = defaultdict(list)
    
    # Grouper par patterns de Description
    for idx, row in df.iterrows():
        desc = str(row.get('Description', '')).lower().strip()
        if len(desc) > 5:  # Descriptions significatives
            signature = {}
            for col in ['Nature', 'Descrip', 'Vessel', 'Service', 'Reference']:
                if col in df.columns:
                    value = str(row.get(col, '')).strip()
                    # ✅ CORRECTION: Prendre toutes les valeurs, même "N/A"
                    if value and value != 'nan' and value != '':
                        signature[col] = value
                    else:
                        signature[col] = ''  # Valeur vide
            
            pattern_groups[desc].append(signature)
    
    # Analyser chaque pattern
    for desc_pattern, signatures in pattern_groups.items():
        if len(signatures) >= 3:  # Minimum 3 occurrences
            rule = {
                'pattern': desc_pattern,
                'support': len(signatures),
                'rule_type': 'complete_fill',
                'fixed_columns': {},
                'variable_columns': {},
                'global_confidence': 0
            }
            
            confidences = []
            
            for col in ['Nature', 'Descrip', 'Vessel', 'Service', 'Reference']:
                values = [sig.get(col, '') for sig in signatures]
                # ✅ CORRECTION: Inclure toutes les valeurs non vides (même "N/A")
                non_empty = [v for v in values if v]
                
                if non_empty:
                    value_counts = Counter(non_empty)
                    most_common, count = value_counts.most_common(1)[0]
                    confidence = count / len(non_empty)
                    
                    # ✅ CORRECTION: Seuils ajustés
                    if confidence >= 0.85:  # Haute confiance -> colonne fixe
                        rule['fixed_columns'][col] = most_common
                        confidences.append(confidence)
                    elif confidence >= 0.65:  # Confiance modérée -> colonne variable
                        rule['variable_columns'][col] = {
                            'default_value': most_common,
                            'confidence': confidence
                        }
                        confidences.append(confidence)
            
            # Calculer confiance globale
            if confidences:
                rule['global_confidence'] = sum(confidences) / len(confidences)
                
                # Déterminer le type de règle
                if len(rule['fixed_columns']) >= 2:  # Au moins 2 colonnes fixes
                    rule['rule_type'] = 'complete_fill'
                elif len(rule['fixed_columns']) >= 1:
                    rule['rule_type'] = 'hybrid_fill'
                else:
                    rule['rule_type'] = 'conditional_fill'
                
                rules.append(rule)
    
    # Trier par confiance et support
    rules.sort(key=lambda r: (r['global_confidence'], r['support']), reverse=True)
    
    print(f"✅ {len(rules)} règles intelligentes extraites")
    return rules

def analyze_pattern_details(df, pattern, max_examples=10):
    """Analyse détaillée d'un pattern pour debug"""
    print(f"\n🔍 Analyse détaillée du pattern: '{pattern}'")
    
    # Trouver toutes les lignes avec ce pattern
    mask = df['Description'].str.lower().str.contains(re.escape(pattern.lower()), na=False)
    matching_rows = df[mask]
    
    print(f"📊 {len(matching_rows)} lignes trouvées")
    
    if len(matching_rows) > 0:
        print("📋 Exemples:")
        for idx, (_, row) in enumerate(matching_rows.head(max_examples).iterrows()):
            print(f"  {idx+1}. Nature='{row.get('Nature', '')}', Descrip='{row.get('Descrip', '')}', Vessel='{row.get('Vessel', '')}', Service='{row.get('Service', '')}', Reference='{row.get('Reference', '')}'")
        
        # Analyser la distribution des valeurs
        print("\n📊 Distribution des valeurs:")
        for col in ['Nature', 'Descrip', 'Vessel', 'Service', 'Reference']:
            if col in matching_rows.columns:
                values = matching_rows[col].fillna('').astype(str)
                value_counts = values.value_counts()
                print(f"  {col}:")
                for value, count in value_counts.head(3).items():
                    percentage = (count / len(matching_rows)) * 100
                    print(f"    '{value}': {count} fois ({percentage:.1f}%)")

def save_rules_with_analysis(rules, df):
    """Sauvegarde les règles avec analyse détaillée"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"rules_corrected_{timestamp}.json"
    
    # Analyser quelques patterns pour vérification
    print("\n🔍 ANALYSE DES TOP RÈGLES:")
    for i, rule in enumerate(rules[:3]):
        pattern = rule['pattern']
        analyze_pattern_details(df, pattern, max_examples=5)
    
    rules_data = {
        'extraction_time': datetime.now().isoformat(),
        'total_rules': len(rules),
        'system_type': 'rules_corrected',
        'extraction_method': 'include_all_values_including_constants',
        'rules': rules
    }
    
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(rules_data, f, indent=2, ensure_ascii=False)
        
        print(f"\n💾 Règles corrigées sauvegardées: {filename}")
        return filename
        
    except Exception as e:
        print(f"❌ Erreur sauvegarde: {e}")
        return None

def main():
    """Fonction principale CORRIGÉE"""
    print("🚀 ENTRAÎNEMENT RÈGLES INTELLIGENTES CORRIGÉES")
    print("=" * 60)
    print("Mode: Extraction COMPLÈTE avec valeurs constantes")
    print()
    
    # 1. Charger données
    df = load_training_data()
    if df is None:
        return
    
    # 2. Extraire règles intelligentes corrigées
    rules = extract_intelligent_rules_corrected(df)
    
    # 3. Sauvegarder avec analyse
    rules_file = save_rules_with_analysis(rules, df)
    
    # 4. Résumé
    print("\n" + "=" * 60)
    print("🎉 EXTRACTION CORRIGÉE TERMINÉE!")
    print(f"✅ {len(rules)} règles intelligentes extraites")
    print(f"💾 Fichier: {rules_file}")
    
    # 5. Exemples de règles avec détails
    print("\n📋 TOP 5 RÈGLES CORRIGÉES:")
    for i, rule in enumerate(rules[:5], 1):
        pattern = rule['pattern'][:40]
        support = rule['support']
        confidence = rule['global_confidence']
        fixed_cols = rule.get('fixed_columns', {})
        variable_cols = rule.get('variable_columns', {})
        
        print(f"  {i}. Pattern: '{pattern}...'")
        print(f"     Support: {support}, Confiance: {confidence:.2f}")
        print(f"     Colonnes fixes: {list(fixed_cols.keys())}")
        if fixed_cols:
            for col, val in list(fixed_cols.items())[:3]:  # Top 3 colonnes
                print(f"       {col} = '{val}'")
        print()
    
    print("🚀 Système de règles corrigé prêt!")
    print("Test: python src/main.py")

if __name__ == "__main__":
    main()

# Ajouter à la fin de src/routes/excel.py

class SimpleRulesPredictor:
    """Prédicteur simple avec règles intelligentes CORRIGÉES"""
    
    def __init__(self):
        self.rules = []
    
    def load_rules(self):
        """Charge les règles depuis le fichier le plus récent"""
        import glob
        
        # Chercher les fichiers de règles (priorité aux corrigées)
        rule_files = (
            glob.glob("rules_corrected_*.json") + 
            glob.glob("rules_only_*.json") + 
            glob.glob("intelligent_rules_*.json")
        )
        
        if rule_files:
            # Prendre le plus récent
            latest_file = max(rule_files, key=os.path.getctime)
            print(f"📋 Chargement des règles corrigées: {latest_file}")
            
            try:
                with open(latest_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                self.rules = data.get('rules', [])
                print(f"✅ {len(self.rules)} règles corrigées chargées")
                
                # Afficher résumé
                rule_types = {}
                for rule in self.rules:
                    rule_type = rule.get('rule_type', 'unknown')
                    rule_types[rule_type] = rule_types.get(rule_type, 0) + 1
                
                for rule_type, count in rule_types.items():
                    print(f"   📋 {rule_type}: {count} règles")
                
                # Montrer exemple de règle "comfort"
                comfort_rules = [r for r in self.rules if 'comfort' in r.get('pattern', '').lower()]
                if comfort_rules:
                    print(f"\n🎯 Exemple règle 'comfort':")
                    rule = comfort_rules[0]
                    for col, val in rule.get('fixed_columns', {}).items():
                        print(f"   {col} = '{val}'")
                
                return True
                
            except Exception as e:
                print(f"❌ Erreur chargement règles: {e}")
                return False
        else:
            print("⚠️ Aucun fichier de règles trouvé")
            return False
    
    def apply_rules_to_dataframe(self, df):
        """Applique les règles corrigées au DataFrame"""
        if not self.rules:
            print("⚠️ Aucune règle chargée")
            return df
        
        print(f"🎯 Application de {len(self.rules)} règles corrigées...")
        
        # Suivre les modifications
        total_filled = 0
        applied_rules = 0
        
        # Prendre les top 100 règles pour performance
        for rule in self.rules[:100]:
            pattern = rule.get('pattern', '')
            
            if len(pattern) < 3:
                continue
            
            try:
                # Trouver les lignes qui correspondent au pattern
                mask = df['Description'].str.lower().str.contains(
                    re.escape(pattern.lower()), na=False, regex=True
                )
                
                if mask.sum() == 0:
                    continue
                
                rule_filled = 0
                
                # Appliquer TOUTES les colonnes fixes (y compris "N/A")
                for col, value in rule.get('fixed_columns', {}).items():
                    if col in df.columns:
                        # Identifier les cellules vides
                        empty_mask = (df.loc[mask, col].isna() | (df.loc[mask, col] == ''))
                        
                        if empty_mask.sum() > 0:
                            df.loc[mask & empty_mask, col] = value
                            rule_filled += empty_mask.sum()
                
                # Appliquer les colonnes variables (haute confiance seulement)
                for col, var_info in rule.get('variable_columns', {}).items():
                    if col in df.columns and isinstance(var_info, dict):
                        confidence = var_info.get('confidence', 0)
                        if confidence > 0.8:  # Seuil élevé
                            default_value = var_info.get('default_value')
                            if default_value:
                                empty_mask = (df.loc[mask, col].isna() | (df.loc[mask, col] == ''))
                                
                                if empty_mask.sum() > 0:
                                    df.loc[mask & empty_mask, col] = default_value
                                    rule_filled += empty_mask.sum()
                
                if rule_filled > 0:
                    applied_rules += 1
                    total_filled += rule_filled
                    print(f"  ✅ Règle '{pattern[:30]}...' → {rule_filled} cellules remplies")
                
            except Exception as e:
                print(f"  ⚠️ Erreur règle '{pattern[:20]}...': {e}")
                continue
        
        print(f"✅ Résultat: {applied_rules} règles appliquées, {total_filled} cellules remplies")
        return df

def apply_rules(df):
    """Version CORRIGÉE - Règles intelligentes avec valeurs constantes"""
    print("🔧 Application des règles intelligentes CORRIGÉES...")

    try:
        # Utiliser le système de règles corrigé
        predictor = SimpleRulesPredictor()
        
        # Charger les règles corrigées
        if predictor.load_rules():
            # Appliquer les règles
            df_result = predictor.apply_rules_to_dataframe(df.copy())
            return df_result
        else:
            print("❌ Impossible de charger les règles")
            return df.copy()
        
    except Exception as e:
        print(f"⚠️ Erreur règles corrigées: {e}")
        return df.copy()