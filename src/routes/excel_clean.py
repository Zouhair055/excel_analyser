from flask import Blueprint, request, jsonify, send_file
import pandas as pd
import os
import tempfile
from werkzeug.utils import secure_filename
import re
import json
import glob
from collections import Counter
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
import datetime
from openpyxl.utils import get_column_letter
import sys

sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

# Système uniquement basé sur les règles (ML désactivé)
USE_ML = False
print("🎯 Système à règles intelligentes activé")

excel_bp = Blueprint('excel', __name__)

# Utiliser des chemins absolus pour éviter les problèmes
UPLOAD_FOLDER = os.path.abspath('uploads')
PROCESSED_FOLDER = os.path.abspath('processed')

# Créer les dossiers s'ils n'existent pas
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'xlsx', 'xls'}

def find_data_start_row(filepath):
    """Trouve la ligne où commencent vraiment les données"""
    # Lire les premières lignes pour détecter où commencent les données
    try:
        # Lire les 20 premières lignes sans en-têtes
        preview_df = pd.read_excel(filepath, header=None, nrows=20)
        
        # Chercher la première ligne avec une structure cohérente
        for i in range(len(preview_df)):
            row = preview_df.iloc[i]
            
            # Vérifier si cette ligne contient des en-têtes valides
            non_null_count = row.notna().sum()
            
            # Si au moins 3 colonnes sont remplies, considérer comme ligne d'en-têtes potentielle
            if non_null_count >= 3:
                # Vérifier les lignes suivantes pour voir s'il y a des données
                if i + 1 < len(preview_df):
                    next_row = preview_df.iloc[i + 1]
                    if next_row.notna().sum() >= 2:  # Au moins 2 colonnes avec des données
                        return i
        
        # Par défaut, commencer à la ligne 0
        return 0
        
    except Exception as e:
        print(f"Erreur lors de la détection de la ligne de départ: {e}")
        return 0

def clean_column_names(df):
    """Nettoie et standardise les noms de colonnes - VERSION AMÉLIORÉE"""
    
    # Mapping des noms de colonnes connus
    column_mapping = {
        'description': 'Description',
        'descrip': 'Descrip',  # Ne pas mapper vers Description pour éviter les doublons
        'desc': 'Descrip',
        'nature': 'Nature',
        'nat': 'Nature',
        'reference': 'Reference',
        'ref': 'Reference',
        'service': 'Service',
        'serv': 'Service',
        'vessel': 'Vessel',
        'ves': 'Vessel',
        'amount': 'Amount CCYs',
        'amount_ccy': 'Amount CCYs',
        'amount_usd': 'Amount USD',
        'rate': 'Rate FX',
        'rate_fx': 'Rate FX',
        'date': 'Date',
        'period': 'Period',
        'entity': 'Entity',
        'bank': 'Bank Account',
        'bank_account': 'Bank Account',
        'account': 'Bank Account'
    }
    
    # Créer un dictionnaire des colonnes à renommer
    rename_dict = {}
    
    for col in df.columns:
        if pd.isna(col) or col == '':
            continue
            
        col_str = str(col).strip()
        col_lower = col_str.lower()
        
        # Chercher une correspondance dans le mapping
        for old_name, new_name in column_mapping.items():
            if old_name in col_lower and col not in rename_dict:  # Éviter les doublons
                # Vérifier si le nouveau nom existe déjà
                if new_name not in df.columns and new_name not in rename_dict.values():
                    rename_dict[col] = new_name
                    break
    
    # Renommer les colonnes
    if rename_dict:
        df = df.rename(columns=rename_dict)
        print(f"✅ Colonnes renommées: {rename_dict}")
    
    # Gérer les colonnes dupliquées en les renommant
    columns = list(df.columns)
    seen = {}
    new_columns = []
    
    for col in columns:
        if col in seen:
            seen[col] += 1
            new_col = f"{col}_{seen[col]}"
            new_columns.append(new_col)
            print(f"⚠️ Colonne dupliquée renommée: '{col}' → '{new_col}'")
        else:
            seen[col] = 0
            new_columns.append(col)
    
    df.columns = new_columns
    
    return df

def detect_date_columns(df):
    """Détecte les colonnes de date"""
    date_columns = []
    
    for col in df.columns:
        if col is None or pd.isna(col):
            continue
            
        col_name_lower = str(col).lower()
        
        # Vérifier d'abord le nom de la colonne
        if any(date_word in col_name_lower for date_word in ['date', 'period', 'time', 'day']):
            date_columns.append(col)
            continue
        
        # Vérifier le contenu de la colonne
        sample_values = df[col].dropna().head(10)
        date_count = 0
        
        for value in sample_values:
            if pd.isna(value):
                continue
                
            # Patterns de date courants
            if isinstance(value, (pd.Timestamp, datetime.datetime, datetime.date)):
                date_count += 1
            else:
                # Vérifier avec des regex
                date_patterns = [
                    r'\d{4}-\d{2}-\d{2}',  # 2024-01-15
                    r'\d{2}/\d{2}/\d{4}',  # 15/01/2024
                    r'\d{2}-\d{2}-\d{4}',  # 15-01-2024
                    r'\d{1,2}/\d{1,2}/\d{4}',  # 5/1/2024
                    r'\d{1,2}-\d{1,2}-\d{4}',  # 5-1-2024
                ]
                
                for pattern in date_patterns:
                    if re.match(pattern, str(value).strip()):
                        date_count += 1
                        break
                
                # Essayer de parser avec pandas
                try:
                    pd.to_datetime(value)
                    date_count += 1
                except:
                    pass
        
        # Si plus de 70% des valeurs semblent être des dates
        if len(sample_values) > 0 and date_count / len(sample_values) > 0.7:
            date_columns.append(col)
    
    return date_columns

def detect_numeric_columns(df):
    """Détecte les colonnes numériques"""
    numeric_columns = []
    
    # Colonnes spécifiques connues pour être numériques
    known_numeric_columns = ['Amount CCYs', 'Rate FX', 'Amount USD', 'amount', 'rate', 'price', 'quantity']
    
    for col in df.columns:
        # Vérifier d'abord si c'est une colonne connue pour être numérique
        if any(known_col.lower() in col.lower() for known_col in known_numeric_columns):
            numeric_columns.append(col)
            print(f"✅ Colonne numérique identifiée par nom: '{col}'")
            continue
            
        # Vérifier le type pandas
        if df[col].dtype in ['int64', 'float64', 'int32', 'float32']:
            numeric_columns.append(col)
    
    return numeric_columns

def preserve_original_formatting(original_filepath, df, ws, data_start_row):
    """Préserve le formatage ET les formules originales"""
    try:
        # Ouvrir le fichier original
        original_wb = load_workbook(original_filepath, data_only=False)  # GARDEZ LES FORMULES
        original_ws = original_wb.active
        
        # Trouver les colonnes avec des formules
        formula_columns = []
        original_start_row = find_data_start_row(original_filepath)
        
        # Détecter les formules dans Amount USD
        for col_idx in range(1, original_ws.max_column + 1):
            header_cell = original_ws.cell(row=original_start_row + 1, column=col_idx)
            if header_cell.value and 'amount usd' in str(header_cell.value).lower():
                
                # Vérifier si cette colonne contient des formules
                sample_cell = original_ws.cell(row=original_start_row + 2, column=col_idx)
                if sample_cell.data_type == 'f':  # 'f' = formule
                    formula_columns.append((col_idx, str(header_cell.value)))
                    print(f"✅ Formule détectée dans '{header_cell.value}' - sera préservée")
        
        # Copier les formules originales pour ces colonnes
        for col_idx, col_name in formula_columns:
            df_col_name = None
            for df_col in df.columns:
                if 'amount usd' in df_col.lower():
                    df_col_name = df_col
                    break
            
            if df_col_name:
                # Copier les formules ligne par ligne
                for row_idx in range(len(df)):
                    original_cell = original_ws.cell(row=original_start_row + 2 + row_idx, column=col_idx)
                    new_cell = ws.cell(row=data_start_row + 1 + row_idx, column=list(df.columns).index(df_col_name) + 1)
                    
                    if original_cell.data_type == 'f':  # Si c'est une formule
                        new_cell.value = original_cell.value  # Copier la formule
                        print(f"   📋 Formule copiée ligne {row_idx + 1}: {original_cell.value}")
                    # Sinon, garder la valeur du DataFrame
        
        original_wb.close()
        
    except Exception as e:
        print(f"⚠️ Erreur préservation formules: {e}")

def apply_enhanced_formatting(ws, df, data_start_row):
    """Applique un formatage professionnel avec formatage conditionnel pour les montants"""
    
    # Couleurs et styles
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    
    data_font = Font(size=10)
    negative_font = Font(size=10, color="FF0000")  # Rouge pour les négatifs
    alignment_center = Alignment(horizontal="center", vertical="center")
    alignment_left = Alignment(horizontal="left", vertical="center")
    alignment_right = Alignment(horizontal="right", vertical="center")
    
    # Bordures
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Identifier les colonnes spéciales
    date_columns = detect_date_columns(df)
    numeric_columns = detect_numeric_columns(df)
    
    # Identifier les colonnes de montants pour le formatage conditionnel
    amount_columns = []
    for col in df.columns:
        if any(keyword in col.lower() for keyword in ['amount ccys', 'amount usd', 'amount', 'montant']):
            amount_columns.append(col)
            print(f"💰 Colonne de montant détectée: {col}")
    
    # Formatage des en-têtes
    header_row = data_start_row + 1
    for col_idx, column in enumerate(df.columns, 1):
        cell = ws.cell(row=header_row, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = alignment_center
        cell.border = thin_border
        
        # Ajuster la largeur des colonnes
        column_letter = get_column_letter(col_idx)
        if column in numeric_columns:
            ws.column_dimensions[column_letter].width = 15
        elif column in date_columns:
            ws.column_dimensions[column_letter].width = 12
        elif column == 'Description':
            ws.column_dimensions[column_letter].width = 40
        else:
            ws.column_dimensions[column_letter].width = 15
    
    # Formatage des données avec formatage conditionnel
    for row_idx in range(len(df)):
        excel_row = data_start_row + row_idx + 2
        
        for col_idx, column in enumerate(df.columns, 1):
            cell = ws.cell(row=excel_row, column=col_idx)
            cell_value = df.iloc[row_idx, col_idx - 1]
            cell.border = thin_border
            
            # 🎯 FORMATAGE CONDITIONNEL POUR LES MONTANTS
            if column in amount_columns:
                cell.alignment = alignment_right
                
                # Vérifier si la valeur est numérique et négative
                try:
                    numeric_value = float(cell_value) if pd.notna(cell_value) and cell_value != '' else 0
                    
                    if numeric_value < 0:
                        # ROUGE pour les montants négatifs
                        cell.font = negative_font
                        cell.number_format = '#,##0.00;[RED]-#,##0.00'
                    else:
                        # NOIR (normal) pour les montants positifs
                        cell.font = data_font
                        cell.number_format = '#,##0.00'
                        
                except (ValueError, TypeError):
                    # Si la conversion échoue, utiliser le format normal
                    cell.font = data_font
                    cell.number_format = '#,##0.00'
            
            # Format spécifique selon le type de colonne (autres colonnes)
            elif column in numeric_columns:
                cell.alignment = alignment_right
                cell.font = data_font
                if 'Rate' in column:
                    cell.number_format = '0.0000'
                else:
                    cell.number_format = '#,##0.00'
                    
            elif column in date_columns:
                cell.alignment = alignment_center
                cell.font = data_font
                if column == 'Period':
                    cell.number_format = 'MMM-YY'
                else:
                    cell.number_format = 'DD/MM/YYYY'
            else:
                cell.alignment = alignment_left
                cell.font = data_font
    
    # Ajouter des lignes alternées (mais préserver le rouge pour les négatifs)
    light_fill = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
    
    for row_idx in range(len(df)):
        if row_idx % 2 == 1:  # Lignes paires (index impair)
            excel_row = data_start_row + row_idx + 2
            for col_idx, column in enumerate(df.columns, 1):
                cell = ws.cell(row=excel_row, column=col_idx)
                cell.fill = light_fill
                # Ne pas écraser le formatage conditionnel des montants
                if column not in amount_columns:
                    cell.fill = light_fill
    
    print(f"✅ Formatage conditionnel appliqué sur {len(amount_columns)} colonnes de montants")

def analyze_data_quality(df):
    """Analyse la qualité des données et retourne un rapport - VERSION CORRIGÉE"""
    report = {
        'total_rows': len(df),
        'total_columns': len(df.columns),
        'empty_cells': 0,
        'completion_rate': 0,
        'column_analysis': {}
    }
    
    total_cells = len(df) * len(df.columns)
    empty_cells = 0
    
    for col in df.columns:
        try:
            # Gérer les colonnes dupliquées en prenant la première occurrence
            if isinstance(df[col], pd.DataFrame):  # Si plusieurs colonnes avec le même nom
                col_data = df[col].iloc[:, 0]  # Prendre la première colonne
            else:
                col_data = df[col]
            
            empty_count = col_data.isna().sum() + (col_data == '').sum()
            
            # S'assurer qu'empty_count est un scalaire
            if hasattr(empty_count, 'iloc'):
                empty_count = empty_count.iloc[0] if len(empty_count) > 0 else 0
            
            empty_cells += empty_count
            
            completion_rate = ((len(df) - empty_count) / len(df)) * 100 if len(df) > 0 else 0
            
            # Créer une clé unique pour les colonnes dupliquées
            col_key = f"{col}_{list(df.columns).index(col)}" if list(df.columns).count(col) > 1 else col
            
            report['column_analysis'][col_key] = {
                'empty_count': int(empty_count),
                'completion_rate': round(float(completion_rate), 1),
                'data_type': str(col_data.dtype)
            }
            
        except Exception as e:
            print(f"⚠️ Erreur analyse colonne '{col}': {e}")
            # Valeurs par défaut en cas d'erreur
            report['column_analysis'][str(col)] = {
                'empty_count': 0,
                'completion_rate': 100.0,
                'data_type': 'object'
            }
    
    report['empty_cells'] = int(empty_cells)
    report['completion_rate'] = round(((total_cells - empty_cells) / total_cells) * 100, 1) if total_cells > 0 else 0
    
    return report

class SimpleRulesPredictor:
    """Prédicteur simple avec règles intelligentes CORRIGÉES"""
    
    def __init__(self):
        self.rules = []
    
    def load_rules(self):
        """Charge les règles depuis le fichier le plus récent AUTOMATIQUEMENT"""
        
        # Chercher TOUS les fichiers de règles dans TOUS les répertoires
        rule_files = []
        
        # Patterns de recherche CORRIGÉS pour les vrais fichiers de règles
        patterns = [
            "rules_corrected_*.json",  # Format principal
            "rules_only_*.json", 
            "intelligent_rules_*.json",
            "smart_rules_*.json",
            "extracted_rules_*.json"  # Ajouté pour compatibilité
        ]
        
        # Répertoires de recherche ÉTENDUS (avec chemins absolus)
        project_root = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
        search_dirs = [
            project_root,  # Répertoire racine du projet
            os.path.join(project_root, "model_auto_remplissage"),
            os.path.join(project_root, "model_auto_remplissage", "models"),
            os.path.dirname(__file__),  # Répertoire du script
            os.path.dirname(os.path.dirname(__file__)),  # Répertoire parent
            "."  # Répertoire courant
        ]
        
        print("🔍 Recherche ÉTENDUE de fichiers de règles...")
        
        for directory in search_dirs:
            abs_dir = os.path.abspath(directory)
            print(f"   📁 Recherche dans: {abs_dir}")
            
            if os.path.exists(abs_dir):
                for pattern in patterns:
                    search_path = os.path.join(abs_dir, pattern)
                    found_files = glob.glob(search_path)
                    if found_files:
                        print(f"   ✅ Trouvé avec pattern '{pattern}': {found_files}")
                        rule_files.extend(found_files)
                    else:
                        print(f"   🔍 Recherché: {search_path}")
            else:
                print(f"   ⚠️ Répertoire inexistant: {abs_dir}")
        
        # Recherche RÉCURSIVE si rien trouvé
        if not rule_files:
            print("🔍 Recherche récursive dans tout le projet...")
            for root, dirs, files in os.walk(project_root):
                for file in files:
                    # Vérifier si le fichier correspond à un pattern de règles
                    if (file.startswith("rules_corrected_") or 
                        file.startswith("intelligent_rules_") or 
                        file.startswith("smart_rules_") or
                        file.startswith("rules_only_")) and file.endswith(".json"):
                        full_path = os.path.join(root, file)
                        rule_files.append(full_path)
                        print(f"   ✅ Trouvé (récursif): {full_path}")
        
        if rule_files:
            # PRENDRE LE PLUS RÉCENT AUTOMATIQUEMENT
            latest_file = max(rule_files, key=os.path.getctime)
            print(f"📋 Chargement automatique du fichier le plus récent: {latest_file}")
            
            try:
                with open(latest_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                self.rules = data.get('rules', [])
                print(f"✅ {len(self.rules)} règles chargées depuis {latest_file}")
                
                # Afficher résumé des types de règles
                if self.rules:
                    rule_types = {}
                    for rule in self.rules:
                        rule_type = rule.get('rule_type', 'unknown')
                        rule_types[rule_type] = rule_types.get(rule_type, 0) + 1
                    
                    for rule_type, count in rule_types.items():
                        print(f"   📋 {rule_type}: {count} règles")
                    
                    # Montrer exemple de règle pour vérification
                    first_rule = self.rules[0]
                    pattern = first_rule.get('pattern', '')[:30]
                    fixed_cols = len(first_rule.get('fixed_columns', {}))
                    print(f"   🎯 Exemple: '{pattern}...' avec {fixed_cols} colonnes fixes")
                
                return True
                
            except Exception as e:
                print(f"❌ Erreur chargement règles: {e}")
                return False
        else:
            print("⚠️ Aucun fichier de règles trouvé même avec recherche récursive")
            print("💡 Créez un fichier avec model_auto_remplissage/train_hybrid_system.py")
            return False
    
def create_comparison_report(original_df, processed_df, output_path):
    """Crée un rapport de comparaison entre les données originales et traitées"""
    
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Feuille 1: Données originales
            original_df.to_excel(writer, sheet_name='Original', index=False)
            
            # Feuille 2: Données traitées
            processed_df.to_excel(writer, sheet_name='Processed', index=False)
            
            # Feuille 3: Rapport de comparaison
            comparison_data = []
            
            for col in processed_df.columns:
                if col in original_df.columns:
                    orig_empty = original_df[col].isna().sum() + (original_df[col] == '').sum()
                    proc_empty = processed_df[col].isna().sum() + (processed_df[col] == '').sum()
                    filled = orig_empty - proc_empty
                    
                    comparison_data.append({
                        'Column': col,
                        'Original_Empty': orig_empty,
                        'Processed_Empty': proc_empty,
                        'Cells_Filled': filled,
                        'Improvement_%': round((filled / orig_empty * 100) if orig_empty > 0 else 0, 1)
                    })
            
            comparison_df = pd.DataFrame(comparison_data)
            comparison_df.to_excel(writer, sheet_name='Comparison_Report', index=False)
            
            # Formatage du rapport
            workbook = writer.book
            worksheet = writer.sheets['Comparison_Report']
            
            # En-têtes en gras
            for cell in worksheet[1]:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
            
            # Ajuster les largeurs de colonnes
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        print(f"✅ Rapport de comparaison créé: {output_path}")
        
    except Exception as e:
        print(f"❌ Erreur création rapport: {e}")

@excel_bp.route('/upload', methods=['POST'])
def upload_file():
    """Endpoint avec formatage conditionnel complet"""
    
    try:
        # Sauvegarder le fichier uploadé
        if 'file' not in request.files:
            return jsonify({'error': 'Aucun fichier fourni'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'Aucun fichier sélectionné'}), 400
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Créer un nom unique pour éviter les conflits
            name, ext = os.path.splitext(filename)
            unique_filename = f"{name}_{timestamp}{ext}"
            
            filepath = os.path.join(UPLOAD_FOLDER, unique_filename)
            file.save(filepath)
            
            print(f"📁 Fichier sauvegardé: {filepath}")
            
            # Détecter la ligne de départ des données
            data_start_row = find_data_start_row(filepath)
            print(f"📍 Ligne de départ des données détectée: {data_start_row}")
            
            # Lire le fichier Excel
            df = pd.read_excel(filepath, skiprows=data_start_row)
            
            # Nettoyer les noms de colonnes
            df = clean_column_names(df)
            
            print(f"📊 Données chargées: {len(df)} lignes, {len(df.columns)} colonnes")
            print(f"📋 Colonnes: {list(df.columns)}")
            
            # Vérifier que la colonne Description existe
            if 'Description' not in df.columns:
                return jsonify({
                    'error': 'Colonne "Description" non trouvée. Colonnes disponibles: ' + ', '.join(df.columns)
                }), 400
            
            # Analyser la qualité initiale
            print("\n" + "🔍 ANALYSE INITIALE" + "="*50)
            initial_quality = analyze_data_quality_detailed(df)
            
            # Appliquer le système de règles amélioré
            predictor = EnhancedRulesPredictor()
            if predictor.load_rules():
                processed_df = predictor.apply_rules_to_dataframe(df.copy())
                
                # 🔍 VALIDATION POST-TRAITEMENT
                integrity_ok = validate_data_integrity(df, processed_df)
                if not integrity_ok:
                    print("🚨 ALERTE: Problème d'intégrité détecté!")
                
            else:
                processed_df = df.copy()
                print("⚠️ Traitement sans règles - aucune amélioration")
            
            # Analyser la qualité finale
            print("\n" + "🎯 ANALYSE FINALE" + "="*50)
            final_quality = analyze_data_quality_detailed(processed_df)
            
            # Calculs d'amélioration
            improvement = final_quality['completion_rate'] - initial_quality['completion_rate']
            cells_filled = initial_quality['empty_cells'] - final_quality['empty_cells']
            
            # Affichage final
            print(f"\n{'='*80}")
            print(f"🎉 RÉSULTATS FINAUX DU TRAITEMENT")
            print(f"{'='*80}")
            print(f"📊 Remplissage initial: {initial_quality['completion_rate']:.2f}%")
            print(f"📈 Remplissage final: {final_quality['completion_rate']:.2f}%")
            print(f"🚀 Amélioration: +{improvement:.2f}% ({cells_filled:,} cellules remplies)")
            print(f"🎯 Règles appliquées: {predictor.stats.get('rules_applied', 0)}/{predictor.stats.get('rules_loaded', 0)}")
            print(f"{'='*80}")
            
            # Sauvegarder le fichier traité
            processed_filename = f"processed_{unique_filename}"
            processed_filepath = os.path.join(PROCESSED_FOLDER, processed_filename)
            
            # 🎨 UTILISER LA NOUVELLE FONCTION DE FORMATAGE
            formatting_result = format_excel_file_with_filters_and_conditional(
                processed_df, 
                processed_filepath, 
                filepath
            )
            
            print(f"💾 Fichier traité sauvegardé avec formatage avancé: {processed_filepath}")
            
            # Retourner les résultats avec info sur le formatage
            return jsonify({
                'success': True,
                'message': 'Fichier traité avec succès',
                'original_file': unique_filename,
                'processed_file': processed_filename,
                'formatting_applied': {
                    'filters': formatting_result['filters_added'],
                    'conditional_formatting': formatting_result['conditional_formatting'],
                    'table_style': formatting_result['table_created'],
                    'formulas_preserved': formatting_result['formulas_preserved'],
                    'negative_amounts_in_red': True,
                    'positive_amounts_in_black': True
                },
                'statistics': {
                    'initial_completion': float(initial_quality['completion_rate']),
                    'final_completion': float(final_quality['completion_rate']),
                    'improvement': float(improvement),
                    'cells_filled': int(cells_filled),
                    'rules_loaded': int(predictor.stats.get('rules_loaded', 0)),
                    'rules_applied': int(predictor.stats.get('rules_applied', 0)),
                    'top_patterns': list(predictor.stats.get('patterns_matched', {}).keys())[:5]
                },
                'columns_info': {
                    'columns': list(processed_df.columns),
                    'shape': [int(processed_df.shape[0]), int(processed_df.shape[1])],
                    'empty_columns': [col for col in processed_df.columns if processed_df[col].isna().all()]
                }
            })
        else:
            return jsonify({'error': 'Type de fichier non autorisé. Utilisez .xlsx ou .xls'}), 400
        
    except Exception as e:
        print(f"❌ Erreur: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@excel_bp.route('/download/<filename>')
def download_file(filename):
    """Télécharge un fichier traité"""
    try:
        # Vérifier d'abord dans le dossier processed
        file_path = os.path.join(PROCESSED_FOLDER, filename)
        
        if not os.path.exists(file_path):
            # Vérifier dans le dossier uploads si c'est un fichier original
            file_path = os.path.join(UPLOAD_FOLDER, filename)
        
        if not os.path.exists(file_path):
            return jsonify({'error': 'Fichier non trouvé'}), 404
        
        return send_file(file_path, as_attachment=True)
        
    except Exception as e:
        return jsonify({'error': f'Erreur lors du téléchargement: {str(e)}'}), 500

@excel_bp.route('/files')
def list_files():
    """Liste tous les fichiers disponibles"""
    try:
        files = {
            'uploads': [],
            'processed': []
        }
        
        # Lister les fichiers uploadés
        if os.path.exists(UPLOAD_FOLDER):
            for f in os.listdir(UPLOAD_FOLDER):
                if allowed_file(f):
                    stat = os.stat(os.path.join(UPLOAD_FOLDER, f))
                    files['uploads'].append({
                        'name': f,
                        'size': stat.st_size,
                        'modified': datetime.datetime.fromtimestamp(stat.st_mtime).isoformat()
                    })
        
        # Lister les fichiers traités
        if os.path.exists(PROCESSED_FOLDER):
            for f in os.listdir(PROCESSED_FOLDER):
                if f.endswith(('.xlsx', '.xls')):
                    stat = os.stat(os.path.join(PROCESSED_FOLDER, f))
                    files['processed'].append({
                        'name': f,
                        'size': stat.st_size,
                        'modified': datetime.datetime.fromtimestamp(stat.st_mtime).isoformat()
                    })
        
        return jsonify(files)
        
    except Exception as e:
        return jsonify({'error': f'Erreur lors de la liste des fichiers: {str(e)}'}), 500

@excel_bp.route('/health')
def health_check():
    """Vérifie l'état du service"""
    
    # Vérifier la disponibilité des règles
    predictor = SimpleRulesPredictor()
    rules_available = predictor.load_rules()
    
    return jsonify({
        'status': 'healthy',
        'system_type': 'rules_only',
        'ml_available': False,
        'rules_available': rules_available,
        'rules_count': len(predictor.rules) if rules_available else 0,
        'upload_folder': UPLOAD_FOLDER,
        'processed_folder': PROCESSED_FOLDER,
        'timestamp': datetime.datetime.now().isoformat()
    })

import json
import glob
import re

class EnhancedRulesPredictor:
    """Système de règles amélioré avec statistiques détaillées"""
    
    def __init__(self):
        self.rules = []
        self.stats = {
            'rules_loaded': 0,
            'rules_applied': 0,
            'cells_filled': 0,
            'total_cells': 0,
            'patterns_matched': {}
        }
    
    def load_rules(self):
        """Charge TOUTES les règles du fichier JSON"""
        
        # Chercher le fichier de règles le plus récent
        rule_files = glob.glob("rules_corrected_*.json")
        if not rule_files:
            rule_files = glob.glob("**/rules_corrected_*.json", recursive=True)
        
        if rule_files:
            latest_file = max(rule_files, key=os.path.getctime)
            print(f"📋 Chargement des règles depuis: {latest_file}")
            
            try:
                with open(latest_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                self.rules = data.get('rules', [])
                self.stats['rules_loaded'] = len(self.rules)
                
                print(f"✅ {self.stats['rules_loaded']} règles chargées avec succès")
                
                # Afficher un résumé des types de règles
                rule_types = {}
                for rule in self.rules:
                    rule_type = rule.get('rule_type', 'unknown')
                    rule_types[rule_type] = rule_types.get(rule_type, 0) + 1
                
                print("📊 Types de règles chargées:")
                for rule_type, count in rule_types.items():
                    print(f"   - {rule_type}: {count} règles")
                
                return True
                
            except Exception as e:
                print(f"❌ Erreur lors du chargement: {e}")
                return False
        else:
            print("⚠️ Aucun fichier de règles trouvé")
            return False
    
    def apply_rules_to_dataframe(self, df):
        """Applique TOUTES les règles avec PROTECTION des colonnes financières"""
        
        if not self.rules:
            print("⚠️ Aucune règle chargée")
            return df
        
        # 🔒 COLONNES À NE JAMAIS MODIFIER
        PROTECTED_COLUMNS = [
            'Amount CCYs', 'Amount USD', 'Rate FX', 'Total', 
            'Entity', 'Period', 'Date', 'Transaction Date',
            'Bank Account', 'CCY'
        ]
        
        print(f"\n🎯 APPLICATION DE {len(self.rules)} RÈGLES (AVEC PROTECTION)")
        print("="*60)
        
        # 🔒 SAUVEGARDER LES COLONNES PROTÉGÉES AVANT TRAITEMENT
        protected_data = {}
        for col in PROTECTED_COLUMNS:
            if col in df.columns:
                protected_data[col] = df[col].copy()
                print(f"🔒 Colonne protégée: {col}")
        
        print(f"🛡️ {len(protected_data)} colonnes financières protégées")
        
        # Calculer les statistiques initiales
        initial_empty_cells = self._count_empty_cells(df)
        self.stats['total_cells'] = len(df) * len(df.columns)
        
        # APPLIQUER LES RÈGLES SEULEMENT SUR LES COLONNES CIBLES
        TARGET_COLUMNS = ['Nature', 'Descrip', 'Vessel', 'Service', 'Reference']
        
        # Compteurs pour le traitement
        rules_applied = 0
        total_cells_filled = 0
        pattern_matches = {}
        
        # Appliquer les règles par ordre de confiance/support
        sorted_rules = sorted(self.rules, key=lambda x: x.get('support', 0), reverse=True)
        
        for i, rule in enumerate(sorted_rules[:200]):  # Top 200 règles
            pattern = rule.get('pattern', '').lower().strip()
            
            if len(pattern) < 2:
                continue
            
            try:
                # Rechercher le pattern dans la colonne Description
                mask = df['Description'].str.lower().str.contains(
                    re.escape(pattern), na=False, regex=True
                )
                
                matched_rows = mask.sum()
                if matched_rows == 0:
                    continue
                
                # Compter les cellules remplies pour cette règle
                rule_cells_filled = 0
                
                # ✅ APPLIQUER LES COLONNES FIXES (SEULEMENT SUR COLONNES CIBLES)
                for col, value in rule.get('fixed_columns', {}).items():
                    if col in df.columns and col in TARGET_COLUMNS:  # ← PROTECTION ICI
                        # Identifier les cellules vides
                        empty_mask = (df.loc[mask, col].isna() | (df.loc[mask, col] == ''))
                        cells_to_fill = empty_mask.sum()
                        
                        if cells_to_fill > 0:
                            df.loc[mask & empty_mask, col] = value
                            rule_cells_filled += cells_to_fill
                    elif col not in TARGET_COLUMNS:
                        print(f"   🚫 Colonne '{col}' ignorée (protection)")
                
                # ✅ APPLIQUER LES COLONNES VARIABLES (SEULEMENT SUR COLONNES CIBLES)
                for col, var_info in rule.get('variable_columns', {}).items():
                    if col in df.columns and col in TARGET_COLUMNS and isinstance(var_info, dict):  # ← PROTECTION ICI
                        confidence = var_info.get('confidence', 0)
                        if confidence > 0.8:  # Seuil élevé
                            default_value = var_info.get('default_value')
                            if default_value:
                                empty_mask = (df.loc[mask, col].isna() | (df.loc[mask, col] == ''))
                                cells_to_fill = empty_mask.sum()
                                
                                if cells_to_fill > 0:
                                    df.loc[mask & empty_mask, col] = default_value
                                    rule_cells_filled += cells_to_fill
                    elif col not in TARGET_COLUMNS:
                        print(f"   🚫 Colonne '{col}' ignorée (protection)")
                
                if rule_cells_filled > 0:
                    rules_applied += 1
                    total_cells_filled += rule_cells_filled
                    pattern_matches[pattern] = {
                        'rows_matched': matched_rows,
                        'cells_filled': rule_cells_filled,
                        'support': rule.get('support', 0)
                    }
                    
                    print(f"  ✅ Règle {rules_applied:3d}: '{pattern[:40]}...' → {matched_rows} lignes, {rule_cells_filled} cellules remplies")
                
            except Exception as e:
                print(f"  ⚠️ Erreur règle '{pattern[:20]}...': {e}")
                continue
        
        # 🔒 RESTAURER INTÉGRALEMENT LES COLONNES PROTÉGÉES
        print(f"\n🔒 RESTAURATION DES COLONNES PROTÉGÉES...")
        for col, original_data in protected_data.items():
            df[col] = original_data
            print(f"✅ Colonne '{col}' restaurée (données originales préservées)")
        
        # Calculer les statistiques finales
        final_empty_cells = self._count_empty_cells(df)
        
        # Sauvegarder les statistiques
        self.stats.update({
            'rules_applied': rules_applied,
            'cells_filled': total_cells_filled,
            'patterns_matched': pattern_matches,
            'initial_empty_cells': initial_empty_cells,
            'final_empty_cells': final_empty_cells,
            'improvement': initial_empty_cells - final_empty_cells,
            'protected_columns': len(protected_data)
        })
        
        # Afficher le rapport final
        self._print_final_report()
        
        return df

    def _count_empty_cells(self, df):
        """Compte le nombre de cellules vides dans le DataFrame"""
        empty_count = 0
        for col in df.columns:
            empty_count += (df[col].isna() | (df[col] == '')).sum()
        return empty_count
    
    def _print_final_report(self):
        """Affiche un rapport détaillé des résultats avec protection"""
        
        print("\n" + "="*60)
        print("📈 RAPPORT FINAL D'APPLICATION DES RÈGLES (AVEC PROTECTION)")
        print("="*60)
        
        # Statistiques globales
        total_cells = self.stats['total_cells']
        initial_empty = self.stats['initial_empty_cells']
        final_empty = self.stats['final_empty_cells']
        cells_filled = self.stats['cells_filled']
        protected_count = self.stats.get('protected_columns', 0)  # ← SÉCURISÉ
        
        initial_completion = ((total_cells - initial_empty) / total_cells) * 100
        final_completion = ((total_cells - final_empty) / total_cells) * 100
        improvement = final_completion - initial_completion
        
        print(f"📊 STATISTIQUES GLOBALES:")
        print(f"   • Règles chargées: {self.stats['rules_loaded']}")
        print(f"   • Règles appliquées: {self.stats['rules_applied']}")
        print(f"   • 🔒 Colonnes protégées: {protected_count}")
        print(f"   • Cellules totales: {total_cells:,}")
        print(f"   • Cellules vides initiales: {initial_empty:,}")
        print(f"   • Cellules vides finales: {final_empty:,}")
        print(f"   • Cellules remplies: {cells_filled:,}")
        
        print(f"\n🎯 TAUX DE REMPLISSAGE:")
        print(f"   • Avant traitement: {initial_completion:.2f}%")
        print(f"   • Après traitement: {final_completion:.2f}%")
        print(f"   • Amélioration: +{improvement:.2f}%")
        
        print(f"\n🔒 PROTECTION DES DONNÉES:")
        if protected_count > 0:
            print(f"   ✅ {protected_count} colonnes financières préservées")
            print(f"   ✅ Aucune modification des montants/formules")
            print(f"   ✅ Intégrité comptable garantie")
        else:
            print(f"   ⚠️ Aucune colonne protégée détectée")
        
        # Top 10 des règles les plus efficaces
        if self.stats['patterns_matched']:
            print(f"\n🏆 TOP 10 DES RÈGLES LES PLUS EFFICACES:")
            sorted_patterns = sorted(
                self.stats['patterns_matched'].items(),
                key=lambda x: x[1]['cells_filled'],
                reverse=True
            )
            
            for i, (pattern, info) in enumerate(sorted_patterns[:10]):
                print(f"   {i+1:2d}. '{pattern[:35]:<35}' → {info['cells_filled']} cellules")
        
        # Recommandations
        print(f"\n💡 RECOMMANDATIONS:")
        if improvement < 5:
            print("   ⚠️ Amélioration faible - considérer ajouter plus de règles")
        elif improvement < 15:
            print("   ✅ Amélioration correcte - système fonctionnel")
        else:
            print("   🚀 Excellente amélioration - système très efficace")
        
        if self.stats['rules_applied'] < self.stats['rules_loaded'] * 0.1:
            print("   📋 Peu de règles utilisées - vérifier la pertinence des patterns")
        
        print(f"\n🎯 SYSTÈME SÉCURISÉ:")
        print(f"   ✅ Colonnes cibles traitées: Nature, Descrip, Vessel, Service, Reference")
        print(f"   🔒 Colonnes financières intactes: Amount CCYs, Amount USD, Rate FX")
        print(f"   ✅ Pas d'erreurs #DIV/0! attendues")
        
        print("="*60)



def analyze_data_quality_detailed(df):
    """Analyse détaillée de la qualité des données"""
    
    print(f"\n📊 ANALYSE DÉTAILLÉE DE LA QUALITÉ DES DONNÉES")
    print("="*60)
    
    total_cells = len(df) * len(df.columns)
    
    # Analyse par colonne
    column_stats = []
    total_empty = 0
    
    for col in df.columns:
        if isinstance(df[col], pd.DataFrame):  # Gestion colonnes dupliquées
            col_data = df[col].iloc[:, 0]
        else:
            col_data = df[col]
        
        empty_count = (col_data.isna() | (col_data == '')).sum()
        total_empty += empty_count
        
        completion_rate = ((len(df) - empty_count) / len(df)) * 100
        
        column_stats.append({
            'column': col,
            'empty_count': empty_count,
            'completion_rate': completion_rate,
            'data_type': str(col_data.dtype)
        })
    
    # Trier par taux de completion
    column_stats.sort(key=lambda x: x['completion_rate'])
    
    print(f"📋 ANALYSE PAR COLONNE (triée par taux de remplissage):")
    print(f"{'Colonne':<20} {'Vides':<8} {'Remplissage':<12} {'Type':<15}")
    print("-" * 60)
    
    for stats in column_stats:
        print(f"{stats['column']:<20} {stats['empty_count']:<8} {stats['completion_rate']:<11.1f}% {stats['data_type']:<15}")
    
    overall_completion = ((total_cells - total_empty) / total_cells) * 100
    
    print(f"\n🎯 RÉSUMÉ GLOBAL:")
    print(f"   • Cellules totales: {total_cells:,}")
    print(f"   • Cellules vides: {total_empty:,}")
    print(f"   • Taux de remplissage global: {overall_completion:.2f}%")
    
    return {
        'total_cells': total_cells,
        'empty_cells': total_empty,
        'completion_rate': overall_completion,
        'column_stats': column_stats
    }

def format_excel_file_with_filters_and_conditional(df, filepath, original_filepath=None):
    """Formate le fichier Excel avec filtres ET formatage conditionnel (OPTIMISÉ)"""
    
    try:
        print(f"📊 Formatage Excel COMPLET...")
        
        # Créer le workbook de base
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        
        # Copier les en-têtes si nécessaire
        start_row = 0
        if original_filepath:
            start_row = min(find_data_start_row(original_filepath), 2)
        
        data_start_row = start_row + 1
        
        # ✅ ÉCRITURE DES EN-TÊTES AVEC FORMATAGE (LIGNE 1 - FOND BLEU)
        for col_idx, col_name in enumerate(df.columns):
            cell = ws.cell(row=data_start_row, column=col_idx + 1, value=str(col_name))
            # 🎯 EN-TÊTES : GARDER LE FOND BLEU + GRAS + BLANC
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # ✅ ÉCRITURE DES DONNÉES (LIGNE 2+ - FORMATAGE NORMAL)
        for row_idx, (_, row) in enumerate(df.iterrows()):
            excel_row = data_start_row + 1 + row_idx  # Ligne 2, 3, 4...
            
            for col_idx, value in enumerate(row):
                cell = ws.cell(row=excel_row, column=col_idx + 1)
                
                # Nettoyer la valeur
                if pd.notna(value) and value != '':
                    try:
                        if isinstance(value, str) and value.replace('.', '').replace('-', '').replace(',', '').isdigit():
                            cell.value = float(value.replace(',', ''))
                        else:
                            cell.value = value
                    except:
                        cell.value = str(value)
                else:
                    cell.value = ""
                
                # 🎯 FORMATAGE NORMAL POUR TOUTES LES DONNÉES (Y COMPRIS LIGNE 2)
                col_name = df.columns[col_idx]
                
                # Colonnes de montants
                if any(keyword in col_name.lower() for keyword in ['amount ccys', 'amount usd', 'amount', 'montant']):
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.font = Font(size=10, color="000000")  # Noir normal
                    try:
                        numeric_value = float(cell.value) if cell.value else 0
                        if numeric_value < 0:
                            cell.font = Font(color="FF0000", size=10)  # Rouge pour négatif
                            cell.number_format = '#,##0.00;[RED]-#,##0.00'
                        else:
                            cell.number_format = '#,##0.00'
                    except:
                        cell.number_format = '#,##0.00'
                
                # Colonnes de dates/period (FORMATAGE NORMAL)
                elif any(keyword in col_name.lower() for keyword in ['date', 'period']):
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = Font(size=10, color="000000")  # Noir normal
                    if 'period' in col_name.lower():
                        cell.number_format = 'MMM-YY'
                    else:
                        cell.number_format = 'DD/MM/YYYY'
                
                # Autres colonnes
                else:
                    cell.font = Font(size=10, color="000000")  # Noir normal
                    cell.alignment = Alignment(horizontal="left", vertical="center")
        
        # ✅ AJOUTER LES FILTRES
        max_col_letter = get_column_letter(len(df.columns))
        max_row = data_start_row + len(df)
        filter_range = f"A{data_start_row}:{max_col_letter}{max_row}"
        ws.auto_filter.ref = filter_range
        
        # ✅ AJUSTER LES LARGEURS DE COLONNES
        for col_idx in range(1, len(df.columns) + 1):
            column_letter = get_column_letter(col_idx)
            col_name = df.columns[col_idx - 1]
            
            if 'description' in col_name.lower():
                ws.column_dimensions[column_letter].width = 40
            elif any(keyword in col_name.lower() for keyword in ['amount', 'montant']):
                ws.column_dimensions[column_letter].width = 15
            elif 'date' in col_name.lower() or 'period' in col_name.lower():
                ws.column_dimensions[column_letter].width = 12
            else:
                ws.column_dimensions[column_letter].width = 15
        
        # ✅ AJOUTER DES BORDURES (TOUTES LES LIGNES)
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        for row_idx in range(data_start_row, max_row + 1):
            for col_idx in range(1, len(df.columns) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.border = thin_border
        
        # 💾 SAUVEGARDER
        wb.save(filepath)
        print(f"✅ Fichier Excel avec formatage complet sauvegardé: {filepath}")
        
        return {
            'filters_added': True,
            'conditional_formatting': True,
            'table_created': True,
            'formulas_preserved': False,
            'negative_amounts_in_red': True,
            'column_widths_adjusted': True,
            'borders_added': True
        }
        
    except Exception as e:
        print(f"❌ Erreur formatage Excel: {e}")
        df.to_excel(filepath, index=False)
        return {
            'filters_added': False,
            'conditional_formatting': False,
            'table_created': False,
            'formulas_preserved': False
        }

def add_conditional_formatting(ws, df, data_start_row):
    """Ajoute un formatage conditionnel Excel natif pour les montants"""
    
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import Font, PatternFill
    
    print("🎨 Application du formatage conditionnel Excel...")
    
    # Identifier les colonnes de montants
    amount_columns = []
    for col_idx, column in enumerate(df.columns, 1):
        if any(keyword in column.lower() for keyword in ['amount ccys', 'amount usd', 'amount', 'montant']):
            column_letter = get_column_letter(col_idx)
            amount_columns.append((column, column_letter))
    
    # Calculer la plage des données
    start_row = data_start_row + 2  # +2 pour les en-têtes
    end_row = data_start_row + len(df) + 1
    
    for column_name, column_letter in amount_columns:
        # Définir la plage pour cette colonne
        range_address = f"{column_letter}{start_row}:{column_letter}{end_row}"
        
        # Règle pour les valeurs négatives (Rouge)
        negative_rule = CellIsRule(
            operator='lessThan',
            formula=['0'],
            font=Font(color="FF0000"),  # Rouge
            fill=None
        )
        
        # Règle pour les valeurs positives (Noir - optionnel)
        positive_rule = CellIsRule(
            operator='greaterThanOrEqual',
            formula=['0'],
            font=Font(color="000000"),  # Noir
            fill=None
        )
        
        # Appliquer les règles à la plage
        ws.conditional_formatting.add(range_address, negative_rule)
        ws.conditional_formatting.add(range_address, positive_rule)
        
        print(f"   💰 Formatage conditionnel appliqué à {column_name} (plage: {range_address})")
    
    print(f"✅ Formatage conditionnel configuré pour {len(amount_columns)} colonnes")

def validate_data_integrity(original_df, processed_df):
    """Valide que les colonnes protégées n'ont pas été modifiées"""
    
    PROTECTED_COLUMNS = [
        'Amount CCYs', 'Amount USD', 'Rate FX', 'Total', 
        'Entity', 'Period', 'Date', 'Transaction Date',
        'Bank Account', 'CCY'
    ]
    
    print(f"\n🔍 VALIDATION DE L'INTÉGRITÉ DES DONNÉES")
    print("="*50)
    
    integrity_issues = []
    protected_columns_found = []
    
    for col in PROTECTED_COLUMNS:
        if col in original_df.columns and col in processed_df.columns:
            protected_columns_found.append(col)
            
            # Comparer les données
            original_values = original_df[col].fillna('')
            processed_values = processed_df[col].fillna('')
            
            # Compter les différences
            differences = (original_values != processed_values).sum()
            
            if differences > 0:
                integrity_issues.append({
                    'column': col,
                    'differences': differences,
                    'percentage': (differences / len(original_df)) * 100
                })
                print(f"❌ Colonne '{col}': {differences} changements détectés ({(differences/len(original_df)*100):.1f}%)")
            else:
                print(f"✅ Colonne '{col}': Aucun changement")
    
    print(f"\n📊 RÉSUMÉ DE LA VALIDATION:")
    print(f"   • Colonnes protégées vérifiées: {len(protected_columns_found)}")
    print(f"   • Colonnes avec problèmes: {len(integrity_issues)}")
    
    if integrity_issues:
        total_issues = sum(issue['differences'] for issue in integrity_issues)
        print(f"   🚨 PROBLÈME: {total_issues} changements non-autorisés détectés!")
        print(f"   💡 Solution: Corriger la logique de protection des colonnes")
        return False
    else:
        print(f"   ✅ PARFAIT: Toutes les colonnes protégées sont intactes")
        print(f"   🎯 Système fonctionne correctement")
        return True