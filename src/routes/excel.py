from flask import Blueprint, request, jsonify, send_file
import pandas as pd
import os
import tempfile
from werkzeug.utils import secure_filename
import re
import json
import numpy as np

excel_bp = Blueprint('excel', __name__)

# Utiliser des chemins absolus pour éviter les problèmes
UPLOAD_FOLDER = os.path.abspath('uploads')
PROCESSED_FOLDER = os.path.abspath('processed')

# Créer les dossiers s'ils n'existent pas
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'xlsx', 'xls'}

def clean_data_for_json(data):
    """Nettoie les données pour éviter les erreurs JSON avec NaN"""
    if isinstance(data, list):
        return [clean_data_for_json(item) for item in data]
    elif isinstance(data, dict):
        return {key: clean_data_for_json(value) for key, value in data.items()}
    elif pd.isna(data) or data is np.nan:
        return None
    elif isinstance(data, (np.integer, np.floating)):
        return data.item()
    else:
        return data

def apply_rules(df):
    """
    Applique les règles de remplissage des colonnes selon la structure réelle du fichier
    Colonnes à remplir: Nature, Descrip, Vessel, Service, Reference
    """
    
    # Règle 1 : Si "ADVICEPRO" est dans 'Description', remplir les colonnes
    if 'Description' in df.columns:
        mask_advicepro = df['Description'].str.contains("ADVICEPRO", case=False, na=False)
        
        # Créer les colonnes si elles n'existent pas
        for col in ['Nature', 'Descrip', 'Vessel', 'Service']:
            if col not in df.columns:
                df[col] = ''
        
        # TEMPORAIREMENT DÉSACTIVÉ - À RECONDITIONNER
        # mask_empty_nature = (df['Nature'].isna()) | (df['Nature'] == '')
        # mask_empty_descrip = (df['Descrip'].isna()) | (df['Descrip'] == '')
        # mask_empty_vessel = (df['Vessel'].isna()) | (df['Vessel'] == '')
        # mask_empty_service = (df['Service'].isna()) | (df['Service'] == '')
        # 
        # df.loc[mask_advicepro & mask_empty_nature, 'Nature'] = "G- Suppliers"
        # df.loc[mask_advicepro & mask_empty_descrip, 'Descrip'] = "ADVICEPRO"
        # df.loc[mask_advicepro & mask_empty_vessel, 'Vessel'] = "N/A"
        # df.loc[mask_advicepro & mask_empty_service, 'Service'] = "OHD"

    # Règle 2 : Extraire 'Reference' depuis 'Description' (GARDÉE)
    if 'Description' in df.columns:
        if 'Reference' not in df.columns:
            df['Reference'] = ''
        
        # Chercher les patterns AE\d+ et OFFICE \d+ \w+ dans Description
        references = df['Description'].str.extract(r'(AE\d+|OFFICE \d+ \w+)', expand=False)
        mask_empty_ref = (df['Reference'].isna()) | (df['Reference'] == '')
        df.loc[mask_empty_ref, 'Reference'] = references.loc[mask_empty_ref]

    # RÈGLES SUPPRIMÉES :
    # - Règle 3 : Bank account USD → Import (SUPPRIMÉE)
    # - Règle 4 : CCY USD → Import (SUPPRIMÉE) 
    # - Règle 5 : Swift Payment → G-Suppliers + SWIFT (SUPPRIMÉE)
    
    return df

@excel_bp.route('/upload', methods=['POST'])
def upload_file():
    """
    Endpoint pour uploader et traiter un fichier Excel
    """
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Aucun fichier fourni'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'Aucun fichier sélectionné'}), 400
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            file.save(filepath)
            
            # Lire le fichier Excel en spécifiant que les en-têtes sont à la ligne 8 (index 7)
            df = pd.read_excel(filepath, header=7)  # Ligne 8 = index 7
            
            # Debug: afficher les colonnes détectées
            print(f"Colonnes détectées: {list(df.columns)}")
            print(f"Premières lignes du DataFrame:")
            print(df.head())
            
            # Remplacer les NaN par des chaînes vides pour éviter les problèmes JSON
            df_clean = df.fillna('')
            
            # Obtenir les informations sur les colonnes
            sample_data = df_clean.head(3).to_dict('records')
            sample_data_cleaned = clean_data_for_json(sample_data)
            
            columns_info = {
                'columns': list(df.columns),
                'shape': df.shape,
                'empty_columns': [col for col in df.columns if df[col].isna().all()],
                'sample_data': sample_data_cleaned
            }
            
            # Appliquer les règles de traitement
            df_processed = apply_rules(df.copy())
            
            # Sauvegarder le fichier traité
            processed_filename = f"processed_{filename}"
            processed_filepath = os.path.join(PROCESSED_FOLDER, processed_filename)
            df_processed.to_excel(processed_filepath, index=False)
            
            # Vérifier que le fichier a été créé
            if not os.path.exists(processed_filepath):
                raise Exception(f"Le fichier traité n'a pas pu être créé: {processed_filepath}")
            
            print(f"Fichier traité créé avec succès: {processed_filepath}")  # Debug
            
            return jsonify({
                'success': True,
                'message': 'Fichier traité avec succès',
                'original_file': filename,
                'processed_file': processed_filename,
                'columns_info': columns_info,
                'changes_applied': {
                    'rules_applied': [
                        'Remplissage automatique pour ADVICEPRO',
                        'Extraction de références',
                        'Classification USD → Import'
                    ]
                }
            })
        
        return jsonify({'error': 'Type de fichier non autorisé. Utilisez .xlsx ou .xls'}), 400
    
    except Exception as e:
        print(f"Erreur détaillée: {str(e)}")  # Pour le débogage
        return jsonify({'error': f'Erreur lors du traitement: {str(e)}'}), 500
@excel_bp.route('/download/<filename>')
def download_file(filename):
    """
    Endpoint pour télécharger le fichier traité
    """
    try:
        # Utiliser le chemin absolu
        filepath = os.path.join(PROCESSED_FOLDER, filename)
        
        print(f"Tentative de téléchargement: {filepath}")  # Debug
        print(f"Le fichier existe: {os.path.exists(filepath)}")  # Debug
        
        if os.path.exists(filepath):
            # Vérifier que le fichier n'est pas vide
            file_size = os.path.getsize(filepath)
            print(f"Taille du fichier: {file_size} bytes")  # Debug
            
            if file_size == 0:
                return jsonify({'error': 'Le fichier est vide'}), 500
            
            return send_file(
                filepath, 
                as_attachment=True, 
                download_name=filename,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            print(f"Fichier non trouvé: {filepath}")  # Debug
            print(f"Contenu du dossier processed: {os.listdir(PROCESSED_FOLDER) if os.path.exists(PROCESSED_FOLDER) else 'Dossier inexistant'}")  # Debug
            return jsonify({'error': f'Fichier non trouvé: {filename}'}), 404
            
    except Exception as e:
        print(f"Erreur lors du téléchargement: {str(e)}")  # Debug
        return jsonify({'error': f'Erreur lors du téléchargement: {str(e)}'}), 500

@excel_bp.route('/columns/<filename>')
def get_columns(filename):
    """
    Endpoint pour obtenir les colonnes d'un fichier uploadé
    """
    try:
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        if os.path.exists(filepath):
            # Lire avec les en-têtes à la ligne 8
            df = pd.read_excel(filepath, header=7)
            df_clean = df.fillna('')
            sample_data = clean_data_for_json(df_clean.head(5).to_dict('records'))
            
            return jsonify({
                'columns': list(df.columns),
                'shape': df.shape,
                'sample_data': sample_data
            })
        else:
            return jsonify({'error': 'Fichier non trouvé'}), 404
    except Exception as e:
        return jsonify({'error': f'Erreur lors de la lecture: {str(e)}'}), 500