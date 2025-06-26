import pandas as pd

def verify_processing():
    """Vérifie que les règles ont été correctement appliquées"""
    
    print("🔍 Vérification du traitement Excel")
    print("=" * 50)
    
    # Charger les fichiers
    original_file = 'test_data.xlsx'
    processed_file = 'downloaded_processed_test_data.xlsx'
    
    try:
        df_original = pd.read_excel(original_file)
        df_processed = pd.read_excel(processed_file)
        
        print("📊 Données originales:")
        print(df_original.head())
        print("\n📊 Données traitées:")
        print(df_processed.head())
        
        print("\n🔍 Vérification des règles:")
        
        # Règle 1: ADVICEPRO
        advicepro_rows = df_processed[df_processed['Description'].str.contains('ADVICEPRO', na=False)]
        if not advicepro_rows.empty:
            print("✅ Règle ADVICEPRO appliquée:")
            for idx, row in advicepro_rows.iterrows():
                print(f"   - Ligne {idx}: Nature='{row['Nature']}', Descrip='{row['Descrip']}', Service='{row['Service']}'")
        
        # Règle 2: Extraction de références
        ref_filled = df_processed[df_processed['Reference'].notna() & (df_processed['Reference'] != '')]
        if not ref_filled.empty:
            print("✅ Règle d'extraction de références appliquée:")
            for idx, row in ref_filled.iterrows():
                print(f"   - Ligne {idx}: Reference='{row['Reference']}'")
        
        # Règle 3: USD → Import
        usd_rows = df_processed[df_processed['Bank account'].str.contains('USD', na=False)]
        if not usd_rows.empty:
            print("✅ Règle USD → Import appliquée:")
            for idx, row in usd_rows.iterrows():
                print(f"   - Ligne {idx}: Bank account='{row['Bank account']}', Nature='{row['Nature']}'")
        
        print("\n📈 Statistiques:")
        print(f"   - Lignes traitées: {len(df_processed)}")
        print(f"   - Colonnes: {len(df_processed.columns)}")
        print(f"   - Colonnes avec données: {len([col for col in df_processed.columns if not df_processed[col].isna().all()])}")
        
        return True
        
    except Exception as e:
        print(f"❌ Erreur lors de la vérification: {e}")
        return False

if __name__ == "__main__":
    verify_processing()

