import requests
import os

# Configuration
API_URL = 'http://localhost:5001/api/excel'
TEST_FILE = 'test_data.xlsx'

def test_upload():
    """Test l'upload et le traitement d'un fichier Excel"""
    
    print("🧪 Test de l'API Excel Analyzer")
    print("=" * 50)
    
    # Vérifier que le fichier de test existe
    if not os.path.exists(TEST_FILE):
        print("❌ Fichier de test non trouvé:", TEST_FILE)
        return False
    
    print(f"📁 Fichier de test: {TEST_FILE}")
    
    try:
        # Test 1: Upload du fichier
        print("\n1️⃣ Test d'upload...")
        
        with open(TEST_FILE, 'rb') as f:
            files = {'file': f}
            response = requests.post(f'{API_URL}/upload', files=files)
        
        if response.status_code == 200:
            print("✅ Upload réussi!")
            data = response.json()
            print(f"   - Fichier original: {data['original_file']}")
            print(f"   - Fichier traité: {data['processed_file']}")
            print(f"   - Colonnes détectées: {len(data['columns_info']['columns'])}")
            print(f"   - Lignes: {data['columns_info']['shape'][0]}")
            
            # Test 2: Téléchargement du fichier traité
            print("\n2️⃣ Test de téléchargement...")
            
            download_response = requests.get(f"{API_URL}/download/{data['processed_file']}")
            
            if download_response.status_code == 200:
                print("✅ Téléchargement réussi!")
                
                # Sauvegarder le fichier téléchargé
                output_file = f"downloaded_{data['processed_file']}"
                with open(output_file, 'wb') as f:
                    f.write(download_response.content)
                print(f"   - Fichier sauvegardé: {output_file}")
                
                return True
            else:
                print(f"❌ Erreur de téléchargement: {download_response.status_code}")
                return False
        else:
            print(f"❌ Erreur d'upload: {response.status_code}")
            print(f"   Message: {response.text}")
            return False
            
    except requests.exceptions.ConnectionError:
        print("❌ Impossible de se connecter au serveur")
        print("   Vérifiez que le serveur Flask fonctionne sur le port 5001")
        return False
    except Exception as e:
        print(f"❌ Erreur inattendue: {e}")
        return False

if __name__ == "__main__":
    success = test_upload()
    print("\n" + "=" * 50)
    if success:
        print("🎉 Tous les tests sont passés avec succès!")
    else:
        print("💥 Certains tests ont échoué")

