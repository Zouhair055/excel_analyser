import pandas as pd
import os
from datetime import datetime
import tkinter as tk
from tkinter import filedialog

class TrainingVsOutputComparator:
    """Compare le fichier d'entraînement (vraies valeurs) avec la sortie du système"""
    
    def __init__(self):
        self.target_columns = ['Nature', 'Descrip', 'Vessel', 'Service', 'Reference']
        
    def select_files_gui(self):
        """Sélection des fichiers avec des labels corrects"""
        
        root = tk.Tk()
        root.withdraw()
        
        print("🎯 COMPARAISON ENTRAÎNEMENT vs SORTIE SYSTÈME")
        print("=" * 60)
        print("📋 Workflow:")
        print("   1. Fichier d'entraînement = vraies valeurs complètes")
        print("   2. Fichier de sortie = colonnes remplies par le système")
        print("   3. Comparaison = performance du système de règles")
        print()
        
        # Sélection fichier d'entraînement (vraies valeurs)
        print("📁 Sélection du fichier D'ENTRAÎNEMENT (vraies valeurs)...")
        training_file = filedialog.askopenfilename(
            title="📁 Fichier d'Entraînement (avec VRAIES valeurs)",
            filetypes=[("Fichiers Excel", "*.xlsx *.xls"), ("Tous les fichiers", "*.*")]
        )
        
        if not training_file:
            print("❌ Aucun fichier d'entraînement sélectionné")
            root.destroy()
            return None, None
        
        print(f"✅ Fichier d'entraînement: {os.path.basename(training_file)}")
        
        # Sélection fichier de sortie (rempli par le système)
        print("\n📁 Sélection du fichier de SORTIE (rempli par le système)...")
        output_file = filedialog.askopenfilename(
            title="📁 Fichier de Sortie (rempli par le système)",
            filetypes=[("Fichiers Excel", "*.xlsx *.xls"), ("Tous les fichiers", "*.*")]
        )
        
        if not output_file:
            print("❌ Aucun fichier de sortie sélectionné")
            root.destroy()
            return None, None
        
        print(f"✅ Fichier de sortie: {os.path.basename(output_file)}")
        
        root.destroy()
        return training_file, output_file
    
    def find_data_start_row(self, filepath):
        """Trouve où commencent les données (même logique que l'app)"""
        try:
            df_preview = pd.read_excel(filepath, header=None, nrows=15)
            
            for idx, row in df_preview.iterrows():
                row_str = ' '.join([str(cell) for cell in row if pd.notna(cell)]).lower()
                if 'description' in row_str:
                    return idx
            
            for idx, row in df_preview.iterrows():
                non_null_count = row.notna().sum()
                if non_null_count >= 5:
                    row_str = ' '.join([str(cell) for cell in row if pd.notna(cell)]).lower()
                    keywords = ['date', 'amount', 'entity', 'nature', 'vessel', 'service', 'period']
                    if any(keyword in row_str for keyword in keywords):
                        return idx
            return 0
        except Exception:
            return 0
    
    def read_excel_smart(self, filepath):
        """Lecture intelligente du fichier"""
        start_row = self.find_data_start_row(filepath)
        df = pd.read_excel(filepath, header=start_row)
        df.columns = [str(col).strip() if pd.notna(col) else f'Unnamed_{i}' for i, col in enumerate(df.columns)]
        df = df.dropna(how='all')
        return df
    
    def compare_training_vs_output(self, training_file, output_file):
        """Compare l'entraînement (vraies valeurs) avec la sortie du système"""
        
        print("\n" + "🎯 ANALYSE DE PERFORMANCE DU SYSTÈME" + "="*30)
        print(f"📚 Fichier d'entraînement: {os.path.basename(training_file)}")
        print(f"🤖 Fichier de sortie: {os.path.basename(output_file)}")
        
        try:
            # Charger les fichiers
            df_training = self.read_excel_smart(training_file)
            df_output = self.read_excel_smart(output_file)
            
            print(f"\n📊 Structure des fichiers:")
            print(f"   • Entraînement: {len(df_training)} lignes, {len(df_training.columns)} colonnes")
            print(f"   • Sortie: {len(df_output)} lignes, {len(df_output.columns)} colonnes")
            
            # Vérifier que les structures correspondent
            if len(df_training) != len(df_output):
                print("⚠️ Nombre de lignes différent - alignement des fichiers")
                min_rows = min(len(df_training), len(df_output))
                df_training = df_training.head(min_rows)
                df_output = df_output.head(min_rows)
            
            # Analyser chaque colonne cible
            self._analyze_system_performance(df_training, df_output)
            
            # Vérifier que les autres colonnes n'ont pas changé
            self._verify_unchanged_columns(df_training, df_output)
            
            # Rapport final de performance
            self._print_performance_report()
            
        except Exception as e:
            print(f"❌ Erreur: {e}")
            import traceback
            traceback.print_exc()
    
    def _analyze_system_performance(self, df_training, df_output):
        """Analyse la performance du système pour chaque colonne"""
        
        print("\n" + "🎯 PERFORMANCE PAR COLONNE CIBLE" + "="*35)
        
        self.performance_stats = {}
        self.total_correct = 0
        self.total_predicted = 0
        self.total_should_predict = 0
        
        for column in self.target_columns:
            print(f"\n📍 Analyse de la colonne: {column}")
            print("-" * 50)
            
            if column not in df_training.columns:
                print(f"❌ Colonne '{column}' manquante dans l'entraînement")
                continue
                
            if column not in df_output.columns:
                print(f"❌ Colonne '{column}' manquante dans la sortie")
                continue
            
            # Statistiques pour cette colonne
            correct_predictions = 0
            total_predictions = 0
            false_positives = 0
            missed_predictions = 0
            
            print(f"📊 Échantillon de comparaison (10 premières lignes):")
            print(f"{'Ligne':<6} {'Vraie Valeur':<25} {'Prédiction':<25} {'Résultat':<15}")
            print("-" * 75)
            
            for idx in range(min(len(df_training), 10)):
                true_val = self._normalize_value(df_training.iloc[idx][column])
                pred_val = self._normalize_value(df_output.iloc[idx][column])
                
                if true_val != '' and pred_val != '':
                    # Le système a prédit quelque chose pour une vraie valeur
                    if true_val == pred_val:
                        result = "✅ CORRECT"
                        correct_predictions += 1
                    else:
                        result = "❌ INCORRECT"
                        false_positives += 1
                    total_predictions += 1
                    print(f"{idx+1:<6} {true_val[:24]:<25} {pred_val[:24]:<25} {result:<15}")
                    
                elif true_val != '' and pred_val == '':
                    # Le système a raté une valeur qu'il aurait dû prédire
                    result = "⚠️ RATÉ"
                    missed_predictions += 1
                    if idx < 5:  # Afficher seulement les 5 premières
                        print(f"{idx+1:<6} {true_val[:24]:<25} {'(vide)':<25} {result:<15}")
                
                elif true_val == '' and pred_val != '':
                    # Le système a prédit là où il n'y avait rien
                    result = "🔄 NOUVEAU"
                    total_predictions += 1
                    if idx < 3:  # Afficher seulement les 3 premières
                        print(f"{idx+1:<6} {'(vide)':<25} {pred_val[:24]:<25} {result:<15}")
            
            # Calculer les statistiques complètes
            total_correct_for_col = 0
            total_predicted_for_col = 0
            total_should_predict_for_col = 0
            total_missed_for_col = 0
            
            for idx in range(len(df_training)):
                true_val = self._normalize_value(df_training.iloc[idx][column])
                pred_val = self._normalize_value(df_output.iloc[idx][column])
                
                if true_val != '':
                    total_should_predict_for_col += 1
                    
                if pred_val != '':
                    total_predicted_for_col += 1
                    
                if true_val != '' and pred_val == true_val:
                    total_correct_for_col += 1
                    
                if true_val != '' and pred_val == '':
                    total_missed_for_col += 1
            
            # Métriques de performance
            accuracy = (total_correct_for_col / total_should_predict_for_col * 100) if total_should_predict_for_col > 0 else 0
            precision = (total_correct_for_col / total_predicted_for_col * 100) if total_predicted_for_col > 0 else 0
            recall = (total_correct_for_col / total_should_predict_for_col * 100) if total_should_predict_for_col > 0 else 0
            
            print(f"\n📊 Métriques pour {column}:")
            print(f"   • Valeurs à prédire: {total_should_predict_for_col}")
            print(f"   • Prédictions faites: {total_predicted_for_col}")
            print(f"   • Prédictions correctes: {total_correct_for_col}")
            print(f"   • Valeurs ratées: {total_missed_for_col}")
            print(f"   • 🎯 Précision: {accuracy:.1f}% (bonnes réponses/total à prédire)")
            print(f"   • 🔍 Rappel: {recall:.1f}% (bonnes réponses/total à prédire)")
            print(f"   • ⚡ Efficacité: {precision:.1f}% (bonnes réponses/prédictions faites)")
            
            # Stocker les stats
            self.performance_stats[column] = {
                'total_should_predict': total_should_predict_for_col,
                'total_predicted': total_predicted_for_col,
                'correct_predictions': total_correct_for_col,
                'missed': total_missed_for_col,
                'accuracy': accuracy,
                'precision': precision,
                'recall': recall
            }
            
            # Mettre à jour les totaux
            self.total_correct += total_correct_for_col
            self.total_predicted += total_predicted_for_col
            self.total_should_predict += total_should_predict_for_col
    
    def _verify_unchanged_columns(self, df_training, df_output):
        """Vérifie que les autres colonnes n'ont pas changé"""
        
        print("\n" + "🔒 VÉRIFICATION DES AUTRES COLONNES" + "="*25)
        
        # Colonnes qui ne doivent pas changer
        common_cols = set(df_training.columns) & set(df_output.columns)
        other_cols = [col for col in common_cols if col not in self.target_columns]
        
        changes_detected = 0
        
        for col in other_cols:
            col_changes = 0
            for idx in range(len(df_training)):
                train_val = self._normalize_value(df_training.iloc[idx][col])
                out_val = self._normalize_value(df_output.iloc[idx][col])
                
                if train_val != out_val:
                    col_changes += 1
            
            if col_changes > 0:
                changes_detected += col_changes
                print(f"   ⚠️ {col}: {col_changes} changements détectés")
        
        if changes_detected == 0:
            print("   ✅ Parfait! Aucune colonne non-cible n'a été modifiée")
        else:
            print(f"   🚨 PROBLÈME: {changes_detected} changements dans les colonnes non-cibles")
        
        self.other_columns_stable = changes_detected == 0
    
    def _normalize_value(self, value):
        """Normalise une valeur"""
        if pd.isna(value) or value is None:
            return ''
        return str(value).strip()
    
    def _print_performance_report(self):
        """Rapport final de performance du système"""
        
        print("\n" + "="*80)
        print("📊 RAPPORT DE PERFORMANCE FINAL DU SYSTÈME")
        print("="*80)
        
        # Performance globale
        overall_accuracy = (self.total_correct / self.total_should_predict * 100) if self.total_should_predict > 0 else 0
        overall_precision = (self.total_correct / self.total_predicted * 100) if self.total_predicted > 0 else 0
        
        print(f"🎯 PERFORMANCE GLOBALE:")
        print(f"   • Total de valeurs à prédire: {self.total_should_predict:,}")
        print(f"   • Total de prédictions faites: {self.total_predicted:,}")
        print(f"   • Total de prédictions correctes: {self.total_correct:,}")
        print(f"   • 📊 Précision générale: {overall_accuracy:.1f}%")
        print(f"   • ⚡ Efficacité générale: {overall_precision:.1f}%")
        
        # Performance par colonne
        print(f"\n📋 PERFORMANCE PAR COLONNE:")
        print(f"{'Colonne':<12} {'À Prédire':<10} {'Prédites':<10} {'Correctes':<10} {'Précision':<10}")
        print("-" * 60)
        
        for col, stats in self.performance_stats.items():
            print(f"{col:<12} {stats['total_should_predict']:<10} {stats['total_predicted']:<10} {stats['correct_predictions']:<10} {stats['accuracy']:<9.1f}%")
        
        # Classement des colonnes par performance
        sorted_cols = sorted(self.performance_stats.items(), key=lambda x: x[1]['accuracy'], reverse=True)
        
        print(f"\n🏆 CLASSEMENT PAR PERFORMANCE:")
        for i, (col, stats) in enumerate(sorted_cols, 1):
            emoji = "🥇" if i == 1 else "🥈" if i == 2 else "🥉" if i == 3 else "📊"
            print(f"   {emoji} {i}. {col}: {stats['accuracy']:.1f}% de précision")
        
        # Stabilité des autres colonnes
        print(f"\n🔒 STABILITÉ DES AUTRES COLONNES:")
        if self.other_columns_stable:
            print("   ✅ Toutes les autres colonnes sont stables")
        else:
            print("   ❌ Certaines colonnes non-cibles ont été modifiées")
        
        # Verdict final et recommandations
        print(f"\n💡 VERDICT ET RECOMMANDATIONS:")
        
        if overall_accuracy >= 90:
            print("   🎉 EXCELLENT! Le système de règles fonctionne très bien")
            print("   ✅ Performance élevée, système prêt pour la production")
        elif overall_accuracy >= 70:
            print("   👍 BON! Performance satisfaisante")
            print("   🔧 Quelques ajustements de règles pourraient améliorer")
        elif overall_accuracy >= 50:
            print("   ⚠️ MOYEN! Performance modérée")
            print("   🛠️ Révision des règles recommandée")
        else:
            print("   🚨 FAIBLE! Performance insuffisante")
            print("   🔄 Révision complète du système de règles nécessaire")
        
        # Recommandations spécifiques
        worst_col = min(sorted_cols, key=lambda x: x[1]['accuracy'])
        best_col = max(sorted_cols, key=lambda x: x[1]['accuracy'])
        
        print(f"\n🎯 ACTIONS PRIORITAIRES:")
        print(f"   📈 Améliorer les règles pour '{worst_col[0]}' ({worst_col[1]['accuracy']:.1f}%)")
        print(f"   ✨ Reproduire le succès de '{best_col[0]}' ({best_col[1]['accuracy']:.1f}%)")
        
        if not self.other_columns_stable:
            print(f"   🔧 Corriger les modifications non-désirées dans les autres colonnes")
        
        print("="*80)

def main():
    """Fonction principale"""
    
    comparator = TrainingVsOutputComparator()
    
    # Sélection des fichiers
    training_file, output_file = comparator.select_files_gui()
    
    if training_file and output_file:
        # Analyser la performance du système
        comparator.compare_training_vs_output(training_file, output_file)
        
        print(f"\n🎉 Analyse de performance terminée!")
        input("\nAppuyez sur Entrée pour fermer...")
    else:
        print("❌ Analyse annulée - fichiers non sélectionnés")

if __name__ == "__main__":
    main()