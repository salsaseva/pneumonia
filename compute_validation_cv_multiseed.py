"""
Script untuk menghitung secara komprehensif:
1. Metrik Hasil Validasi Tetap (Validation Set P2: 1.047 citra)
2. CV 5-Fold Berkelompok per Pasien (StratifiedGroupKFold, n_splits=5) pada Data Latih Resmi (5.232 citra) -> Mean ± SD
3. Multi-Seed Stability Analysis (5 seed: 42, 123, 456, 789, 2026)
4. Tabel Berdampingan: Validasi vs CV 5-Fold vs Uji Resmi (624 citra)
"""
import os
import json
import numpy as np
import pandas as pd
from sklearn.neighbors import KNeighborsClassifier
from sklearn.svm import SVC
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import StratifiedGroupKFold
from imblearn.over_sampling import SMOTE

from src.config import RANDOM_SEED, ADDITIONAL_SEEDS, TABLES_DIR, CACHE_DIR
from src.features import extract_features_dataset
from src.pipelines import build_pipeline, create_voting_ensemble
from src.evaluation import compute_all_metrics

ALL_SEEDS = [RANDOM_SEED] + ADDITIONAL_SEEDS  # [42, 123, 456, 789, 2026]

def run():
    print("=" * 70)
    print("MENGHITUNG METRIK VALIDASI, CV 5-FOLD, MULTI-SEED, & TABEL BERDAMPINGAN")
    print("=" * 70)

    # 1. Load Data Inventory & Splits
    df_inventory = pd.read_csv("results/audit/dataset_inventory.csv")
    with open("results/splits/splits_seed_42.json", "r") as f:
        splits_42 = json.load(f)

    # Load baseline cached features (basic, Ng=256, GLCM-6)
    X_all, y_all = extract_features_dataset(df_inventory, mode='basic', Ng=256, variant='GLCM-6', use_cache=True)
    patient_ids = df_inventory['patient_id'].values

    # P2 Indices
    p2_tr_idx = np.array(splits_42['P2']['train_idx'])
    p2_val_idx = np.array(splits_42['P2']['val_idx'])
    p2_test_idx = np.array(splits_42['P2']['test_idx'])

    # Data latih resmi utuh (train + val = 5.232 citra)
    p2_train_full_idx = np.concatenate([p2_tr_idx, p2_val_idx])

    X_train_sub = X_all[p2_tr_idx]
    y_train_sub = y_all[p2_tr_idx]

    X_val = X_all[p2_val_idx]
    y_val = y_all[p2_val_idx]

    X_train_full = X_all[p2_train_full_idx]
    y_train_full = y_all[p2_train_full_idx]
    patient_ids_full = patient_ids[p2_train_full_idx]

    X_test = X_all[p2_test_idx]
    y_test = y_all[p2_test_idx]

    print(f"Data Latih Sub-Split: {len(X_train_sub)} | Validasi: {len(X_val)} | Uji Resmi: {len(X_test)}")
    print(f"Data Latih Resmi Utuh: {len(X_train_full)} citra")

    # Define model configurations
    models_config = {
        'KNN (Tuned: k=11)': KNeighborsClassifier(n_neighbors=11, metric='euclidean'),
        'SVM (Tuned: RBF, C=100)': SVC(C=100.0, kernel='rbf', gamma='scale', probability=True, random_state=42),
        'Random Forest (Tuned)': RandomForestClassifier(n_estimators=300, max_depth=None, min_samples_leaf=2, random_state=42),
        'Ensemble Soft Voting': None, # will be created from tuned
        'SVM Default [15] (Linear)': SVC(C=1.0, kernel='linear', probability=True, random_state=42),
        'RF Default [15] (100 pohon)': RandomForestClassifier(n_estimators=100, min_samples_leaf=2, random_state=42),
        'KNN Default [15] (k=65)': KNeighborsClassifier(n_neighbors=65, metric='euclidean'),
        'SVM + SMOTE': SVC(C=100.0, kernel='rbf', gamma='scale', probability=True, random_state=42),
        'Voting + SMOTE': None
    }

    # -------------------------------------------------------------------------
    # A. METRIK VALIDASI TETAP (1.047 CITRA)
    # -------------------------------------------------------------------------
    print("\n>>> 1. Menghitung Metrik pada Validasi Tetap (P2, 1.047 Citra)...")
    val_results = []
    trained_sub_models = {}

    for name, clf in models_config.items():
        use_smote = ('SMOTE' in name)
        if 'Voting' in name:
            # build voting from tuned models
            s = trained_sub_models['SVM (Tuned: RBF, C=100)'].named_steps['model']
            r = trained_sub_models['Random Forest (Tuned)'].named_steps['model']
            k = trained_sub_models['KNN (Tuned: k=11)'].named_steps['model']
            v_clf = create_voting_ensemble(s, r, k)
            pipe = build_pipeline(v_clf, scaler=True, smote=use_smote)
        else:
            pipe = build_pipeline(clf, scaler=True, smote=use_smote)
        
        pipe.fit(X_train_sub, y_train_sub)
        trained_sub_models[name] = pipe

        val_pred = pipe.predict(X_val)
        val_proba = pipe.predict_proba(X_val)[:, 1] if hasattr(pipe, 'predict_proba') else None
        m = compute_all_metrics(y_val, val_pred, val_proba)
        m['Model'] = name
        val_results.append(m)

    df_val = pd.DataFrame(val_results)
    val_csv_path = os.path.join(TABLES_DIR, "tabel_hasil_validasi_tetap.csv")
    df_val.to_csv(val_csv_path, index=False)
    print(f"Hasil validasi tetap disimpan ke {val_csv_path}")

    # -------------------------------------------------------------------------
    # B. CV 5-FOLD BERKELOMPOK PER PASIEN (StratifiedGroupKFold pada 5.232 Citra)
    # -------------------------------------------------------------------------
    print("\n>>> 2. Menjalankan CV 5-Fold Berkelompok per Pasien pada Data Latih Resmi (5.232 Citra)...")
    sgkf = StratifiedGroupKFold(n_splits=5, shuffle=True, random_state=42)

    cv_results = {name: {'acc': [], 'f1': [], 'sen': [], 'spe': [], 'bal_acc': []} for name in models_config.keys()}

    fold_idx = 0
    for train_idx, val_fold_idx in sgkf.split(X_train_full, y_train_full, groups=patient_ids_full):
        fold_idx += 1
        X_tr_f, y_tr_f = X_train_full[train_idx], y_train_full[train_idx]
        X_va_f, y_va_f = X_train_full[val_fold_idx], y_train_full[val_fold_idx]

        # Verify zero patient overlap in fold
        p_tr = set(patient_ids_full[train_idx])
        p_va = set(patient_ids_full[val_fold_idx])
        assert len(p_tr.intersection(p_va)) == 0, f"Patient leak in fold {fold_idx}!"

        fold_trained = {}
        for name, clf in models_config.items():
            use_smote = ('SMOTE' in name)
            if 'Voting' in name:
                s = fold_trained['SVM (Tuned: RBF, C=100)'].named_steps['model']
                r = fold_trained['Random Forest (Tuned)'].named_steps['model']
                k = fold_trained['KNN (Tuned: k=11)'].named_steps['model']
                v_clf = create_voting_ensemble(s, r, k)
                pipe = build_pipeline(v_clf, scaler=True, smote=use_smote)
            else:
                pipe = build_pipeline(clf, scaler=True, smote=use_smote)

            pipe.fit(X_tr_f, y_tr_f)
            fold_trained[name] = pipe

            pred = pipe.predict(X_va_f)
            m = compute_all_metrics(y_va_f, pred)

            cv_results[name]['acc'].append(m['Accuracy (%)'])
            cv_results[name]['f1'].append(m['F1-Score (%)'])
            cv_results[name]['sen'].append(m['Recall/Sen (%)'])
            cv_results[name]['spe'].append(m['Specificity (%)'])
            cv_results[name]['bal_acc'].append(m['Balanced Accuracy (%)'])

    cv_summary = []
    for name, metrics in cv_results.items():
        cv_summary.append({
            'Model': name,
            'CV Acc Mean (%)': np.mean(metrics['acc']),
            'CV Acc SD (%)': np.std(metrics['acc']),
            'CV F1 Mean (%)': np.mean(metrics['f1']),
            'CV F1 SD (%)': np.std(metrics['f1']),
            'CV Sen Mean (%)': np.mean(metrics['sen']),
            'CV Sen SD (%)': np.std(metrics['sen']),
            'CV Spe Mean (%)': np.mean(metrics['spe']),
            'CV Spe SD (%)': np.std(metrics['spe']),
            'CV Bal Acc Mean (%)': np.mean(metrics['bal_acc']),
            'CV Bal Acc SD (%)': np.std(metrics['bal_acc']),
        })

    df_cv = pd.DataFrame(cv_summary)
    cv_csv_path = os.path.join(TABLES_DIR, "tabel_cv_5fold_berkelompok.csv")
    df_cv.to_csv(cv_csv_path, index=False)
    print(f"Hasil CV 5-fold berkelompok disimpan ke {cv_csv_path}")

    # -------------------------------------------------------------------------
    # C. MULTI-SEED STABILITY ANALYSIS (5 SEED: 42, 123, 456, 789, 2026)
    # -------------------------------------------------------------------------
    print(f"\n>>> 3. Menjalankan Analisis Stabilitas Multi-Seed ({ALL_SEEDS})...")
    seed_records = []

    for seed in ALL_SEEDS:
        # Generate split for this seed on P2 (or train/val split on 5232)
        sgkf_seed = StratifiedGroupKFold(n_splits=5, shuffle=True, random_state=seed)
        tr_i, va_i = next(sgkf_seed.split(X_train_full, y_train_full, groups=patient_ids_full))

        X_tr_s, y_tr_s = X_train_full[tr_i], y_train_full[tr_i]
        X_va_s, y_va_s = X_train_full[va_i], y_train_full[va_i]

        # Evaluate Tuned models and Voting on Val and Test
        # 1. Fit models on train split
        m_knn = build_pipeline(KNeighborsClassifier(n_neighbors=11), scaler=True)
        m_svm = build_pipeline(SVC(C=100.0, kernel='rbf', gamma='scale', probability=True, random_state=seed), scaler=True)
        m_rf = build_pipeline(RandomForestClassifier(n_estimators=300, min_samples_leaf=2, random_state=seed), scaler=True)

        m_knn.fit(X_tr_s, y_tr_s)
        m_svm.fit(X_tr_s, y_tr_s)
        m_rf.fit(X_tr_s, y_tr_s)

        v_pipe = build_pipeline(create_voting_ensemble(
            m_svm.named_steps['model'], m_rf.named_steps['model'], m_knn.named_steps['model']
        ), scaler=True)
        v_pipe.fit(X_tr_s, y_tr_s)

        for m_name, pipe_obj in [('KNN (Tuned)', m_knn), ('SVM (Tuned)', m_svm), ('Random Forest (Tuned)', m_rf), ('Ensemble Soft Voting', v_pipe)]:
            # Score on validation
            val_p = pipe_obj.predict(X_va_s)
            val_m = compute_all_metrics(y_va_s, val_p)

            # Score on official test (once)
            test_p = pipe_obj.predict(X_test)
            test_m = compute_all_metrics(y_test, test_p)

            seed_records.append({
                'Seed': seed,
                'Model': m_name,
                'Val Acc (%)': val_m['Accuracy (%)'],
                'Val F1 (%)': val_m['F1-Score (%)'],
                'Val Sen (%)': val_m['Recall/Sen (%)'],
                'Val Spe (%)': val_m['Specificity (%)'],
                'Test Acc (%)': test_m['Accuracy (%)'],
                'Test F1 (%)': test_m['F1-Score (%)'],
                'Test Sen (%)': test_m['Recall/Sen (%)'],
                'Test Spe (%)': test_m['Specificity (%)'],
            })

    df_seeds = pd.DataFrame(seed_records)
    seeds_csv_path = os.path.join(TABLES_DIR, "tabel_multi_seed_stability.csv")
    df_seeds.to_csv(seeds_csv_path, index=False)
    print(f"Hasil multi-seed stability disimpan ke {seeds_csv_path}")

    # Multi-seed aggregate summary
    seed_agg = df_seeds.groupby('Model').agg({
        'Val Acc (%)': ['mean', 'std'],
        'Val F1 (%)': ['mean', 'std'],
        'Test Acc (%)': ['mean', 'std'],
        'Test F1 (%)': ['mean', 'std'],
    }).reset_index()
    seed_agg.columns = ['Model', 'Val Acc Mean', 'Val Acc SD', 'Val F1 Mean', 'Val F1 SD', 'Test Acc Mean', 'Test Acc SD', 'Test F1 Mean', 'Test F1 SD']
    seed_agg_path = os.path.join(TABLES_DIR, "tabel_multi_seed_summary.csv")
    seed_agg.to_csv(seed_agg_path, index=False)

    # -------------------------------------------------------------------------
    # D. TABEL BERDAMPINGAN: VALIDASI vs CV 5-FOLD vs UJI RESMI
    # -------------------------------------------------------------------------
    print("\n>>> 4. Membuat Tabel Berdampingan: Validasi vs CV 5-Fold vs Uji Resmi...")
    df_test = pd.read_csv(os.path.join(TABLES_DIR, "tabel_evaluasi_model_lengkap.csv"))
    # Match model names
    side_by_side = []
    for _, row_val in df_val.iterrows():
        m_name = row_val['Model']
        
        # Find CV row
        cv_match = df_cv[df_cv['Model'] == m_name]
        cv_acc_str = f"{cv_match['CV Acc Mean (%)'].values[0]:.2f} ± {cv_match['CV Acc SD (%)'].values[0]:.2f}" if len(cv_match) > 0 else "-"
        cv_f1_str = f"{cv_match['CV F1 Mean (%)'].values[0]:.2f} ± {cv_match['CV F1 SD (%)'].values[0]:.2f}" if len(cv_match) > 0 else "-"

        # Find Test row
        # clean m_name for matching
        test_match = df_test[df_test['Model'].str.contains(m_name.split(' (')[0], case=False, regex=False)]
        if 'SMOTE' in m_name:
            test_match = df_test[df_test['Model'].str.contains('SMOTE', case=False) & df_test['Model'].str.contains(m_name.split(' ')[0], case=False)]
        elif 'Default' in m_name:
            test_match = df_test[df_test['Model'].str.contains('Default', case=False) & df_test['Model'].str.contains(m_name.split(' ')[0], case=False)]

        test_acc = f"{test_match['Accuracy (%)'].values[0]:.2f}%" if len(test_match) > 0 else "-"
        test_f1 = f"{test_match['F1-Score (%)'].values[0]:.2f}%" if len(test_match) > 0 else "-"
        test_spe = f"{test_match['Specificity (%)'].values[0]:.2f}%" if len(test_match) > 0 else "-"

        side_by_side.append({
            'Model': m_name,
            'Val Acc (%)': f"{row_val['Accuracy (%)']:.2f}%",
            'Val F1 (%)': f"{row_val['F1-Score (%)']:.2f}%",
            'Val Spe (%)': f"{row_val['Specificity (%)']:.2f}%",
            'CV 5-Fold Acc (Mean ± SD)': cv_acc_str,
            'CV 5-Fold F1 (Mean ± SD)': cv_f1_str,
            'Test Acc (%)': test_acc,
            'Test F1 (%)': test_f1,
            'Test Spe (%)': test_spe,
        })

    df_side = pd.DataFrame(side_by_side)
    side_csv_path = os.path.join(TABLES_DIR, "tabel_validasi_cv_uji_berdampingan.csv")
    df_side.to_csv(side_csv_path, index=False)
    print(f"Tabel berdampingan disimpan ke {side_csv_path}")

    print("\n" + "=" * 70)
    print("RINGKASAN HASIL:")
    print("=" * 70)
    print("\n1. TABEL BERDAMPINGAN (Validasi vs CV 5-Fold vs Uji):")
    print(df_side.to_string(index=False))
    print("\n2. MULTI-SEED SUMMARY (5 Seed: Mean ± SD):")
    print(seed_agg.to_string(index=False))

if __name__ == '__main__':
    run()
