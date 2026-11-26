"""
Master Experiment Runner: Deteksi Pneumonia GLCM & Ensemble
Menjalankan seluruh Tahap 1 hingga Tahap 10 secara terstruktur, reproducible,
dan mengekspor seluruh hasil tabel, figur, dan model ke results/.
"""
import warnings
warnings.filterwarnings('ignore')

import os
import sys
import time
import json
import joblib
import platform
import numpy as np
import pandas as pd
import sklearn
import cv2
import skimage
import imblearn

from src.config import (
    BASE_DIR, DATASET_DIR, RESULTS_DIR, AUDIT_DIR, TABLES_DIR,
    FIGURES_DIR, MODELS_DIR, RANDOM_SEED, ADDITIONAL_SEEDS
)
from src.audit import run_audit
from src.splits import generate_splits
from src.preprocessing import compute_reference_histogram
from src.features import extract_features_dataset, GLCM_PROPERTIES
from src.pipelines import (
    build_pipeline, get_default_models, get_legacy_models,
    get_tuning_hyperparameters, create_voting_ensemble
)
from src.evaluation import (
    compute_all_metrics, plot_confusion_matrices, plot_roc_and_det_curves
)
from src.statistics import (
    run_all_mcnemar_tests, compute_bootstrap_ci, analyze_ensemble_diversity
)
from src.ablation import (
    test_artifact_robustness, compute_permutation_importance
)

from sklearn.neighbors import KNeighborsClassifier
from sklearn.svm import SVC
from sklearn.ensemble import RandomForestClassifier

def main():
    t_start = time.time()
    print("=" * 80)
    print("SISTEM EVALUASI & REVISI ILMIAH DETEKSI PNEUMONIA GLCM")
    print("Pembanding Utama: Ortiz-Toro et al. [15] | Studi Pendahulu: [19]")
    print("=" * 80)

    # -------------------------------------------------------------------------
    # TAHAP 1: Audit Data & ID Pasien
    # -------------------------------------------------------------------------
    print("\n>>> TAHAP 1: Audit Data & Ekstraksi ID Pasien...")
    df_inventory, audit_summary = run_audit(DATASET_DIR, AUDIT_DIR)
    print(f"Total citra: {len(df_inventory)} | ID Overlap Train-Test: {audit_summary['overlapping_patients_count']}")

    # -------------------------------------------------------------------------
    # TAHAP 2: Pembuatan 3 Protokol Pembagian Data (P1, P2, P3)
    # -------------------------------------------------------------------------
    print("\n>>> TAHAP 2: Menghasilkan Protokol Pembagian Data P1, P2, P3...")
    splits = generate_splits(df_inventory, seed=RANDOM_SEED)

    # -------------------------------------------------------------------------
    # TAHAP 3 & 4: Prapengolahan & Ekstraksi Fitur GLCM Baseline
    # -------------------------------------------------------------------------
    print("\n>>> TAHAP 3 & 4: Prapengolahan & Ekstraksi Fitur GLCM...")
    # Hitung histogram referensi HANYA dari data latih P2
    p2_train_paths = df_inventory.iloc[splits['P2']['train_idx']]['file_path'].tolist()
    ref_cdf = compute_reference_histogram(p2_train_paths)

    # Ekstraksi baseline GLCM-6, Ng=256, basic
    X_all, y_all = extract_features_dataset(
        df_inventory, mode='basic', Ng=256, variant='GLCM-6', ref_cdf=ref_cdf
    )

    # -------------------------------------------------------------------------
    # TAHAP 5 & 6: Pipeline & Tuning Model pada Validasi P2
    # -------------------------------------------------------------------------
    print("\n>>> TAHAP 5 & 6: Tuning Hyperparameter pada Set Validasi P2...")
    p2_tr_idx = splits['P2']['train_idx']
    p2_val_idx = splits['P2']['val_idx']
    p2_te_idx = splits['P2']['test_idx']

    X_train, y_train = X_all[p2_tr_idx], y_all[p2_tr_idx]
    X_val, y_val = X_all[p2_val_idx], y_all[p2_val_idx]
    X_test, y_test = X_all[p2_te_idx], y_all[p2_te_idx]

    tuning_space = get_tuning_hyperparameters(n_train_samples=len(X_train))

    # 1. Tuning KNN
    print("  -> Tuning KNN...")
    best_knn_params, best_knn_f1 = None, -1
    for params in tuning_space['KNN']:
        pipe = build_pipeline(KNeighborsClassifier(**params), scaler=True, smote=False)
        pipe.fit(X_train, y_train)
        val_pred = pipe.predict(X_val)
        val_metrics = compute_all_metrics(y_val, val_pred)
        if val_metrics['F1-Score (%)'] > best_knn_f1:
            best_knn_f1 = val_metrics['F1-Score (%)']
            best_knn_params = params

    # 2. Tuning SVM
    print("  -> Tuning SVM...")
    best_svm_params, best_svm_f1 = None, -1
    for params in tuning_space['SVM']:
        pipe = build_pipeline(SVC(**params, random_state=RANDOM_SEED), scaler=True, smote=False)
        pipe.fit(X_train, y_train)
        val_pred = pipe.predict(X_val)
        val_metrics = compute_all_metrics(y_val, val_pred)
        if val_metrics['F1-Score (%)'] > best_svm_f1:
            best_svm_f1 = val_metrics['F1-Score (%)']
            best_svm_params = params

    # 3. Tuning RF
    print("  -> Tuning Random Forest...")
    best_rf_params, best_rf_f1 = None, -1
    for params in tuning_space['RF']:
        pipe = build_pipeline(RandomForestClassifier(**params, random_state=RANDOM_SEED), scaler=True, smote=False)
        pipe.fit(X_train, y_train)
        val_pred = pipe.predict(X_val)
        val_metrics = compute_all_metrics(y_val, val_pred)
        if val_metrics['F1-Score (%)'] > best_rf_f1:
            best_rf_f1 = val_metrics['F1-Score (%)']
            best_rf_params = params

    print(f"  Best KNN: {best_knn_params} (Val F1: {best_knn_f1}%)")
    print(f"  Best SVM: {best_svm_params} (Val F1: {best_svm_f1}%)")
    print(f"  Best RF:  {best_rf_params} (Val F1: {best_rf_f1}%)")

    # Fit final models on P2 Train+Val (Official 5232 images) or evaluate directly
    # Menurut panduan: latih di train (atau train+val), evaluasi sekali jalan di test
    X_train_full = np.vstack([X_train, X_val])
    y_train_full = np.concatenate([y_train, y_val])

    tuned_knn = KNeighborsClassifier(**best_knn_params)
    tuned_svm = SVC(**best_svm_params, random_state=RANDOM_SEED)
    tuned_rf = RandomForestClassifier(**best_rf_params, random_state=RANDOM_SEED)
    tuned_voting = create_voting_ensemble(tuned_svm, tuned_rf, tuned_knn)

    models_to_evaluate = {
        'KNN (Tuned)': build_pipeline(tuned_knn, scaler=True, smote=False),
        'SVM (Tuned)': build_pipeline(tuned_svm, scaler=True, smote=False),
        'Random Forest (Tuned)': build_pipeline(tuned_rf, scaler=True, smote=False),
        'Ensemble Soft Voting': build_pipeline(tuned_voting, scaler=True, smote=False)
    }

    # Model default ala [15]
    default_dict = get_default_models(n_train_samples=len(X_train_full))
    for name, m in default_dict.items():
        models_to_evaluate[f"{name} [15]"] = build_pipeline(m, scaler=True, smote=False)

    # Model dengan SMOTE
    models_to_evaluate['SVM (SMOTE)'] = build_pipeline(SVC(**best_svm_params, random_state=RANDOM_SEED), scaler=True, smote=True)
    models_to_evaluate['RF (SMOTE)'] = build_pipeline(RandomForestClassifier(**best_rf_params, random_state=RANDOM_SEED), scaler=True, smote=True)
    models_to_evaluate['Voting (SMOTE)'] = build_pipeline(create_voting_ensemble(
        SVC(**best_svm_params, random_state=RANDOM_SEED),
        RandomForestClassifier(**best_rf_params, random_state=RANDOM_SEED),
        KNeighborsClassifier(**best_knn_params)
    ), scaler=True, smote=True)

    # -------------------------------------------------------------------------
    # TAHAP 7: Evaluasi Uji Sekali Jalan pada Test Set Resmi (624 Citra)
    # -------------------------------------------------------------------------
    print("\n>>> TAHAP 7: Evaluasi Test Set Resmi (624 citra)...")
    test_eval_results = []
    test_predictions = {}
    test_probabilities = {}
    cms = {}

    for name, pipe in models_to_evaluate.items():
        t0_fit = time.time()
        pipe.fit(X_train_full, y_train_full)
        t_fit = time.time() - t0_fit

        t0_inf = time.time()
        y_pred = pipe.predict(X_test)
        t_inf_per_img = (time.time() - t0_inf) / len(X_test) * 1000.0 # ms

        try:
            y_prob = pipe.predict_proba(X_test)[:, 1]
        except Exception:
            y_prob = None

        m_dict = compute_all_metrics(y_test, y_pred, y_prob, model_name=name)
        m_dict['Waktu Latih (s)'] = round(t_fit, 3)
        m_dict['Inferensi (ms/citra)'] = round(t_inf_per_img, 3)
        test_eval_results.append(m_dict)

        if 'Tuned' in name or 'Voting' in name:
            clean_name = name.replace(' (Tuned)', '')
            test_predictions[clean_name] = y_pred
            if y_prob is not None:
                test_probabilities[clean_name] = y_prob
            from sklearn.metrics import confusion_matrix
            cms[clean_name] = confusion_matrix(y_test, y_pred)

    # Baseline Mayoritas
    maj_pred = np.ones(len(y_test), dtype=int) # All Pneumonia (390/624 = 62.5%)
    maj_metrics = compute_all_metrics(y_test, maj_pred, model_name='Baseline Mayoritas (Pneumonia Only)')
    maj_metrics['Waktu Latih (s)'] = 0.0
    maj_metrics['Inferensi (ms/citra)'] = 0.0
    test_eval_results.append(maj_metrics)

    df_test_metrics = pd.DataFrame(test_eval_results)
    table_eval_path = os.path.join(TABLES_DIR, "tabel_evaluasi_model_lengkap.csv")
    df_test_metrics.to_csv(table_eval_path, index=False)
    print(f"Hasil evaluasi test set disimpan ke {table_eval_path}")

    # Plot Confusion Matrix, ROC & DET
    plot_confusion_matrices(cms, output_path=os.path.join(FIGURES_DIR, "gambar2_confusion_matrix_revisi.png"))
    plot_roc_and_det_curves(test_probabilities, y_test,
                            output_roc=os.path.join(FIGURES_DIR, "kurva_roc.png"),
                            output_det=os.path.join(FIGURES_DIR, "kurva_det.png"))

    # -------------------------------------------------------------------------
    # ANALISIS OPTIMISM GAP (P1 vs P2 vs P3)
    # -------------------------------------------------------------------------
    print("\n>>> Analisis Optimism Gap Antar Protokol (P1 vs P2 vs P3)...")
    protocol_results = []
    for proto_code in ['P1', 'P2', 'P3']:
        tr_i = splits[proto_code]['train_idx']
        val_i = splits[proto_code]['val_idx']
        te_i = splits[proto_code]['test_idx']

        X_tr_proto = np.vstack([X_all[tr_i], X_all[val_i]])
        y_tr_proto = np.concatenate([y_all[tr_i], y_all[val_i]])
        X_te_proto = X_all[te_i]
        y_te_proto = y_all[te_i]

        # Model Voting
        v_pipe = build_pipeline(create_voting_ensemble(
            SVC(**best_svm_params, random_state=RANDOM_SEED),
            RandomForestClassifier(**best_rf_params, random_state=RANDOM_SEED),
            KNeighborsClassifier(**best_knn_params)
        ), scaler=True, smote=False)
        v_pipe.fit(X_tr_proto, y_tr_proto)
        p_pred = v_pipe.predict(X_te_proto)
        p_prob = v_pipe.predict_proba(X_te_proto)[:, 1]

        p_metrics = compute_all_metrics(y_te_proto, p_pred, p_prob, model_name=f"Ensemble Voting ({proto_code})")
        p_metrics['Protokol'] = proto_code
        p_metrics['Deskripsi Protokol'] = splits[proto_code]['protocol']
        protocol_results.append(p_metrics)

    df_proto = pd.DataFrame(protocol_results)
    proto_csv_path = os.path.join(TABLES_DIR, "analisis_optimism_gap_protokol.csv")
    df_proto.to_csv(proto_csv_path, index=False)
    print(f"Perbandingan protokol (Optimism Gap) disimpan ke {proto_csv_path}")

    # -------------------------------------------------------------------------
    # TAHAP 8: Uji Statistik McNemar & Keragaman Ensemble
    # -------------------------------------------------------------------------
    print("\n>>> TAHAP 8: Uji Statistik McNemar, Bootstrap CI, & Keragaman Ensemble...")
    run_all_mcnemar_tests(y_test, test_predictions)
    analyze_ensemble_diversity(y_test, test_predictions)

    # Bootstrap CI
    ci_rows = []
    for m_name, preds in test_predictions.items():
        ci_dict = compute_bootstrap_ci(y_test, preds)
        ci_dict['Model'] = m_name
        ci_rows.append(ci_dict)
    df_ci = pd.DataFrame(ci_rows)
    ci_csv = os.path.join(TABLES_DIR, "bootstrap_confidence_intervals.csv")
    df_ci.to_csv(ci_csv, index=False)
    print(f"Bootstrap CI disimpan ke {ci_csv}")

    # -------------------------------------------------------------------------
    # TAHAP 9: Ablasi & Uji Ketahanan Artefak (Crop Tepi)
    # -------------------------------------------------------------------------
    print("\n>>> TAHAP 9: Ablasi & Uji Ketahanan Artefak...")
    df_test_official = df_inventory.iloc[p2_te_idx].reset_index(drop=True)
    best_pipeline = models_to_evaluate['Ensemble Soft Voting']

    # Uji artefak (crop 10% dan 15%)
    test_artifact_robustness(best_pipeline, df_test_official, X_test, crop_percentages=[0.10, 0.15])

    # Permutation feature importance
    compute_permutation_importance(best_pipeline, X_test, y_test, GLCM_PROPERTIES)

    # -------------------------------------------------------------------------
    # TAHAP 10: Ekspor Model Pipeline Utuh untuk Streamlit
    # -------------------------------------------------------------------------
    print("\n>>> TAHAP 10: Ekspor Model Pipeline Final...")
    final_model_path = os.path.join(MODELS_DIR, "pneumonia_ensemble_pipeline.pkl")
    joblib.dump(best_pipeline, final_model_path)
    print(f"Model pipeline utuh disimpan ke {final_model_path}")

    # Catat metadata reproduksibilitas
    repro_meta = {
        'platform': platform.platform(),
        'python_version': platform.python_version(),
        'scikit_learn_version': sklearn.__version__,
        'imbalanced_learn_version': imblearn.__version__,
        'opencv_version': cv2.__version__,
        'scikit_image_version': skimage.__version__,
        'random_seed': RANDOM_SEED,
        'best_hyperparameters': {
            'knn': best_knn_params,
            'svm': best_svm_params,
            'rf': best_rf_params
        },
        'execution_time_seconds': round(time.time() - t_start, 2)
    }
    meta_path = os.path.join(RESULTS_DIR, "reproducibility_metadata.json")
    with open(meta_path, 'w') as fp:
        json.dump(repro_meta, fp, indent=2)
    print(f"Metadata reproduksibilitas disimpan ke {meta_path}")

    print("\n" + "=" * 80)
    print(f"SELURUH EKSPERIMEN REVISI BERHASIL DILAKSANAKAN DALAM {time.time() - t_start:.2f} DETIK!")
    print("=" * 80)

if __name__ == "__main__":
    main()
