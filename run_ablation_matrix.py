"""
Ablation Study Matrix (Optimized):
Mengevaluasi seluruh kombinasi 3x3x3:
- Preprocessing: basic, clahe, ortiz
- Ng (gray levels): 32, 64, 256
- Fitur: GLCM-5, GLCM-6, GLCM-24
Pada protokol resmi P2 untuk menjawab Research Question 2 (RQ2).
"""
import os
import json
import time
import numpy as np
import pandas as pd
from joblib import Parallel, delayed
from sklearn.svm import SVC
from sklearn.neighbors import KNeighborsClassifier
from sklearn.ensemble import RandomForestClassifier

from src.config import DATASET_DIR, TABLES_DIR, RESULTS_DIR, RANDOM_SEED, CACHE_DIR
from src.audit import run_audit
from src.splits import generate_splits
from src.preprocessing import compute_reference_histogram, preprocess_image
from src.features import extract_glcm_all_variants
from src.pipelines import build_pipeline, create_voting_ensemble
from src.evaluation import compute_all_metrics

def main():
    print("[Ablation Matrix] Memulai Studi Ablasi Preprocessing, Ng, dan Varian Fitur (3x3x3 = 27 Kombinasi)...")
    df_inventory, _ = run_audit(DATASET_DIR)
    
    splits_file = os.path.join(RESULTS_DIR, "splits", f"splits_seed_{RANDOM_SEED}.json")
    with open(splits_file, 'r') as fp:
        splits = json.load(fp)

    p2 = splits['P2']
    tr_val_idx = p2['train_idx'] + p2['val_idx']
    te_idx = p2['test_idx']

    p2_train_paths = df_inventory.iloc[p2['train_idx']]['file_path'].tolist()
    ref_cdf = compute_reference_histogram(p2_train_paths)

    preproc_modes = ['basic', 'clahe', 'ortiz']
    ng_values = [32, 64, 256]
    feature_variants = ['GLCM-5', 'GLCM-6', 'GLCM-24']

    records = []

    for mode in preproc_modes:
        for ng in ng_values:
            print(f"\n>>> Mengekstrak GLCM matriks untuk mode='{mode}', Ng={ng}...")
            t0 = time.time()
            
            # Check if all 3 variants exist in cache
            cached = True
            var_data = {}
            for var in feature_variants:
                cfile = os.path.join(CACHE_DIR, f"features_{mode}_ng{ng}_{var}.npz")
                if os.path.exists(cfile):
                    var_data[var] = np.load(cfile)['X']
                else:
                    cached = False
                    break

            if not cached:
                def _worker(row):
                    img_prep = preprocess_image(row['file_path'], mode=mode, ref_cdf=ref_cdf)
                    return extract_glcm_all_variants(img_prep, Ng=ng)

                res_dicts = Parallel(n_jobs=-1, verbose=0)(
                    delayed(_worker)(row) for _, row in df_inventory.iterrows()
                )

                # Unpack and cache all 3 variants
                y_arr = df_inventory['target'].values.astype(np.int32)
                for var in feature_variants:
                    X_mat = np.array([d[var] for d in res_dicts], dtype=np.float32)
                    var_data[var] = X_mat
                    cfile = os.path.join(CACHE_DIR, f"features_{mode}_ng{ng}_{var}.npz")
                    np.savez_compressed(cfile, X=X_mat, y=y_arr)
                print(f"    Ekstraksi 3 varian selesai dalam {time.time() - t0:.2f}s")
            else:
                print(f"    Memuat cache untuk mode='{mode}', Ng={ng}")

            y_all = df_inventory['target'].values.astype(np.int32)

            for variant in feature_variants:
                X_all = var_data[variant]
                X_train_full = X_all[tr_val_idx]
                y_train_full = y_all[tr_val_idx]
                X_test = X_all[te_idx]
                y_test = y_all[te_idx]

                # Model Ensemble Voting
                svm_m = SVC(kernel='rbf', C=10.0, gamma='scale', probability=True, random_state=RANDOM_SEED)
                rf_m = RandomForestClassifier(n_estimators=100, min_samples_leaf=2, random_state=RANDOM_SEED)
                knn_m = KNeighborsClassifier(n_neighbors=11)
                
                voting = create_voting_ensemble(svm_m, rf_m, knn_m)
                pipe = build_pipeline(voting, scaler=True, smote=False)

                t0_fit = time.time()
                pipe.fit(X_train_full, y_train_full)
                t_fit = time.time() - t0_fit

                y_pred = pipe.predict(X_test)
                y_prob = pipe.predict_proba(X_test)[:, 1]

                m = compute_all_metrics(y_test, y_pred, y_prob, model_name=f"{mode}_{ng}_{variant}")
                m['Preprocessing'] = mode
                m['Ng'] = ng
                m['Varian Fitur'] = variant
                m['Jumlah Fitur'] = X_all.shape[1]
                m['Waktu Latih (s)'] = round(t_fit, 3)

                print(f"    [{mode:5s} | Ng={ng:3d} | {variant:7s}] -> Acc: {m['Accuracy (%)']}%, Sen: {m['Recall/Sen (%)']}%, Spe: {m['Specificity (%)']}%, F1: {m['F1-Score (%)']}%, BalAcc: {m['Balanced Accuracy (%)']}%")
                records.append(m)

    df_abl = pd.DataFrame(records)
    out_csv = os.path.join(TABLES_DIR, "tabel_ablasi_preprocessing_ng_fitur.csv")
    df_abl.to_csv(out_csv, index=False)
    print(f"\n[Ablation Matrix] Seluruh 27 kombinasi berhasil dievaluasi dan disimpan ke {out_csv}!")

if __name__ == "__main__":
    main()
