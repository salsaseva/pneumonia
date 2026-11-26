"""
Tahap 9: Ablasi & Uji Ketahanan Artefak (Artifact Robustness)
1. Ablasi Preprocessing (basic vs clahe vs ortiz), Ng (32, 64, 256), Fitur (GLCM-5, 6, 24).
2. Uji Artefak: Crop tengah / hapus tepi 10% & 15% untuk menguji ketergantungan pada watermark/penanda tepi.
3. Permutation Feature Importance vs Gini Importance.
"""
import os
import cv2
import numpy as np
import pandas as pd
from sklearn.inspection import permutation_importance
from sklearn.metrics import accuracy_score, f1_score
from joblib import Parallel, delayed
from src.config import TABLES_DIR, FIGURES_DIR
from src.features import extract_glcm_single
from src.preprocessing import preprocess_image

def test_artifact_robustness(model_pipeline, df_test, X_test, crop_percentages=[0.10, 0.15], Ng=256, variant='GLCM-6', mode='basic', output_csv=None):
    """
    Uji ketahanan model terhadap artefak tepi/watermark:
    Memotong tepi gambar sebesar 10% dan 15% (hanya menyisakan area paru tengah).
    Mengekstrak fitur dan menguji apakah akurasi / F1 runtuh.
    """
    print("[Ablation] Menjalankan Uji Ketahanan Artefak (Center Crop / Edge Removal)...")
    results = []

    y_test = df_test['target'].values

    # Baseline (tanpa crop)
    y_pred_base = model_pipeline.predict(X_test)
    acc_base = accuracy_score(y_test, y_pred_base)
    f1_base = f1_score(y_test, y_pred_base, zero_division=0)
    results.append({
        'Kondisi': 'Baseline (Citra Utuh 0% Crop)',
        'Akurasi (%)': round(acc_base * 100, 2),
        'F1-Score (%)': round(f1_base * 100, 2),
        'Keterangan': 'Standar evaluasi resmi'
    })

    for crop_p in crop_percentages:
        print(f"  -> Menguji crop tepi {int(crop_p*100)}%...")
        
        def _crop_and_extract(row):
            img = cv2.imread(row['file_path'], cv2.IMREAD_GRAYSCALE)
            h, w = img.shape
            dh, dw = int(h * crop_p), int(w * crop_p)
            cropped = img[dh:h-dh, dw:w-dw]
            cropped_prep = preprocess_image(cropped, mode=mode)
            return extract_glcm_single(cropped_prep, Ng=Ng, variant=variant)

        X_cropped = Parallel(n_jobs=-1)(delayed(_crop_and_extract)(row) for _, row in df_test.iterrows())
        X_cropped = np.array(X_cropped, dtype=np.float32)

        y_pred = model_pipeline.predict(X_cropped)
        acc = accuracy_score(y_test, y_pred)
        f1 = f1_score(y_test, y_pred, zero_division=0)

        results.append({
            'Kondisi': f'Crop Tepi {int(crop_p*100)}%',
            'Akurasi (%)': round(acc * 100, 2),
            'F1-Score (%)': round(f1 * 100, 2),
            'Keterangan': 'Evaluasi robustitas terhadap penanda/tepi'
        })

    df_res = pd.DataFrame(results)
    if output_csv is None:
        os.makedirs(TABLES_DIR, exist_ok=True)
        output_csv = os.path.join(TABLES_DIR, "artifact_robustness_test.csv")
    df_res.to_csv(output_csv, index=False)
    print(f"[Ablation] Hasil uji artefak disimpan ke {output_csv}")
    return df_res

def compute_permutation_importance(fitted_pipeline, X_test, y_test, feature_names, n_repeats=10, seed=42, output_csv=None):
    """
    Menghitung Permutation Feature Importance untuk menghindari bias Gini
    pada fitur yang saling berkorelasi (misal ASM dan Energy).
    """
    print("[Ablation] Menghitung Permutation Feature Importance...")
    perm = permutation_importance(
        fitted_pipeline, X_test, y_test,
        n_repeats=n_repeats, random_state=seed, scoring='f1', n_jobs=-1
    )

    records = []
    for i, name in enumerate(feature_names):
        records.append({
            'Fitur': name,
            'Permutation Importance Mean': round(float(perm.importances_mean[i]), 5),
            'Permutation Importance Std': round(float(perm.importances_std[i]), 5)
        })

    df_perm = pd.DataFrame(records).sort_values(by='Permutation Importance Mean', ascending=False)
    if output_csv is None:
        os.makedirs(TABLES_DIR, exist_ok=True)
        output_csv = os.path.join(TABLES_DIR, "permutation_feature_importance.csv")
    df_perm.to_csv(output_csv, index=False)
    print(f"[Ablation] Permutation importance disimpan ke {output_csv}")
    return df_perm
