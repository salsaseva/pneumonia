"""
Tahap 2: Tiga Protokol Pembagian Data (P1, P2, P3)
P1: Acak 80:20 per citra (kode lama)
P2: Split resmi 5.232 / 624, validasi tetap per-pasien (StratifiedGroupKFold)
P3: Gabungan semua citra, split murni per-pasien (latih / validasi / uji)
"""
import os
import json
import numpy as np
import pandas as pd
from sklearn.model_selection import StratifiedGroupKFold, train_test_split, GroupShuffleSplit
from src.config import RESULTS_DIR, RANDOM_SEED

SPLITS_DIR = os.path.join(RESULTS_DIR, "splits")

def generate_splits(df, seed=RANDOM_SEED, output_dir=SPLITS_DIR):
    os.makedirs(output_dir, exist_ok=True)
    splits = {}

    # =========================================================================
    # PROTOKOL P1: Split Acak 80:20 per citra (Kode Lama / Image-level random)
    # =========================================================================
    # Latih 80%, Uji 20%, dari seluruh dataset atau 80:10:10 train/val/test
    p1_tr_val_idx, p1_te_idx = train_test_split(
        np.arange(len(df)), test_size=0.20, random_state=seed, stratify=df['target']
    )
    p1_tr_idx, p1_val_idx = train_test_split(
        p1_tr_val_idx, test_size=0.125, random_state=seed, stratify=df.iloc[p1_tr_val_idx]['target'] # 0.8 * 0.125 = 0.10 (70:10:20)
    )
    splits['P1'] = {
        'protocol': 'P1 (Random 70:10:20 Image-level)',
        'train_idx': p1_tr_idx.tolist(),
        'val_idx': p1_val_idx.tolist(),
        'test_idx': p1_te_idx.tolist()
    }

    # =========================================================================
    # PROTOKOL P2: Split Resmi 5.232 / 624, Validasi Tetap per Pasien
    # =========================================================================
    train_val_mask = (df['official_split'] == 'train').values
    test_mask = (df['official_split'] == 'test').values
    
    p2_train_val_indices = np.where(train_val_mask)[0]
    p2_test_indices = np.where(test_mask)[0]

    df_p2_tv = df.iloc[p2_train_val_indices].reset_index(drop=True)
    sgkf = StratifiedGroupKFold(n_splits=5, shuffle=True, random_state=seed)
    
    # Ambil fold 0 sebagai validasi tetap (~20% dari 5232 ≈ 1046 citra)
    p2_folds = []
    for fold, (f_tr_sub, f_val_sub) in enumerate(sgkf.split(df_p2_tv, df_p2_tv['target'], df_p2_tv['patient_id'])):
        f_tr_orig = p2_train_val_indices[f_tr_sub].tolist()
        f_val_orig = p2_train_val_indices[f_val_sub].tolist()
        
        # Assert tidak ada ID pasien yang sama di latih dan validasi
        tr_pids = set(df.iloc[f_tr_orig]['patient_id'])
        val_pids = set(df.iloc[f_val_orig]['patient_id'])
        overlap = tr_pids.intersection(val_pids)
        assert len(overlap) == 0, f"P2 Fold {fold} patient overlap: {overlap}"
        
        p2_folds.append({'train_idx': f_tr_orig, 'val_idx': f_val_orig})

    # P2 utama memakai fold 0 sebagai validasi tetap
    splits['P2'] = {
        'protocol': 'P2 (Official 5232/624 with Patient-level Validation)',
        'train_idx': p2_folds[0]['train_idx'],
        'val_idx': p2_folds[0]['val_idx'],
        'test_idx': p2_test_indices.tolist(),
        'cv_5folds': p2_folds
    }

    # =========================================================================
    # PROTOKOL P3: Gabungan Semua Citra, Split Murni per Pasien
    # =========================================================================
    # 70% Train, 15% Val, 15% Test berbasis Patient ID (Strict Zero-Leakage)
    gss_test = GroupShuffleSplit(n_splits=1, test_size=0.15, random_state=seed)
    p3_tr_val_idx, p3_test_idx = next(gss_test.split(df, df['target'], df['patient_id']))
    
    df_p3_tv = df.iloc[p3_tr_val_idx].reset_index(drop=True)
    # Val 15% dari total = 0.15 / 0.85 ≈ 0.1765 dari train_val
    gss_val = GroupShuffleSplit(n_splits=1, test_size=0.1765, random_state=seed)
    p3_sub_tr, p3_sub_val = next(gss_val.split(df_p3_tv, df_p3_tv['target'], df_p3_tv['patient_id']))
    
    p3_train_indices = p3_tr_val_idx[p3_sub_tr]
    p3_val_indices = p3_tr_val_idx[p3_sub_val]

    # Assert mutual exclusiveness of patients
    p3_tr_pids = set(df.iloc[p3_train_indices]['patient_id'])
    p3_val_pids = set(df.iloc[p3_val_indices]['patient_id'])
    p3_te_pids = set(df.iloc[p3_test_idx]['patient_id'])

    assert len(p3_tr_pids.intersection(p3_val_pids)) == 0, "P3 train-val patient overlap!"
    assert len(p3_tr_pids.intersection(p3_te_pids)) == 0, "P3 train-test patient overlap!"
    assert len(p3_val_pids.intersection(p3_te_pids)) == 0, "P3 val-test patient overlap!"

    splits['P3'] = {
        'protocol': 'P3 (Combined All Images, Strict Patient-Level Split)',
        'train_idx': p3_train_indices.tolist(),
        'val_idx': p3_val_indices.tolist(),
        'test_idx': p3_test_idx.tolist()
    }

    splits_file = os.path.join(output_dir, f"splits_seed_{seed}.json")
    with open(splits_file, 'w') as fp:
        json.dump(splits, fp, indent=2)
    print(f"[Splits] Berhasil membuat dan memvalidasi protokol P1, P2, P3. Disimpan ke {splits_file}")

    return splits

if __name__ == "__main__":
    from src.audit import run_audit
    df, _ = run_audit()
    splits = generate_splits(df)
    for p in ['P1', 'P2', 'P3']:
        tr_len = len(splits[p]['train_idx'])
        val_len = len(splits[p]['val_idx'])
        te_len = len(splits[p]['test_idx'])
        print(f"{p}: Train={tr_len}, Val={val_len}, Test={te_len}")
