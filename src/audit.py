"""
Tahap 1: Audit Data & Ekstraksi Patient ID
Inventarisasi duplikat hash, citra korup, label, jumlah per kelas,
ekstraksi ID pasien, dan verifikasi overlap pasien antar split resmi.
"""
import os
import re
import json
import hashlib
from PIL import Image
from collections import defaultdict
import pandas as pd
from src.config import DATASET_DIR, AUDIT_DIR

def run_audit(dataset_dir=DATASET_DIR, output_dir=AUDIT_DIR):
    os.makedirs(output_dir, exist_ok=True)
    records = []
    corrupt_images = []
    hashes = defaultdict(list)

    for split in ['train', 'val', 'test']:
        for label in ['NORMAL', 'PNEUMONIA']:
            folder = os.path.join(dataset_dir, split, label)
            if not os.path.exists(folder):
                continue
            for fname in sorted(os.listdir(folder)):
                if fname.startswith('.'):
                    continue
                fpath = os.path.join(folder, fname)
                
                # Verify image integrity
                try:
                    with Image.open(fpath) as img:
                        img.verify()
                    with Image.open(fpath) as img:
                        w, h = img.size
                        mode = img.mode
                except Exception as e:
                    corrupt_images.append({'path': fpath, 'error': str(e)})
                    continue

                # SHA-256 file hash
                with open(fpath, 'rb') as fp:
                    fhash = hashlib.sha256(fp.read()).hexdigest()
                hashes[fhash].append((split, label, fname))

                # Patient ID extraction
                if label == 'PNEUMONIA':
                    m = re.match(r'^(person\d+)_', fname)
                    pid = m.group(1) if m else f"UNK_PNEU_{fname}"
                    sub_type = 'bacteria' if 'bacteria' in fname else ('virus' if 'virus' in fname else 'unknown')
                else:
                    m = re.match(r'^(NORMAL2-IM-\d+|IM-\d+)', fname)
                    pid = m.group(1) if m else fname.split('.')[0]
                    sub_type = 'normal'

                records.append({
                    'official_split': 'train' if split in ['train', 'val'] else 'test',
                    'raw_split': split,
                    'label': label,
                    'target': 1 if label == 'PNEUMONIA' else 0,
                    'patient_id': pid,
                    'sub_type': sub_type,
                    'filename': fname,
                    'file_path': fpath,
                    'rel_path': os.path.join(split, label, fname),
                    'hash': fhash,
                    'width': w,
                    'height': h,
                    'mode': mode
                })

    df = pd.DataFrame(records)
    csv_path = os.path.join(output_dir, "dataset_inventory.csv")
    df.to_csv(csv_path, index=False)
    print(f"[Audit] Inventarisasi disimpan ke {csv_path} (Total: {len(df)} citra)")

    # Duplicates check
    dup_hashes = {h: lst for h, lst in hashes.items() if len(lst) > 1}
    
    # Overlap between Train+Val and Test
    train_val_patients = set(df[df['official_split'] == 'train']['patient_id'])
    test_patients = set(df[df['official_split'] == 'test']['patient_id'])
    overlap = sorted(list(train_val_patients.intersection(test_patients)))

    # Per-class overlap
    pneu_train_pids = set(df[(df['official_split'] == 'train') & (df['target'] == 1)]['patient_id'])
    pneu_test_pids = set(df[(df['official_split'] == 'test') & (df['target'] == 1)]['patient_id'])
    pneu_overlap = sorted(list(pneu_train_pids.intersection(pneu_test_pids)))

    norm_train_pids = set(df[(df['official_split'] == 'train') & (df['target'] == 0)]['patient_id'])
    norm_test_pids = set(df[(df['official_split'] == 'test') & (df['target'] == 0)]['patient_id'])
    norm_overlap = sorted(list(norm_train_pids.intersection(norm_test_pids)))

    overlap_info = {
        'total_images': len(df),
        'corrupt_images': len(corrupt_images),
        'duplicate_hashes_count': len(dup_hashes),
        'official_train_val_images': len(df[df['official_split'] == 'train']),
        'official_test_images': len(df[df['official_split'] == 'test']),
        'unique_patients_total': df['patient_id'].nunique(),
        'unique_patients_train_val': len(train_val_patients),
        'unique_patients_test': len(test_patients),
        'overlapping_patients_count': len(overlap),
        'pneumonia_test_patients': len(pneu_test_pids),
        'pneumonia_overlapping_patients_count': len(pneu_overlap),
        'pneumonia_overlap_percentage_in_test': round(len(pneu_overlap) / len(pneu_test_pids) * 100, 2),
        'normal_test_patients': len(norm_test_pids),
        'normal_overlapping_patients_count': len(norm_overlap),
        'overlapping_patient_ids': overlap
    }

    summary_path = os.path.join(output_dir, "audit_summary.json")
    with open(summary_path, 'w') as fp:
        json.dump(overlap_info, fp, indent=2)
    print(f"[Audit] Ringkasan audit disimpan ke {summary_path}")

    return df, overlap_info

if __name__ == "__main__":
    run_audit()
