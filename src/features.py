"""
Tahap 4: Ekstraksi Fitur GLCM & First-Order Radiomics
Mendukung:
- Varian: 'GLCM-6' (rerata 4 arah), 'GLCM-24' (per arah), 'GLCM-5' (tanpa Energy), 'GLCM-6-firstorder'
- Ng (gray levels): 32, 64, 256
- Disk Caching untuk mempercepat eksekusi eksperimen
"""
import os
import cv2
import numpy as np
from skimage.feature import graycomatrix, graycoprops
from scipy.stats import skew, kurtosis, entropy
from joblib import Parallel, delayed
from src.config import ANGLES, DISTANCES, GLCM_PROPERTIES, CACHE_DIR
from src.preprocessing import preprocess_image

def quantize_image(image, Ng):
    """Kuantisasi citra ke jumlah level abu-abu Ng (32, 64, atau 256)."""
    if Ng == 256:
        return image
    factor = 256 // Ng
    quantized = (image // factor).astype(np.uint8)
    return np.clip(quantized, 0, Ng - 1)

def extract_glcm_single(image_processed, Ng=256, variant='GLCM-6', include_first_order=False):
    """
    Ekstraksi fitur GLCM dari satu citra yang sudah diproses.
    variant: 'GLCM-6', 'GLCM-24', 'GLCM-5'
    """
    img_q = quantize_image(image_processed, Ng)
    glcm = graycomatrix(
        img_q, distances=DISTANCES, angles=ANGLES,
        levels=Ng, symmetric=True, normed=True
    )

    feature_dict = {}
    for prop in GLCM_PROPERTIES:
        vals = graycoprops(glcm, prop)[0] # array 4 arah: [0, 45, 90, 135]
        feature_dict[prop] = vals

    # Varian fitur
    if variant == 'GLCM-24':
        # 6 properti x 4 arah = 24 fitur
        features = []
        for prop in GLCM_PROPERTIES:
            features.extend(feature_dict[prop])
        features = np.array(features, dtype=np.float32)

    elif variant == 'GLCM-5':
        # Tanpa Energy (karena Energy = sqrt(ASM))
        props_5 = ['contrast', 'dissimilarity', 'homogeneity', 'correlation', 'ASM']
        features = np.array([np.mean(feature_dict[p]) for p in props_5], dtype=np.float32)

    else: # GLCM-6 default
        props_6 = ['contrast', 'dissimilarity', 'homogeneity', 'energy', 'correlation', 'ASM']
        features = np.array([np.mean(feature_dict[p]) for p in props_6], dtype=np.float32)

    if include_first_order:
        flat = image_processed.flatten().astype(np.float64)
        hist, _ = np.histogram(flat, bins=Ng, range=(0, Ng if Ng != 256 else 256), density=True)
        hist = hist[hist > 0]
        fo_mean = np.mean(flat)
        fo_std = np.std(flat)
        fo_skew = float(skew(flat))
        fo_kurt = float(kurtosis(flat))
        fo_ent = float(entropy(hist))
        first_order = np.array([fo_mean, fo_std, fo_skew, fo_kurt, fo_ent], dtype=np.float32)
        features = np.concatenate([features, first_order])

    return features

def extract_glcm_all_variants(image_processed, Ng=256):
    """
    Menghitung GLCM matrix satu kali, lalu langsung mengekstrak
    ketiga varian (GLCM-5, GLCM-6, GLCM-24) secara simultan.
    """
    img_q = quantize_image(image_processed, Ng)
    glcm = graycomatrix(
        img_q, distances=DISTANCES, angles=ANGLES,
        levels=Ng, symmetric=True, normed=True
    )

    feature_dict = {}
    for prop in GLCM_PROPERTIES:
        feature_dict[prop] = graycoprops(glcm, prop)[0]

    # GLCM-24 (6 properti x 4 arah)
    f24 = []
    for prop in GLCM_PROPERTIES:
        f24.extend(feature_dict[prop])
    f24 = np.array(f24, dtype=np.float32)

    # GLCM-6 (rerata 4 arah)
    f6 = np.array([np.mean(feature_dict[p]) for p in GLCM_PROPERTIES], dtype=np.float32)

    # GLCM-5 (tanpa Energy)
    props_5 = ['contrast', 'dissimilarity', 'homogeneity', 'correlation', 'ASM']
    f5 = np.array([np.mean(feature_dict[p]) for p in props_5], dtype=np.float32)

    return {'GLCM-5': f5, 'GLCM-6': f6, 'GLCM-24': f24}

def extract_features_dataset(df, mode='basic', Ng=256, variant='GLCM-6', ref_cdf=None, use_cache=True, n_jobs=-1):
    """
    Ekstraksi fitur GLCM paralel untuk seluruh dataset dengan disk caching.
    """
    os.makedirs(CACHE_DIR, exist_ok=True)
    cache_file = os.path.join(CACHE_DIR, f"features_{mode}_ng{Ng}_{variant}.npz")

    if use_cache and os.path.exists(cache_file):
        data = np.load(cache_file)
        print(f"[Features] Memuat cache fitur dari {cache_file} (Shape: {data['X'].shape})")
        return data['X'], data['y']

    print(f"[Features] Mengekstrak fitur untuk {len(df)} citra (mode={mode}, Ng={Ng}, variant={variant})...")

    def _worker(row):
        fpath = row['file_path']
        img_prep = preprocess_image(fpath, mode=mode, ref_cdf=ref_cdf)
        feat = extract_glcm_single(img_prep, Ng=Ng, variant=variant)
        return feat

    results = Parallel(n_jobs=n_jobs, verbose=0)(
        delayed(_worker)(row) for _, row in df.iterrows()
    )

    X = np.array(results, dtype=np.float32)
    y = df['target'].values.astype(np.int32)

    # Simpan ke cache
    np.savez_compressed(cache_file, X=X, y=y)
    print(f"[Features] Ekstraksi selesai & cache disimpan ke {cache_file} (Shape: {X.shape})")
    return X, y
