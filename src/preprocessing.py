"""
Tahap 3: Prapengolahan Citra
Mode yang didukung:
1. 'basic': Grayscale & resize (150x150)
2. 'clahe': basic + CLAHE (clipLimit=2.0, tileGridSize=(8,8))
3. 'ortiz': pencocokan histogram, BCET, masking persentil 25% teratas, CLAHE.
"""
import os
import cv2
import numpy as np
from src.config import IMG_SIZE, MODELS_DIR

def compute_reference_histogram(image_paths, output_path=None):
    """Menghitung rata-rata histogram kumulatif HANYA dari data latih (atau memuat jika sudah ada)."""
    if output_path is None:
        os.makedirs(MODELS_DIR, exist_ok=True)
        output_path = os.path.join(MODELS_DIR, "ref_histogram.npy")
    if os.path.exists(output_path):
        return np.load(output_path)

    accum_hist = np.zeros(256, dtype=np.float64)
    valid_count = 0
    for path in image_paths:
        img = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
        if img is not None:
            hist, _ = np.histogram(img.flatten(), 256, [0, 256])
            accum_hist += hist / (hist.sum() + 1e-7)
            valid_count += 1
    if valid_count > 0:
        accum_hist /= valid_count
    ref_cdf = np.cumsum(accum_hist)
    ref_cdf = (ref_cdf - ref_cdf.min()) / (ref_cdf.max() - ref_cdf.min() + 1e-7)
    
    if output_path is None:
        os.makedirs(MODELS_DIR, exist_ok=True)
        output_path = os.path.join(MODELS_DIR, "ref_histogram.npy")
    np.save(output_path, ref_cdf)
    print(f"[Preprocessing] Histogram referensi dihitung dari {valid_count} citra latih dan disimpan ke {output_path}")
    return ref_cdf

def match_histogram(image, ref_cdf):
    """Pencocokan histogram citra terhadap CDF referensi."""
    hist, _ = np.histogram(image.flatten(), 256, [0, 256])
    img_cdf = np.cumsum(hist).astype(np.float64)
    img_cdf = (img_cdf - img_cdf.min()) / (img_cdf.max() - img_cdf.min() + 1e-7)
    
    lut = np.interp(img_cdf, ref_cdf, np.arange(256)).astype(np.uint8)
    return cv2.LUT(image, lut)

def apply_bcet(image, L=0, H=255, E=110):
    """
    Balanced Contrast Enhancement Technique (BCET).
    y = A*(x - l)^2 + B*(x - l) + C
    Menyesuaikan min (L), max (H), dan mean (E) gambar.
    """
    img_f = image.astype(np.float64)
    l = np.min(img_f)
    h = np.max(img_f)
    e = np.mean(img_f)
    v = np.mean(img_f ** 2)

    if h == l:
        return image

    # BCET coefficients
    # b = (h - l)*(E - L) - (H - L)*(e - l) / ...
    denom = 2 * (h - l) * (e - l) - (v - l**2)
    if abs(denom) < 1e-5:
        # Fallback to linear min-max stretch
        stretched = (img_f - l) / (h - l) * 255.0
        return np.clip(stretched, 0, 255).astype(np.uint8)

    B = ((h - l)**2 * (E - L) - (H - L) * (v - 2*l*e + l**2)) / ((h - l)**2 * (e - l) - (h - l)*(v - 2*l*e + l**2) + 1e-7)
    A = (H - L - B * (h - l)) / ((h - l)**2 + 1e-7)
    C = L

    out = A * (img_f - l)**2 + B * (img_f - l) + C
    return np.clip(out, 0, 255).astype(np.uint8)

def mask_top_percentile(image, percentile=75):
    """Masking persentil teratas (misal intensitas > persentil 75% di-cap atau di-smooth)."""
    thresh = np.percentile(image, percentile)
    out = image.copy()
    out[out > thresh] = int(thresh)
    return out

def preprocess_image(image_input, mode='basic', ref_cdf=None, target_size=IMG_SIZE):
    """
    Eksekusi prapengolahan citra berdasarkan mode:
    - 'basic': grayscale & resize
    - 'clahe': basic + CLAHE
    - 'ortiz': histogram matching + BCET + top 25% masking + CLAHE
    image_input dapat berupa file_path (str) atau numpy.ndarray.
    """
    if isinstance(image_input, str):
        img = cv2.imread(image_input, cv2.IMREAD_GRAYSCALE)
        if img is None:
            raise FileNotFoundError(f"Cannot read image at {image_input}")
    else:
        if len(image_input.shape) == 3:
            img = cv2.cvtColor(image_input, cv2.COLOR_BGR2GRAY)
        else:
            img = image_input.copy()

    # Resize to target size
    img = cv2.resize(img, target_size, interpolation=cv2.INTER_AREA)

    if mode == 'basic':
        return img
    
    clahe_obj = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))

    if mode == 'clahe':
        return clahe_obj.apply(img)

    if mode == 'ortiz':
        # 1. Pencocokan histogram jika ada referensi
        if ref_cdf is not None:
            img = match_histogram(img, ref_cdf)
        # 2. BCET
        img = apply_bcet(img)
        # 3. Masking persentil 25% teratas (cap at 75th percentile)
        img = mask_top_percentile(img, percentile=75)
        # 4. CLAHE
        img = clahe_obj.apply(img)
        return img

    raise ValueError(f"Unknown preprocessing mode: {mode}")
