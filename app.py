import os
import streamlit as st
import cv2
import numpy as np
import pandas as pd
import joblib
from skimage.feature import graycomatrix, graycoprops

# =============================================================
# 1. KONFIGURASI HALAMAN
# =============================================================
st.set_page_config(
    page_title="Deteksi Pneumonia - GLCM & Ensemble ML (Revisi Ilmiah)",
    page_icon="🫁",
    layout="wide"
)

# Custom CSS ringkas
st.markdown("""
    <style>
    .stApp {
        max-width: 1200px;
        margin: 0 auto;
    }
    .main-title {
        text-align: center;
        margin-bottom: 0.25rem;
    }
    .sub-title {
        text-align: center;
        color: #6c757d;
        margin-bottom: 1.5rem;
    }
    </style>
""", unsafe_allow_html=True)

# =============================================================
# 2. LOAD PIPELINE MODEL UTUH (JOBLIB)
# =============================================================
@st.cache_resource
def load_trained_pipeline():
    model_path = os.path.join(os.path.dirname(__file__), "results", "models", "pneumonia_ensemble_pipeline.pkl")
    if not os.path.exists(model_path):
        # Fallback to local root if needed
        model_path = "model_voting.pkl"
        pipeline = joblib.load(model_path)
        return pipeline, None
    pipeline = joblib.load(model_path)
    
    ref_hist_path = os.path.join(os.path.dirname(__file__), "results", "models", "ref_histogram.npy")
    ref_cdf = np.load(ref_hist_path) if os.path.exists(ref_hist_path) else None
    return pipeline, ref_cdf

pipeline, ref_cdf = load_trained_pipeline()

# =============================================================
# 3. PREPROCESSING & EKSTRAKSI FITUR GLCM
# =============================================================
def extract_glcm_from_bytes(uploaded_file, target_size=(150, 150)):
    file_bytes = np.frombuffer(uploaded_file.getvalue(), dtype=np.uint8)
    img_bgr = cv2.imdecode(file_bytes, cv2.IMREAD_COLOR)
    
    if img_bgr is None:
        return None, None, None, None
        
    # Preprocessing: Grayscale & Resize 150x150
    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    resized = cv2.resize(gray, target_size, interpolation=cv2.INTER_AREA)
    
    # Ekstraksi GLCM 4 arah (0, 45, 90, 135 deg)
    distances = [1]
    angles = [0, np.pi/4, np.pi/2, 3*np.pi/4]
    glcm = graycomatrix(resized, distances=distances, angles=angles, levels=256, symmetric=True, normed=True)
    
    features = {
        'contrast': float(graycoprops(glcm, 'contrast').mean()),
        'dissimilarity': float(graycoprops(glcm, 'dissimilarity').mean()),
        'homogeneity': float(graycoprops(glcm, 'homogeneity').mean()),
        'energy': float(graycoprops(glcm, 'energy').mean()),
        'correlation': float(graycoprops(glcm, 'correlation').mean()),
        'asm': float(graycoprops(glcm, 'ASM').mean())
    }
    
    feat_array = np.array([
        features['contrast'], features['dissimilarity'], features['homogeneity'],
        features['energy'], features['correlation'], features['asm']
    ], dtype=np.float32).reshape(1, -1)
    
    return img_bgr, resized, feat_array, features

# =============================================================
# 4. SIDEBAR & HEADER
# =============================================================
st.markdown("<h1 class='main-title'>🫁 Deteksi Pneumonia Citra Rontgen</h1>", unsafe_allow_html=True)
st.markdown("<p class='sub-title'>Purwarupa Evaluasi Fitur Tekstur <b>GLCM</b> & <b>Ensemble Machine Learning</b></p>", unsafe_allow_html=True)

st.sidebar.header("⚙️ Pengaturan Model")

model_choice = st.sidebar.selectbox(
    "Pilih Model Klasifikasi:",
    [
        "Ensemble Soft Voting (SVM + RF + KNN)",
        "Support Vector Machine (SVM Tuned)",
        "Random Forest (RF Tuned)",
        "K-Nearest Neighbors (KNN Tuned)"
    ]
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 📊 Benchmark Evaluasi Test Resmi")
st.sidebar.caption("Hasil evaluasi pada 624 citra uji resmi (Kermany et al.):")
eval_summary = pd.DataFrame({
    'Metrik': ['Akurasi', 'Sensitivitas (Recall)', 'Spesifisitas', 'F1-Score', 'Balanced Acc'],
    'Ensemble': ['79.33%', '92.82%', '56.84%', '84.88%', '74.83%'],
    'SVM': ['79.33%', '91.03%', '59.83%', '84.62%', '75.43%']
})
st.sidebar.dataframe(eval_summary, hide_index=True)

st.sidebar.markdown("---")
st.sidebar.warning(
    "⚠️ **Disclaimer Klinis:** Purwarupa ini dikembangkan semata-mata untuk demonstrasi penelitian empiris akademis "
    "dan **bukan merupakan alat diagnosis medis pengganti dokter spesialis radiologi**."
)

# =============================================================
# 5. LAYOUT UTAMA (2 KOLOM SIMETRIS)
# =============================================================
col1, col2 = st.columns([1, 1], gap="medium")

# --- KOLOM KIRI: INPUT & IMAGE PREVIEW ---
with col1:
    with st.container(border=True):
        st.subheader("📁 Unggah Citra Rontgen")
        uploaded_file = st.file_uploader("Pilih file citra X-Ray dada (.jpg, .jpeg, .png):", type=["jpg", "jpeg", "png"])
        
        if uploaded_file is not None:
            img_bgr, resized_gray, feat_array, feat_dict = extract_glcm_from_bytes(uploaded_file)
            
            if img_bgr is not None:
                tab_orig, tab_proc = st.tabs(["🖼️ Citra Asli", "⚙️ Preprocessed (150x150)"])
                with tab_orig:
                    st.image(cv2.cvtColor(img_bgr, cv2.COLOR_BGR2RGB), use_container_width=True)
                with tab_proc:
                    st.image(resized_gray, caption="Output Grayscale & Resized 150x150 untuk GLCM", use_container_width=True)
            else:
                st.error("Gagal membaca file citra. Pastikan format file gambar valid.")
        else:
            st.info("👈 Silakan unggah citra rontgen dada melalui kotak di atas.")

# --- KOLOM KANAN: HASIL ANALISIS & DIAGNOSIS ---
with col2:
    if uploaded_file is not None and img_bgr is not None:
        # Menentukan estimator berdasarkan pilihan sidebar
        scaler = pipeline.named_steps.get('scaler', None)
        voting_clf = pipeline.named_steps.get('model', pipeline)

        if scaler is not None:
            X_scaled = scaler.transform(feat_array)
        else:
            X_scaled = feat_array

        if "Voting" in model_choice:
            active_model = voting_clf
        elif "SVM" in model_choice:
            active_model = voting_clf.named_estimators_['svm']
        elif "Random Forest" in model_choice:
            active_model = voting_clf.named_estimators_['rf']
        else:
            active_model = voting_clf.named_estimators_['knn']

        pred = active_model.predict(X_scaled)[0]
        prob = active_model.predict_proba(X_scaled)[0]

        # 1. Card Hasil Utama
        with st.container(border=True):
            st.subheader("🔍 Hasil Diagnosis Prediktif")
            if pred == 1:
                st.error("### ⚠️ Terindikasi: PNEUMONIA")
                confidence = prob[1] * 100
            else:
                st.success("### ✅ Terindikasi: NORMAL")
                confidence = prob[0] * 100
                
            st.write(f"Tingkat Keyakinan Model ({model_choice.split('(')[0].strip()}): **{confidence:.2f}%**")
            st.progress(float(confidence / 100.0))

        # 2. Card Ekstraksi Fitur GLCM (Grid 3x2)
        with st.container(border=True):
            st.subheader("📊 Fitur Tekstur GLCM (Rerata 4 Arah)")
            m1, m2, m3 = st.columns(3)
            m1.metric("Contrast", f"{feat_dict['contrast']:.4f}")
            m2.metric("Dissimilarity", f"{feat_dict['dissimilarity']:.4f}")
            m3.metric("Homogeneity", f"{feat_dict['homogeneity']:.4f}")
            
            m4, m5, m6 = st.columns(3)
            m4.metric("Energy", f"{feat_dict['energy']:.4f}")
            m5.metric("Correlation", f"{feat_dict['correlation']:.4f}")
            m6.metric("ASM", f"{feat_dict['asm']:.4f}")

        # 3. Card Breakdown Model Individual (Jika memilih Voting)
        if "Voting" in model_choice and hasattr(voting_clf, 'named_estimators_'):
            with st.container(border=True):
                st.subheader("🤖 Keputusan Model Komponen")
                p_svm = "PNEUMONIA" if voting_clf.named_estimators_['svm'].predict(X_scaled)[0] == 1 else "NORMAL"
                p_rf  = "PNEUMONIA" if voting_clf.named_estimators_['rf'].predict(X_scaled)[0] == 1 else "NORMAL"
                p_knn = "PNEUMONIA" if voting_clf.named_estimators_['knn'].predict(X_scaled)[0] == 1 else "NORMAL"

                b1, b2, b3 = st.columns(3)
                with b1:
                    st.caption("**SVM Tuned**")
                    if p_svm == "PNEUMONIA": st.error("PNEUMONIA")
                    else: st.success("NORMAL")
                with b2:
                    st.caption("**Random Forest Tuned**")
                    if p_rf == "PNEUMONIA": st.error("PNEUMONIA")
                    else: st.success("NORMAL")
                with b3:
                    st.caption("**KNN Tuned**")
                    if p_knn == "PNEUMONIA": st.error("PNEUMONIA")
                    else: st.success("NORMAL")
    else:
        with st.container(border=True):
            st.subheader("🔍 Hasil Analisis & Prediksi")
            st.info("👈 Silakan unggah citra rontgen dada di panel sebelah kiri untuk melihat hasil klasifikasi dan fitur GLCM.")
