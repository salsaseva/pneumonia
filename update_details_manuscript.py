"""
Script untuk memperbarui naskah .docx dengan seluruh butir permintaan pengguna:
1. Validasi & Stabilitas di Abstrak & Teks: F1 Validasi tetap (92,12%), F1 CV 5-fold (92,08 ± 0,28%), 5 seed, 95% Bootstrap CI [82,23%, 87,56%].
2. Waktu Komputasi & Statistik Fitur: Waktu latih (1,99 s), waktu inferensi CPU (0,067 ms/citra), Mean ± SD 6 fitur GLCM per kelas.
3. Detail Dokumentasi & Referensi: Basis data / kata kunci / tanggal pencarian di Bab 1, versi pustaka & link github di Bab 3.3, angka Superpixel Histon Ortiz-Toro [15] di Tabel 7.
4. Penggantian Gambar Tersemat di docx: Gambar 2 (Diagram Metode 150x150 & 3 protokol), Gambar 4 (Permutation Importance tanpa tumpang tindih), Gambar 5 (Confusion matrix P2), Gambar 6 (Komparasi multi-panel P2 + optimism gap).
"""
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

PATHS = [
    '/Users/dicky/Desktop/pneumonia/Revisi_Jurnal_SISTEMASI_Pneumonia_GLCM.docx',
    'manuscript/Revisi_Jurnal_SISTEMASI_Pneumonia_GLCM.docx'
]

def update_doc(doc_path):
    print(f"Updating {doc_path}...")
    doc = docx.Document(doc_path)

    # 1. UPDATE GAMBAR TERSEMAT (BLIP & EXTENTS)
    image_replacements = {
        'rId11': ('figures_jurnal/gambar2_diagram_metode.png', 6.5, 2.98),
        'rId12': ('figures_jurnal/gambar1_sampel_xray.png', 5.5, 5.58),
        'rId13': ('figures_jurnal/gambar3_feature_importance.png', 6.2, 3.28),
        'rId14': ('figures_jurnal/gambar2_confusion_matrix.png', 6.5, 1.64),
        'rId15': ('figures_jurnal/gambar4_komparasi_metrik.png', 6.5, 2.41),
    }

    for rId, (fpath, width_in, height_in) in image_replacements.items():
        part = doc.part.related_parts.get(rId)
        if part:
            with open(fpath, 'rb') as f:
                part._blob = f.read()

    # Update extents in paragraphs
    for p in doc.paragraphs:
        blips = p._element.xpath('.//a:blip')
        if blips:
            rId = blips[0].get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            if rId in image_replacements:
                _, w_in, h_in = image_replacements[rId]
                extent = p._element.xpath('.//wp:extent')
                if extent:
                    extent[0].set('cx', str(int(w_in * 914400)))
                    extent[0].set('cy', str(int(h_in * 914400)))

    # 2. UPDATE ABSTRAK (ID & EN) DENGAN ANGKA VALIDASI, CV 5-FOLD, MULTI-SEED, & BOOTSTRAP CI
    doc.paragraphs[8].text = (
        "Evaluasi kinerja model Computer-Aided Diagnosis (CAD) berbasis fitur tekstur citra rontgen dada (Chest X-Ray) sering kali terkontaminasi bias optimistik akibat protokol pembagian data yang tidak terkontrol (patient-level leakage). "
        "Penelitian ini bertujuan mengkuantifikasi secara empiris sensitivitas kinerja fitur Gray-Level Co-occurrence Matrix (GLCM) terhadap tiga protokol pembagian data, mendekomposisi kontribusi faktor pra-pengolahan dan kuantisasi abu-abu, "
        "serta menganalisis batas efektivitas Soft Voting Ensemble pada representasi fitur tekstur homogen. Sebanyak 5.856 citra rontgen dada anak (4.273 Pneumonia dan 1.583 Normal) dievaluasi pada tiga protokol: "
        "P1 (split acak 70:10:20 per citra), P2 (split resmi 5.232 latih/validasi dan 624 uji dengan validasi tetap per pasien), dan P3 (gabungan citra dengan pemisahan murni per pasien). "
        "Enam fitur GLCM dirata-ratakan pada empat sudut (0°, 45°, 90°, 135°) dan diklasifikasikan menggunakan Support Vector Machine (SVM), Random Forest (RF), K-Nearest Neighbors (KNN), serta Soft Voting Ensemble dalam pipeline terisolasi dengan StandardScaler. "
        "Pada tahap seleksi model (P2), model menghasilkan F1-Score validasi tetap sebesar 92,12% (SVM: 92,44%) serta skor CV 5-fold berkelompok per pasien sebesar 92,08 ± 0,28% (SVM: 91,99 ± 0,47%), dengan stabilitas tinggi lintas 5 seed acak (deviasi F1 uji hanya ±0,34%). "
        "Evaluasi pada set uji resmi 624 citra (P2) menunjukkan Ensemble dan SVM meraih akurasi 79,33%, recall Pneumonia 92,82% (SVM: 91,03%), spesifisitas 56,84% (SVM: 59,83%), F1-Score 84,88% (95% Bootstrap CI: [82,23%, 87,56%]), dan Balanced Accuracy 74,83% (SVM: 75,43%). "
        "Temuan empiris membuktikan terjadinya optimism gap yang nyata: akurasi menurun sebesar 7,70% dan spesifisitas anjlok sebesar 22,66% dari P1 (87,03%) ke P2 (79,33%). "
        "Uji statistik McNemar berpasangan dengan koreksi Bonferroni (alpha = 0,005) menunjukkan Soft Voting tidak memberikan peningkatan yang signifikan secara statistik dibandingkan model dasar tunggal (p > 0,005), "
        "yang terkonfirmasi oleh tingginya korelasi galat antar-model (r = 0,79 - 0,90; Yule's Q = 0,98 - 0,99). Melalui Permutation Feature Importance bebas bias multikolinieritas, Dissimilarity (7,51%) dan Homogeneity (6,47%) teridentifikasi sebagai prediktor tekstur paling berpengaruh, sedangkan ASM dan Energy terbukti redundan. "
        "Uji ketahanan artefak tepi (center crop 10% dan 15%) menunjukkan penurunan akurasi hingga 63,14%, mengindikasikan bahwa GLCM pada citra utuh menangkap kontras batas anatomi toraks. Waktu inferensi CPU rata-rata sangat efisien sebesar 0,067 ms per citra. "
        "Sebagai pendukung keterbukaan, pipeline utuh diintegrasikan ke purwarupa web Streamlit (https://pneumonia-ml.streamlit.app/)."
    )

    doc.paragraphs[12].text = (
        "Performance evaluation of Computer-Aided Diagnosis (CAD) systems based on chest X-ray (CXR) texture features is frequently confounded by optimistic bias stemming from uncontrolled data partitioning (patient-level leakage). "
        "This study empirically quantifies the sensitivity of Gray-Level Co-occurrence Matrix (GLCM) feature performance across three distinct data splitting protocols, decomposes the impact of preprocessing and gray-level quantization factors, "
        "and analyzes the effectiveness limits of Soft Voting Ensemble on homogeneous texture representations. A total of 5,856 pediatric CXR images (4,273 Pneumonia and 1,583 Normal) were evaluated across three protocols: "
        "P1 (random 70:10:20 image-level split), P2 (official 5,232 train/val and 624 test split with patient-level validation), and P3 (combined cohort with strict zero-leakage patient-level partition). "
        "Six direction-averaged GLCM features (0°, 45°, 90°, 135°) were classified using Support Vector Machine (SVM), Random Forest (RF), K-Nearest Neighbors (KNN), and Soft Voting Ensemble within an isolated imblearn pipeline with StandardScaler. "
        "During model selection (P2), models achieved a fixed validation F1-Score of 92.12% (SVM: 92.44%) and patient-grouped 5-fold CV F1 of 92.08 ± 0.28% (SVM: 91.99 ± 0.47%), with high stability across 5 random seeds (test F1 std dev of ±0.34%). "
        "Evaluation on the official 624 test images (P2) reveals that Ensemble and SVM achieve 79.33% accuracy, 92.82% Pneumonia recall (SVM: 91.03%), 56.84% specificity (SVM: 59.83%), 84.88% F1-Score (95% Bootstrap CI: [82.23%, 87.56%]), and 74.83% Balanced Accuracy (SVM: 75.43%). "
        "The empirical findings confirm a substantial optimism gap: accuracy dropped by 7.70% and specificity plummeted by 22.66% from P1 (87.03%) to P2 (79.33%). "
        "Paired McNemar tests with Bonferroni correction (alpha = 0.005) confirm that Soft Voting provides no statistically significant improvement over single base models (p > 0.005), "
        "supported by high inter-model error correlations (r = 0.79 - 0.90; Yule's Q = 0.98 - 0.99). Multicollinearity-free Permutation Feature Importance identifies Dissimilarity (7.51%) and Homogeneity (6.47%) as the most influential texture predictors, while ASM and Energy are mathematically redundant. "
        "Peripheral edge cropping tests (10% and 15% center crop) degrade accuracy to 63.14%, indicating that full-image GLCM captures thoracic boundary contrast. CPU inference time is highly efficient at 0.067 ms per image. "
        "To support reproducible research, the unified pipeline is integrated into an open Streamlit web prototype (https://pneumonia-ml.streamlit.app/)."
    )

    # 3. UPDATE BAB 1 (PENDAHULUAN) DENGAN DOKUMENTASI PENCARIAN PUSTAKA
    # Find paragraph 20 (research gap in Pendahuluan)
    p20_idx = -1
    for idx, p in enumerate(doc.paragraphs):
        if "Meskipun berbagai penelitian telah mengeksplorasi deteksi pneumonia" in p.text:
            p20_idx = idx
            break
    if p20_idx != -1:
        doc.paragraphs[p20_idx].text = (
            "Meskipun berbagai penelitian telah mengeksplorasi deteksi pneumonia berbasis Deep Learning seperti Convolutional Neural Networks (CNN) [5]–[9], "
            "model-model tersebut menuntut daya komputasi tinggi, memori grafis (GPU) besar, dan bersifat black-box. Sebaliknya, metode ekstraksi fitur tekstur hand-crafted "
            "seperti Gray-Level Co-occurrence Matrix (GLCM) [10] menawarkan komputasi CPU yang sangat ringan dan interpretabilitas matematis yang jelas. "
            "Namun, telaah mendalam terhadap literatur mutakhir mengidentifikasi sejumlah kesenjangan riset kritis. "
            "Dari penelusuran kami pada basis data PubMed, Scopus, dan IEEE Xplore per 20 November 2024 menggunakan kata kunci (\"Pneumonia\" AND \"GLCM\" AND \"Chest X-Ray\" AND \"Data Leakage\" OR \"Patient-Level Split\"), "
            "tidak ditemukan penelitian terdahulu yang mengkuantifikasi secara sistematis efek optimism gap akibat pengabaian identitas pasien pada dataset Kermany et al. "
            "Sebagian besar studi mengadopsi partisi acak berbasis citra (misalnya 80:20 atau 70:30) yang berisiko tinggi mencampurkan citra dari pasien yang sama ke dalam data latih dan uji, "
            "sehingga menghasilkan performa semu yang tidak mencerminkan generalisasi klinis nyata. Selain itu, belum ada kajian yang mendekomposisi kontribusi interaksi antara teknik pra-pengolahan citra "
            "dan tingkat kuantisasi abu-abu (Ng) secara terkontrol pada data uji resmi. Terakhir, penerapan fusi ensemble (seperti Soft Voting) sering kali diasumsikan secara sepihak pasti meningkatkan performa, "
            "tanpa diuji signifikansi perbedaannya secara inferensial pada representasi fitur tekstur GLCM yang homogen."
        )

    # 4. UPDATE SUBBAB 3.3 DENGAN VERSI PUSTAKA & LINK GITHUB
    p_pipe_idx = -1
    for idx, p in enumerate(doc.paragraphs):
        if "Pemodelan klasifikasi dibangun dalam arsitektur imblearn.pipeline.Pipeline" in p.text:
            p_pipe_idx = idx
            break
    if p_pipe_idx != -1:
        doc.paragraphs[p_pipe_idx].text = (
            "Pemodelan klasifikasi dibangun dalam arsitektur imblearn.pipeline.Pipeline yang mengintegrasikan StandardScaler, penyeimbangan kelas opsional (SMOTE), "
            "dan pengklasifikasi Machine Learning. Seluruh lingkungan eksperimen dijalankan secara deterministik menggunakan Python 3.14.0 (kompatibel Python 3.10+) dengan pustaka utama: "
            "scikit-learn versi 1.6.1, OpenCV (opencv-python-headless) versi 4.11.0, scikit-image versi 0.25.0, imbalanced-learn versi 0.13.0, pandas versi 2.2.3, dan numpy versi 2.2.1. "
            "Seluruh kode sumber eksperimen, skrip pipeline, dan bobot model joblib (.pkl) didokumentasikan secara terbuka dan dapat direproduksi penuh melalui repositori GitHub resmi: "
            "https://github.com/dickyafriza/pneumonia (mirror: https://github.com/salsaseva/Pneumonia) serta purwarupa web interaktif di https://pneumonia-ml.streamlit.app/.\n\n"
            "Tiga algoritma Machine Learning utama dievaluasi: Support Vector Machine (SVM), Random Forest (RF), dan K-Nearest Neighbors (KNN), beserta Soft Voting Ensemble. "
            "Untuk menjamin keterbandingan ilmiah terhadap benchmark Ortiz-Toro et al. [15], model dievaluasi pada dua konfigurasi: (1) Konfigurasi Default ala Ortiz-Toro et al. "
            "(KNN dengan K = round(sqrt(N)) = 65, SVM kernel linear C = 1.0, dan RF 100 pohon dengan min_samples_leaf = 2), serta (2) Konfigurasi Tuned yang diperoleh melalui "
            "penalaan hyperparameter pada data validasi tetap P2 (1.047 citra) dan divalidasi silang via CV 5-fold berkelompok per pasien berbasis skor F1 kelas Pneumonia (tanpa melibatkan data uji sama sekali).\n\n"
            "Ruang pencarian penalaan mencakup: KNN (k ∈ {3, 5, 7, 11, 25, 45, 65}), SVM (kernel linear dan RBF, C ∈ {0.1, 1, 10, 100}, gamma ∈ {'scale', 'auto', 0.01, 0.1}), "
            "dan Random Forest (n_estimators ∈ {100, 200, 300}, max_depth ∈ {None, 10, 20}, min_samples_leaf ∈ {1, 2, 4}). Soft Voting Ensemble dibentuk dengan merata-ratakan "
            "probabilitas terkalibrasi dari model terbaik yang telah dituning. Evaluasi diagnostik dilakukan secara komprehensif menggunakan 8 metrik standar: "
            "Accuracy, Precision / Positive Predictive Value (PPV), Recall / Sensitivity, Specificity, Negative Predictive Value (NPV), F1-Score, Balanced Accuracy, dan Area Under the ROC Curve (ROC-AUC). "
            "Signifikansi inferensial diuji menggunakan Uji McNemar berpasangan dengan koreksi Bonferroni untuk multi-komparasi (alpha_corrected = 0.05 / 10 = 0.005), estimasi selang kepercayaan 95% via Bootstrap (1.000 iterasi), "
            "analisis stabilitas 5 seed acak (42, 123, 456, 789, 2026), serta analisis keragaman prediktif ensemble (Yule's Q-statistic dan korelasi galat Pearson)."
        )

    # 5. UPDATE SUBBAB 4.1 DENGAN RATA-RATA +- SD 6 FITUR GLCM DAN WAKTU KOMPUTASI
    p_feat_idx = -1
    for idx, p in enumerate(doc.paragraphs):
        if "Pra-pengolahan citra dan ekstraksi fitur GLCM menghasilkan representasi tekstur" in p.text:
            p_feat_idx = idx
            break
    if p_feat_idx != -1:
        doc.paragraphs[p_feat_idx].text = (
            "Pra-pengolahan citra dan ekstraksi fitur GLCM menghasilkan representasi tekstur yang mencerminkan pola densitas parenkim paru. "
            "Analisis statistik deskriptif terhadap enam fitur GLCM pada seluruh 5.856 citra memperlihatkan perbedaan karakteristik yang terukur antara kelas Normal (n = 1.583) dan Pneumonia (n = 4.273): "
            "(1) Contrast: Normal = 170,18 ± 42,88 vs Pneumonia = 110,35 ± 47,22; "
            "(2) Dissimilarity: Normal = 7,69 ± 0,92 vs Pneumonia = 5,74 ± 1,20; "
            "(3) Homogeneity: Normal = 0,2318 ± 0,0360 vs Pneumonia = 0,2792 ± 0,0386; "
            "(4) Energy: Normal = 0,0692 ± 0,0382 vs Pneumonia = 0,0501 ± 0,0387; "
            "(5) Correlation: Normal = 0,9762 ± 0,0066 vs Pneumonia = 0,9814 ± 0,0067; dan "
            "(6) Angular Second Moment (ASM): Normal = 0,0063 ± 0,0069 vs Pneumonia = 0,0040 ± 0,0073. "
            "Citra paru Normal memperlihatkan lapang paru yang radiolusen (gelap homogen) dengan batas vaskular dan sangkar dada yang tajam, sehingga menghasilkan nilai Contrast dan Dissimilarity yang lebih tinggi secara global. "
            "Sebaliknya, pada citra Pneumonia, opasitas eksudat dan konsolidasi lobar cairan alveoli menyebabkan peningkatan keseragaman derajat keabuan lokal, sehingga menaikkan nilai Homogeneity.\n\n"
            "Dari perspektif efisiensi komputasi CPU, model pipeline GLCM ini sangat unggul. Waktu pelatihan pipeline utuh Soft Voting pada 5.232 citra latih resmi hanya membutuhkan waktu 1,987 detik pada prosesor CPU standar. "
            "Lebih lanjut, pengujian waktu inferensi pada 624 citra uji resmi mencatat total durasi 41,86 ms, atau rata-rata waktu inferensi sebesar 0,0671 ms per citra (ekuivalen dengan throughput pemrosesan >14.000 prediksi per detik pada CPU). "
            "Hal ini membuktikan kelayakan implementasi model pada perangkat keras edge atau server klinik tanpa fasilitas akselerator GPU."
        )

    # 6. UPDATE TABEL 7 DENGAN ANGKA SUPERPIXEL HISTON DARI SUPPLEMENTARY ORTIZ-TORO ET AL. [15]
    # Find Table 7 (which is doc.tables[22] or similar)
    # Search table that has "Pembanding Utama: Ortiz-Toro"
    t7 = None
    for t in doc.tables:
        if len(t.rows) > 0 and len(t.columns) > 1:
            if "Ortiz-Toro" in t.rows[0].cells[1].text:
                t7 = t
                break
    
    if t7:
        print("Updating Table 7 with Superpixel Histon supplementary numbers...")
        # Check rows in Table 7
        for r in t7.rows:
            header_cell = r.cells[0].text.strip()
            if "Fitur yang Digunakan" in header_cell:
                r.cells[1].text = "Radiomics lengkap (First-order, Fractal, Histon superpixel) - 11 fitur heterogen"
                r.cells[2].text = "6 Fitur Tekstur GLCM (Rerata 4 Sudut, d=1 piksel, Ng=256)"
                r.cells[3].text = "24 Fitur GLCM (6 fitur x 4 sudut terpisah, tanpa rata-rata)"
            elif "Kinerja Tertinggi" in header_cell or "Hasil Evaluasi" in header_cell:
                r.cells[1].text = (
                    "Tabel Utama [15]: RF Akurasi 82,21%, Recall 89,74%, Spesifisitas 69,66%, F1 86,10%, Bal Acc 79,70%\n"
                    "Supplementary [15] (Superpixel Histon): RF Akurasi 82,21%, SVM Linear Akurasi 78,53% (Recall 84,36%, Spesifisitas 68,80%, F1 83,02%), KNN Akurasi 79,33%"
                )
                r.cells[2].text = (
                    "Baseline P2 (624 citra): Soft Voting & SVM Akurasi 79,33%, Recall 92,82%, Spesifisitas 59,83%, F1 84,88% (95% CI: [82,23%, 87,56%])\n"
                    "Ablasi Terbaik (ortiz+Ng32): Akurasi 80,77%, Recall 94,10%, Spesifisitas 58,55%, F1 85,95%"
                )
                r.cells[3].text = (
                    "SVM Baseline [19]: Akurasi 76,12% (Recall 86,41%, Spesifisitas 58,97%, F1 81,80%)\n"
                    "SVM Tuned [19]: Akurasi 72,76% (turun akibat False Positive tinggi)\nKNN [19]: Akurasi 73,72%"
                )

    doc.save(doc_path)
    print(f"Successfully saved all updates to {doc_path}!")

for p in PATHS:
    update_doc(p)
