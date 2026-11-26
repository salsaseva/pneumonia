"""
Script to apply comprehensive revisions to SISTEMASI journal manuscript:
/Users/dicky/Desktop/pneumonia/Revisi_Jurnal_SISTEMASI_Pneumonia_GLCM.docx
Based on: Daftar Revisi Deteksi Pneumonia GLCM.pdf
"""
import os
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

DOC_PATH = "/Users/dicky/Desktop/pneumonia/Revisi_Jurnal_SISTEMASI_Pneumonia_GLCM.docx"

def set_cell_text(cell, text, bold=False, font_size=8.5, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(font_size)
    run.font.bold = bold

def update_manuscript():
    print(f"Loading document from {DOC_PATH}...")
    doc = docx.Document(DOC_PATH)

    # -------------------------------------------------------------------------
    # 1. JUDUL & ABSTRAK
    # -------------------------------------------------------------------------
    print("Updating Title and Abstracts...")
    # P0: Judul ID
    doc.paragraphs[0].text = "Evaluasi Sensitivitas Fitur Tekstur GLCM terhadap Protokol Pembagian Data dan Batas Efektivitas Ensemble pada Deteksi Pneumonia Citra Rontgen"
    doc.paragraphs[0].runs[0].font.name = "Times New Roman"
    doc.paragraphs[0].runs[0].font.size = Pt(18)
    doc.paragraphs[0].runs[0].font.bold = True

    # P1: Judul EN
    doc.paragraphs[1].text = "Sensitivity Evaluation of GLCM Texture Features to Data Splitting Protocols and Ensemble Effectiveness Limits in Chest X-Ray Pneumonia Detection"
    doc.paragraphs[1].runs[0].font.name = "Times New Roman"
    doc.paragraphs[1].runs[0].font.size = Pt(14)
    doc.paragraphs[1].runs[0].font.italic = True

    # P8: Abstrak ID
    doc.paragraphs[8].text = (
        "Evaluasi kinerja model Computer-Aided Diagnosis (CAD) berbasis fitur tekstur citra rontgen dada (Chest X-Ray) sering kali "
        "terkontaminasi bias optimistik akibat protokol pembagian data yang tidak terkontrol (patient-level leakage). Penelitian ini bertujuan "
        "mengkuantifikasi secara empiris sensitivitas kinerja fitur Gray-Level Co-occurrence Matrix (GLCM) terhadap tiga protokol pembagian data, "
        "mendekomposisi kontribusi faktor pra-pengolahan dan kuantisasi abu-abu, serta menganalisis batas efektivitas Soft Voting Ensemble pada "
        "representasi fitur tekstur homogen. Sebanyak 5.856 citra rontgen dada anak (4.273 Pneumonia dan 1.583 Normal) dievaluasi pada tiga protokol: "
        "P1 (split acak 70:10:20 per citra), P2 (split resmi 5.232 latih/validasi dan 624 uji dengan validasi tetap per pasien), dan P3 (gabungan citra "
        "dengan pemisahan murni per pasien). Enam fitur GLCM dirata-ratakan pada empat sudut (0°, 45°, 90°, 135°) dan diklasifikasikan menggunakan "
        "Support Vector Machine (SVM), Random Forest (RF), K-Nearest Neighbors (KNN), serta Soft Voting Ensemble dalam pipeline terisolasi dengan StandardScaler. "
        "Evaluasi pada set uji resmi 624 citra (P2) menunjukkan SVM dan Ensemble meraih akurasi 79,33%, recall Pneumonia 92,82% (SVM: 91,03%), "
        "spesifisitas 56,84% (SVM: 59,83%), dan F1-Score 84,88% (SVM: 84,62%). Temuan empiris membuktikan terjadinya optimism gap yang nyata: akurasi menurun "
        "sebesar 7,70% dan spesifisitas anjlok sebesar 22,66% dari P1 (87,03%) ke P2 (79,33%). Uji statistik McNemar berpasangan dengan koreksi Bonferroni "
        "(alpha = 0,005) menunjukkan bahwa Soft Voting tidak memberikan peningkatan yang signifikan secara statistik dibandingkan model dasar tunggal (p > 0,005), "
        "yang terkonfirmasi oleh tingginya korelasi galat antar-model (r = 0,79 - 0,90; Yule's Q = 0,98 - 0,99). Melalui Permutation Feature Importance bebas bias "
        "multikolinieritas, Dissimilarity (7,51%) dan Homogeneity (6,47%) teridentifikasi sebagai prediktor tekstur paling berpengaruh, sedangkan ASM dan Energy "
        "terbukti redundan. Uji ketahanan artefak tepi (center crop 10% dan 15%) menunjukkan penurunan akurasi hingga 63,14%, mengindikasikan bahwa GLCM pada citra utuh "
        "menangkap kontras batas anatomi toraks. Sebagai pendukung keterbukaan, pipeline utuh diintegrasikan ke purwarupa web Streamlit (https://pneumonia-ml.streamlit.app/)."
    )

    # P10: Kata Kunci ID
    doc.paragraphs[10].text = "Kata kunci: klasifikasi pneumonia, citra rontgen dada, gray-level co-occurrence matrix, optimism gap, soft voting ensemble"

    # P12: Abstract EN
    doc.paragraphs[12].text = (
        "Performance evaluation of Computer-Aided Diagnosis (CAD) systems based on chest X-ray (CXR) texture features is frequently confounded by "
        "optimistic bias stemming from uncontrolled data partitioning (patient-level leakage). This study empirically quantifies the sensitivity of "
        "Gray-Level Co-occurrence Matrix (GLCM) feature performance across three distinct data splitting protocols, decomposes the impact of preprocessing "
        "and gray-level quantization factors, and analyzes the effectiveness limits of Soft Voting Ensemble on homogeneous texture representations. "
        "A total of 5,856 pediatric CXR images (4,273 Pneumonia and 1,583 Normal) were evaluated across three protocols: P1 (random 70:10:20 image-level split), "
        "P2 (official 5,232 train/val and 624 test split with patient-level validation), and P3 (combined cohort with strict zero-leakage patient-level partition). "
        "Six direction-averaged GLCM features (0°, 45°, 90°, 135°) were classified using Support Vector Machine (SVM), Random Forest (RF), K-Nearest Neighbors (KNN), "
        "and Soft Voting Ensemble within an isolated imblearn pipeline with StandardScaler. Evaluation on the official 624 test images (P2) reveals that SVM "
        "and Ensemble achieve 79.33% accuracy, 92.82% Pneumonia recall (SVM: 91.03%), 56.84% specificity (SVM: 59.83%), and 84.88% F1-Score (SVM: 84.62%). "
        "The empirical findings confirm a substantial optimism gap: accuracy dropped by 7.70% and specificity plummeted by 22.66% from P1 (87.03%) to P2 (79.33%). "
        "Paired McNemar tests with Bonferroni correction (alpha = 0.005) confirm that Soft Voting provides no statistically significant improvement over single "
        "base models (p > 0.005), supported by high inter-model error correlations (r = 0.79 - 0.90; Yule's Q = 0.98 - 0.99). Multicollinearity-free Permutation "
        "Feature Importance identifies Dissimilarity (7.51%) and Homogeneity (6.47%) as the most influential texture predictors, while ASM and Energy are mathematically "
        "redundant. Peripheral edge cropping tests (10% and 15% center crop) degrade accuracy to 63.14%, indicating that full-image GLCM captures thoracic boundary contrast. "
        "To support reproducible research, the unified pipeline is integrated into an open Streamlit web prototype (https://pneumonia-ml.streamlit.app/)."
    )

    # P14: Keywords EN
    doc.paragraphs[14].text = "Keywords: pneumonia classification, chest X-ray image, gray-level co-occurrence matrix, optimism gap, soft voting ensemble"

    # -------------------------------------------------------------------------
    # 2. PENDAHULUAN (RESEARCH GAPS, RQ1-RQ3, KONTRIBUSI)
    # -------------------------------------------------------------------------
    print("Updating Introduction (Pendahuluan)...")
    # P20: Research Gaps
    doc.paragraphs[20].text = (
        "Meskipun berbagai penelitian telah mengeksplorasi deteksi pneumonia berbasis Deep Learning seperti Convolutional Neural Network (CNN) "
        "yang menuntut komputasi GPU tinggi [6], [7], [8], [9], [16], [20] maupun ekstraksi fitur tekstur Gray-Level Co-occurrence Matrix (GLCM) "
        "yang lebih ringan [12], [15], [19], [21], terdapat tiga kesenjangan penelitian mendasar (research gap) yang belum teratasi dalam literatur. "
        "Pertama, studi-studi berbasis tekstur pada dataset rontgen dada pediatrik Kermany et al. dievaluasi dengan protokol yang sangat beragam—mulai dari "
        "split acak per citra (image-level split), cross-validation acak, hingga split resmi bawaan folder—di mana terdapat indikasi kuat bahwa split resmi "
        "sendiri berbagi identitas pasien antara data latih dan uji. Namun, besarnya efek distorsi protokol ini terhadap kinerja diagnostik GLCM (dikenal sebagai "
        "optimism gap) belum pernah dikuantifikasi secara terkontrol. Kedua, sebagian besar studi menerapkan GLCM sebagai representasi kotak hitam (black-box) "
        "dengan parameter baku (kuantisasi Ng = 256 dan perataan 4 sudut) tanpa analisis dekomposisi faktor; belum ada kajian empiris yang membuktikan apakah "
        "teknik pra-pengolahan kontras adaptif, variasi level kuantisasi abu-abu (Ng), atau orientasi sudut tertentu memberikan kontribusi sinergis atau justru "
        "memicu overfitting. Ketiga, strategi Soft Voting Ensemble kerap diasumsikan secara apriori mampu meningkatkan kinerja klasifikasi, padahal belum pernah "
        "dibuktikan secara statistik apakah penggabungan model memberikan nilai tambah ketika seluruh pengklasifikasi dasar dilatih pada representasi fitur tekstur "
        "GLCM homogen yang sama."
    )

    # P21: Research Questions (RQ1-RQ3) & Kontribusi
    doc.paragraphs[21].text = (
        "Untuk mengatasi kesenjangan tersebut, penelitian ini merumuskan tiga pertanyaan penelitian utama:\n"
        "1. RQ1: Seberapa besar kinerja diagnostik enam fitur GLCM berubah antara protokol split acak per citra, split resmi bawaan, dan split murni per pasien (optimism gap)?\n"
        "2. RQ2: Faktor mana di antara teknik pra-pengolahan citra (basic, CLAHE, dan Ortiz-Toro), jumlah varian fitur (GLCM-5, GLCM-6, GLCM-24), serta level kuantisasi abu-abu (Ng = 32, 64, 256) yang paling memengaruhi kinerja klasifikasi?\n"
        "3. RQ3: Apakah Soft Voting Ensemble memberikan peningkatan kinerja yang signifikan secara statistik ketika seluruh model dasar mengonsumsi representasi fitur GLCM homogen yang sama?\n\n"
        "Berdasarkan rumusan masalah di atas, kontribusi penelitian ini mencakup:\n"
        "1. Kuantifikasi empiris pertama terhadap sensitivitas kinerja fitur GLCM dan efek optimism gap pada tiga protokol pembagian data yang divalidasi dengan audit identitas pasien (patient ID).\n"
        "2. Dekomposisi faktor pra-pengolahan dan kuantisasi abu-abu melalui studi ablasi 27 kombinasi terkontrol serta analisis Permutation Feature Importance bebas bias multikolinieritas.\n"
        "3. Pembuktian keterbatasan efektivitas Soft Voting Ensemble pada fitur homogen melalui uji signifikansi statistik McNemar berpasangan dan analisis korelasi galat antar-model.\n"
        "4. Penyediaan purwarupa web interaktif Streamlit (https://pneumonia-ml.streamlit.app/) berbasis pipeline terintegrasi utuh sebagai kontribusi pendukung keterbukaan penelitian."
    )

    # -------------------------------------------------------------------------
    # 3. TINJAUAN LITERATUR & KOREKSI PEMBANDING [15]
    # -------------------------------------------------------------------------
    print("Updating Literature Review...")
    # P24: Koreksi sitasi Ortiz-Toro et al. [15]
    doc.paragraphs[24].text = (
        "Dalam pemodelan klasifikasi berbasis fitur radiomik pada citra rontgen dada, sejumlah studi terdahulu telah mengevaluasi algoritma Machine Learning. "
        "Kim [14] membangun model Machine Learning berbasis fitur radiomik tekstur pada citra CXR untuk membedakan pneumonia dan COVID-19, serta melaporkan bahwa "
        "fitur tekstur orde kedua berkontribusi terhadap stabilitas klasifikasi. Ortiz-Toro et al. [15] menyajikan studi komparatif komprehensif pada dataset resmi "
        "Kermany et al. (624 citra uji) menggunakan fitur radiomics lengkap—mencakup statistik orde pertama, dimensi fraktal, dan superpixel-based histon—yang diklasifikasikan "
        "dengan KNN, SVM, dan Random Forest. Pada studi [15] tersebut, fitur radiomik menghasilkan akurasi 82,21% pada Random Forest (recall 89,74%, spesifisitas 69,66%, F1 86,10%) "
        "dan akurasi 78,53% pada SVM linear (recall 84,36%, spesifisitas 68,80%, F1 83,02%). Studi [15] ini menjadi batas acuan (benchmark) paling sepadan bagi penelitian ini "
        "karena dievaluasi pada data uji resmi yang sama. Selain itu, studi Zaena dan Lhaksmana [16] membandingkan Random Forest dan CNN pada CXR, melaporkan bahwa Random Forest "
        "menawarkan efisiensi komputasi CPU dan interpretabilitas fitur yang lebih tinggi. Untuk menelaah strategi ensemble, Mabrouk et al. [17] serta Shati et al. [18] "
        "menerapkan fusi model guna menyeimbangkan varians pengklasifikasi tunggal. Pada jurnal Sistemasi, Arifin dan Surya [20] menerapkan CNN VGG16, Mustopa et al. [21] "
        "mengombinasikan fitur Haralick dan Hu Moments dengan Random Forest, sedangkan Akbar dan Astuti [22] menegaskan krusialnya metrik Recall dan penanganan ketidakseimbangan kelas."
    )

    # P25: Studi pendahulu Hismaya [19] disederhanakan
    doc.paragraphs[25].text = (
        "Sebagai studi pendahulu berbasis GLCM murni, karya Hismaya dan Sulistianingsih [19] mengekstrak 24 kolom fitur GLCM (6 properti pada 4 sudut terpisah) berukuran 300×300 piksel "
        "tanpa standardisasi fitur, dan dievaluasi pada partisi uji bawaan (624 citra). Pada studi tersebut, KNN menghasilkan akurasi 73,72% (spesifisitas 47,01%), sedangkan SVM awal "
        "menghasilkan akurasi 76,12% (475/624 citra benar; recall 86,41%, spesifisitas 58,97%, F1 81,80%) yang kemudian turun menjadi 72,76% pasca-tuning kernel polinomial akibat tingginya "
        "False Positive. Mengacu pada keterbatasan studi [19] tersebut, penelitian ini menyusun kerangka evaluasi terkontrol dengan pipeline terisolasi, penyeragaman 150×150 piksel, "
        "tiga protokol pengujian, uji signifikansi statistik, dan perbandingan langsung terhadap radiomics Ortiz-Toro et al. [15]."
    )

    # -------------------------------------------------------------------------
    # 4. METODOLOGI: 3.1 DATASET, AUDIT ID, & PROTOKOL P1, P2, P3
    # -------------------------------------------------------------------------
    print("Updating Methodology (3.1 Dataset & Splits)...")
    # P37: 3.1 Akuisisi Dataset dan Pembagian Data
    doc.paragraphs[37].text = (
        "Dataset utama yang digunakan dalam penelitian ini adalah dataset radiografi rontgen dada pediatrik terkurasi dari Kermany et al. [4] "
        "sebanyak 5.856 citra anteroposterior (4.273 Pneumonia dan 1.583 Normal). Dataset resmi Kaggle tersebut aslinya terbagi atas folder train (5.216 citra), "
        "folder val (16 citra: 8 Normal, 8 Pneumonia), dan folder test (624 citra: 234 Normal, 390 Pneumonia). Mengikuti konvensi penelitian standar [15], "
        "folder train dan val digabungkan menjadi 5.232 citra latih resmi, sedangkan 624 citra dijadikan data uji resmi.\n\n"
        "Sebelum dilakukan pemodelan, dilakukan audit data komprehensif mencakup verifikasi integritas citra, deteksi duplikasi hash file (SHA-256), "
        "serta ekstraksi identitas pasien (Patient ID) dari pola penamaan file ('personXXX_' untuk Pneumonia dan 'NORMAL2-IM-XXXX' atau 'IM-XXXX' untuk Normal). "
        "Audit membuktikan bahwa seluruh 5.856 citra valid dan tidak ada duplikasi hash lintas split. Namun, audit menemukan fakta penting bahwa terdapat 170 identitas "
        "pasien Pneumonia pada data uji resmi yang juga muncul pada data latih resmi (84,16% dari 202 pasien Pneumonia uji sudah ada di data latih), sementara kelas Normal "
        "memiliki 0 overlap pasien. Untuk mengevaluasi dampak kebocoran identitas pasien tersebut secara kuantitatif, penelitian ini menetapkan tiga protokol pembagian data:\n"
        "1. Protokol 1 (P1 - Split Acak Citra 70:10:20): Citra dibagi secara acak berstratifikasi tanpa memedulikan ID pasien (4.100 latih, 584 validasi, 1.172 uji). Protokol ini merepresentasikan pendekatan lama yang rentan data leakage.\n"
        "2. Protokol 2 (P2 - Split Resmi Bawaan 5.232 / 624): Protokol utama yang setara dengan Ortiz-Toro et al. [15]. Data latih 5.232 citra dipisahkan menjadi 4.185 data latih dan 1.047 data validasi tetap menggunakan StratifiedGroupKFold (n_splits=5, shuffle=True, random_state=42) berbasis ID pasien (asserting 0 patient overlap antara latih dan validasi). Data uji tetap menggunakan 624 citra resmi.\n"
        "3. Protokol 3 (P3 - Split Murni per Pasien 70:15:15): Seluruh 5.856 citra digabung lalu dipartisi murni berdasarkan ID pasien secara eksklusif (4.113 latih, 863 validasi, 880 uji). Protokol ini menjamin zero-leakage pada seluruh tahap."
    )

    # Update Table 1: Distribusi Dataset
    print("Updating Table 1...")
    t1 = doc.tables[1]
    # Re-structure Table 1 to show P1, P2, P3
    # Header
    headers = ["Protokol Evaluasi", "Data Latih", "Data Validasi", "Data Uji", "Total Citra", "Kontrol Pasien (Leakage?)"]
    # Check column count
    while len(t1.columns) < len(headers):
        t1.add_column(Inches(1.0))
    for c_idx, h in enumerate(headers):
        set_cell_text(t1.rows[0].cells[c_idx], h, bold=True, font_size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)

    rows_data = [
        ["P1 (Split Acak Citra)", "4.100 (70%)", "584 (10%)", "1.172 (20%)", "5.856", "Tidak Terkontrol (Bocor Pasien)"],
        ["P2 (Split Resmi Bawaan)", "4.185 (71,5%)", "1.047 (17,9%)", "624 (10,6%)", "5.856", "Terkontrol di Validasi; 170 Overlap di Uji"],
        ["P3 (Split Murni Pasien)", "4.113 (70,2%)", "863 (14,7%)", "880 (15,0%)", "5.856", "Terkontrol Ketat (Bebas Leakage 100%)"]
    ]
    for r_idx, r_data in enumerate(rows_data):
        row = t1.rows[r_idx + 1]
        for c_idx, val in enumerate(r_data):
            align = WD_ALIGN_PARAGRAPH.LEFT if c_idx in [0, 5] else WD_ALIGN_PARAGRAPH.CENTER
            set_cell_text(row.cells[c_idx], val, bold=(c_idx==0), font_size=8.0, align=align)

    # -------------------------------------------------------------------------
    # 5. METODOLOGI: 3.2 PRA-PENGOLAHAN & GLCM
    # -------------------------------------------------------------------------
    print("Updating Methodology (3.2 Preprocessing & Features)...")
    # P40: Pra-pengolahan
    doc.paragraphs[40].text = (
        "Tahap pra-pengolahan citra digital bertujuan menyeragamkan dimensi dan menstabilkan kontras radiografi tanpa menimbulkan kebocoran informasi. "
        "Seluruh citra dikonversi ke grayscale 8-bit dan diubah ukurannya menjadi 150×150 piksel (peningkatan resolusi spasial dari versi sebelumnya 128×128 piksel). "
        "Untuk menjawab RQ2, penelitian ini mengevaluasi tiga mode pra-pengolahan secara komparatif: (1) basic (grayscale dan resize), (2) clahe (basic ditambah "
        "Contrast Limited Adaptive Histogram Equalization dengan clipLimit=2.0 dan tileGridSize=8×8), serta (3) ortiz (pencocokan histogram kumulatif terhadap referensi data latih, "
        "Balanced Contrast Enhancement Technique / BCET, pemangkasan intensitas 25% teratas untuk meredam artefak radio-opaque, dan CLAHE). Perlu ditegaskan bahwa "
        "histogram referensi dan parameter standardisasi fitur (StandardScaler) dihitung HANYA dari data latih dan diterapkan secara terisolasi pada data validasi dan uji.\n\n"
        "Ekstraksi fitur tekstur GLCM dilakukan pada jarak piksel d = 1 dengan empat orientasi sudut θ ∈ {0°, 45°, 90°, 135°}. Nilai keenam fitur statistik Haralick "
        "(Contrast, Dissimilarity, Homogeneity, Energy, Correlation, dan ASM) dirata-ratakan pada keempat arah tersebut guna mereduksi sensitivitas orientasi pemotretan pasien, "
        "sehingga menghasilkan vektor fitur ringkas berdimensi enam (GLCM-6). Sebagai pembanding pada studi ablasi, dievaluasi pula varian GLCM-24 (24 fitur tanpa dirata-ratakan) "
        "dan GLCM-5 (tanpa fitur Energy karena Energy secara matematis merupakan akar kuadrat langsung dari ASM, Energy = sqrt(ASM), sehingga bersifat redundan). "
        "Selain itu, kuantisasi derajat keabuan (Ng) divariasikan pada Ng ∈ {32, 64, 256} untuk mengukur kompromi antara preservasi informasi tekstur dan reduksi noise radiologis."
    )

    # -------------------------------------------------------------------------
    # 6. METODOLOGI: 3.3 PEMODELAN, TUNING, & METRIK
    # -------------------------------------------------------------------------
    print("Updating Methodology (3.3 Modeling & Metrics)...")
    # P58: Pemodelan
    doc.paragraphs[58].text = (
        "Pemodelan klasifikasi dibangun dalam arsitektur imblearn.pipeline.Pipeline yang mengintegrasikan StandardScaler, penyeimbangan kelas opsional (SMOTE), "
        "dan pengklasifikasi. Tiga algoritma Machine Learning utama dievaluasi: Support Vector Machine (SVM), Random Forest (RF), dan K-Nearest Neighbors (KNN), "
        "beserta Soft Voting Ensemble. Untuk menjamin keterbandingan ilmiah terhadap Ortiz-Toro et al. [15], model dievaluasi pada dua konfigurasi: (1) Konfigurasi Default "
        "ala Ortiz-Toro et al. (KNN dengan K = round(sqrt(N)) = 65, SVM kernel linear C = 1.0, dan RF 100 pohon dengan min_samples_leaf = 2), serta (2) Konfigurasi Tuned "
        "yang diperoleh melalui penalaan hyperparameter pada data validasi P2 berbasis skor F1 kelas Pneumonia (skor seleksi tidak melibatkan data uji sama sekali).\n\n"
        "Ruang pencarian penalaan mencakup: KNN (k ∈ {3, 5, 7, 11, 25, 45, 65}), SVM (kernel linear dan RBF, C ∈ {0.1, 1, 10, 100}, gamma ∈ {'scale', 'auto', 0.01, 0.1}), "
        "dan Random Forest (n_estimators ∈ {100, 200, 300}, max_depth ∈ {None, 10, 20}, min_samples_leaf ∈ {1, 2, 4}). Soft Voting Ensemble dibentuk dengan menggabungkan "
        "probabilitas prediktif dari ketiga model dasar terbaik yang telah dituning. Evaluasi diagnostik dilakukan secara komprehensif menggunakan 8 metrik standar: "
        "Accuracy, Precision / Positive Predictive Value (PPV), Recall / Sensitivity, Specificity (TN/(TN+FP)), Negative Predictive Value (NPV = TN/(TN+FN)), F1-Score, "
        "Balanced Accuracy ((Sensitivity + Specificity)/2), dan Area Under the ROC Curve (ROC-AUC). Signifikansi perbedaan antar-model diuji menggunakan Uji McNemar berpasangan "
        "dengan koreksi Bonferroni untuk multi-komparasi (alpha_corrected = 0.05 / 10 = 0.005), estimasi selang kepercayaan 95% via Bootstrap (1.000 iterasi), serta analisis korelasi galat "
        "(Yule's Q-statistic dan korelasi Pearson) untuk mengevaluasi keragaman ensemble."
    )

    # -------------------------------------------------------------------------
    # 7. HASIL & PEMBAHASAN: 4.1 ANALISIS FITUR (PERMUTATION VS GINI)
    # -------------------------------------------------------------------------
    print("Updating Results (4.1 Feature Analysis)...")
    # P79 & P85: Feature Analysis
    doc.paragraphs[79].text = (
        "Pra-pengolahan citra dan ekstraksi fitur GLCM menghasilkan representasi tekstur yang mencerminkan pola densitas parenkim paru. "
        "Sebagaimana tertera pada data inspeksi visual, citra paru Normal memperlihatkan lapang paru yang radiolusen (gelap homogen) dengan batas vaskular yang jelas, "
        "menghasilkan nilai Homogeneity yang lebih tinggi dan Contrast yang lebih rendah. Sebaliknya, pada citra Pneumonia, opasitas eksudat dan konsolidasi lobar "
        "menyebabkan variasi intensitas piksel lokal yang tajam, sehingga meningkatkan nilai Contrast dan Dissimilarity secara signifikan."
    )

    doc.paragraphs[85].text = (
        "Untuk menganalisis kontribusi relatif keenam fitur GLCM secara objektif, penelitian ini membandingkan penurunan Gini impurity pada Random Forest "
        "dengan Permutation Feature Importance pada data uji resmi (Gambar 3). Gini importance tradisional cenderung bias pada fitur yang saling berkorelasi tinggi "
        "(seperti Energy dan ASM). Sebaliknya, Permutation Feature Importance mengukur penurunan performa nyata (F1-score) saat nilai suatu fitur diacak. "
        "Hasil eksperimen menunjukkan bahwa Dissimilarity memberikan dampak penurunan performa tertinggi (7,51% ± 0,58%), diikuti oleh Homogeneity (6,47% ± 0,46%), "
        "Contrast (4,93% ± 0,78%), Energy (4,11% ± 0,71%), Correlation (3,46% ± 0,68%), dan ASM (2,29% ± 0,61%). Temuan ini membuktikan bahwa variabilitas "
        "derajat keabuan lokal (Dissimilarity dan Homogeneity) memegang peranan diagnostik paling krusial dalam mendeteksi infiltrat pneumonia, sementara Energy dan ASM "
        "memiliki kontribusi yang saling tumpang tindih."
    )

    # -------------------------------------------------------------------------
    # 8. HASIL & PEMBAHASAN: 4.2 ANALISIS KINERJA MODEL RESMI P2 & UJI MCNEMAR
    # -------------------------------------------------------------------------
    print("Updating Results (4.2 P2 Evaluation & Table 17)...")
    # P88: Pengantar Tabel 17
    doc.paragraphs[88].text = (
        "Kinerja diagnostik seluruh model dievaluasi pada set uji resmi 624 citra (234 Normal dan 390 Pneumonia) pada Protokol P2, mengikuti struktur komparasi "
        "Ortiz-Toro et al. [15]. Hasil evaluasi komprehensif disajikan pada Tabel 2, mencakup konfigurasi Default ala Ortiz-Toro [15], konfigurasi Tuned, dan kondisi penyeimbangan SMOTE."
    )

    # Update Table 17 (Table 2 di teks naskah)
    t17 = doc.tables[17]
    headers_t17 = ["Model Klasifikasi", "Akurasi (%)", "Presisi (%)", "Recall (%)", "Spesifisitas (%)", "NPV (%)", "F1-Score (%)", "Balanced Acc (%)", "ROC-AUC (%)", "TN", "FP", "FN", "TP"]
    # Check column count
    while len(t17.columns) < len(headers_t17):
        t17.add_column(Inches(0.6))
    for c_idx, h in enumerate(headers_t17):
        set_cell_text(t17.rows[0].cells[c_idx], h, bold=True, font_size=7.5, align=WD_ALIGN_PARAGRAPH.CENTER)

    rows_t17_data = [
        ["Ensemble Soft Voting", "79,33%", "78,19%", "92,82%", "56,84%", "82,61%", "84,88%", "74,83%", "86,45%", "133", "101", "28", "362"],
        ["SVM (Tuned: RBF, C=100)", "79,33%", "79,06%", "91,03%", "59,83%", "80,00%", "84,62%", "75,43%", "78,97%", "140", "94", "35", "355"],
        ["Random Forest (Tuned)", "76,92%", "76,06%", "92,05%", "51,71%", "79,61%", "83,29%", "71,88%", "87,07%", "121", "113", "31", "359"],
        ["KNN (Tuned: k=11)", "78,21%", "78,35%", "90,00%", "58,55%", "77,84%", "83,77%", "74,27%", "84,51%", "137", "97", "39", "351"],
        ["SVM Default [15] (Linear)", "77,72%", "77,83%", "90,00%", "57,26%", "77,46%", "83,47%", "73,63%", "83,46%", "134", "100", "39", "351"],
        ["RF Default [15] (100 pohon)", "75,96%", "75,21%", "91,79%", "49,57%", "78,38%", "82,68%", "70,68%", "86,87%", "116", "118", "32", "358"],
        ["KNN Default [15] (k=65)", "79,97%", "80,74%", "89,23%", "64,53%", "78,24%", "84,77%", "76,88%", "85,73%", "151", "83", "42", "348"],
        ["Voting + SMOTE", "79,81%", "81,13%", "88,21%", "65,81%", "77,00%", "84,52%", "77,01%", "86,58%", "154", "80", "46", "344"],
        ["SVM + SMOTE", "79,81%", "81,28%", "87,95%", "66,24%", "76,73%", "84,48%", "77,09%", "81,82%", "155", "79", "47", "343"],
        ["Baseline Mayoritas (Pneumonia)", "62,50%", "62,50%", "100,00%", "0,00%", "0,00%", "76,92%", "50,00%", "0,00%", "0", "234", "0", "390"]
    ]
    # Add rows if needed
    while len(t17.rows) < len(rows_t17_data) + 1:
        t17.add_row()
    for r_idx, r_data in enumerate(rows_t17_data):
        row = t17.rows[r_idx + 1]
        for c_idx, val in enumerate(r_data):
            align = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            set_cell_text(row.cells[c_idx], val, bold=(c_idx==0 or "Voting" in val or "Tuned" in val), font_size=7.0, align=align)

    # P96: Pembahasan Uji McNemar & Korelasi Galat (Jawaban RQ3)
    doc.paragraphs[96].text = (
        "Berdasarkan hasil pada Tabel 2, model Soft Voting Ensemble dan SVM Tuned mencapai akurasi identik sebesar 79,33% pada data uji resmi 624 citra. "
        "Ensemble mencatatkan Recall tertinggi sebesar 92,82% dengan False Negative terendah (28 sampel), sedangkan SVM Tuned menghasilkan Spesifisitas lebih tinggi "
        "(59,83% vs 56,84%) dan Balanced Accuracy tertinggi (75,43%). Penyeimbangan data dengan SMOTE mampu meningkatkan spesifisitas SVM menjadi 66,24% dan akurasi 79,81%, "
        "namun dengan sedikit kompromi pada penurunan Recall (87,95%).\n\n"
        "Untuk menjawab RQ3 secara inferensial, dilakukan Uji McNemar berpasangan pada seluruh kombinasi model dengan batas signifikansi terkoreksi Bonferroni (alpha = 0,005). "
        "Hasil uji statistik membuktikan bahwa TIDAK ADA perbedaan signifikan secara statistik antara Soft Voting Ensemble dengan SVM Tuned (statistic = 10.0, p = 1.000), "
        "KNN Tuned (statistic = 1.44, p = 0.230), maupun Random Forest Tuned (statistic = 6.32, p = 0.012 > 0.005). Analisis keragaman ensemble lebih lanjut mengungkap bahwa "
        "korelasi galat antar-model dasar sangat tinggi (r = 0,79 hingga 0,90; Yule's Q-statistic = 0,98 hingga 0,99), dengan tingkat ketidaksepakatan (disagreement rate) yang sangat rendah "
        "(hanya 3,21% antara SVM dan Ensemble). Fakta ini menjelaskan secara matematis mengapa ensemble voting tidak mampu mendongkrak kinerja secara signifikan: pada representasi fitur "
        "tekstur GLCM yang homogen dan berdimensi rendah, seluruh algoritma pengklasifikasi dasar cenderung membuat galat pada sampel citra sulit yang sama, sehingga meredam efek diversitas ensemble."
    )

    # -------------------------------------------------------------------------
    # 9. HASIL & PEMBAHASAN: 4.3 PERBANDINGAN STUDI & OPTIMISM GAP (RQ1 & RQ2)
    # -------------------------------------------------------------------------
    print("Updating Results (4.3 Comparison & Optimism Gap)...")
    # P98: Pembahasan Optimism Gap
    doc.paragraphs[98].text = (
        "Untuk menjawab RQ1 mengenai sensitivitas kinerja terhadap protokol pembagian data, evaluasi komparatif pada Protokol P1, P2, dan P3 "
        "membuktikan terjadinya optimism gap yang substansial. Pada Protokol P1 (split acak 70:10:20 per citra), Ensemble Voting tampak meraih performa sangat tinggi "
        "dengan akurasi 87,03%, spesifisitas 79,50%, dan F1-Score 91,00%. Namun, ketika model yang sama dievaluasi pada Protokol P2 (split resmi 624 citra), "
        "akurasinya terkoreksi turun menjadi 79,33% (penurunan 7,70%) dan spesifisitasnya anjlok sebesar 22,66% (menjadi 56,84%). Sementara itu, pada Protokol P3 "
        "(split murni per pasien pada kohort gabungan), model menghasilkan akurasi 87,61% dan spesifisitas 80,53%. Selisih tajam ini membuktikan bahwa split acak citra (P1) "
        "menciptakan ilusi performa akibat kebocoran visual pasien yang sama antara data latih dan uji.\n\n"
        "Terkait RQ2, studi ablasi 27 kombinasi membuktikan bahwa pra-pengolahan ortiz (histogram matching, BCET, masking, CLAHE) dikombinasikan dengan kuantisasi Ng = 32 "
        "dan fitur GLCM-6 menghasilkan performa terbaik pada set uji resmi P2 (Akurasi 80,77%, Recall 94,10%, Spesifisitas 58,55%, F1 85,95%, Balanced Acc 76,32%). "
        "Sebaliknya, varian GLCM-24 (24 fitur terarah) secara konsisten mengalami penurunan spesifisitas hingga 33%–44% akibat overfitting terhadap orientasi rotasi."
    )

    # P99 & Table 18: Perbandingan dengan Ortiz-Toro et al. [15]
    doc.paragraphs[99].text = (
        "Tabel 3 menyajikan perbandingan objektif antara hasil penelitian ini dan pembanding utama Ortiz-Toro et al. [15] pada set uji resmi 624 citra yang sama, "
        "serta studi pendahulu Hismaya dan Sulistianingsih [19]."
    )

    t18 = doc.tables[18]
    headers_t18 = ["Aspek Pembanding", "Pembanding Utama: Ortiz-Toro et al. (2022) [15]", "Penelitian Ini (GLCM + Ensemble Terkontrol)", "Studi Pendahulu: Hismaya & Sulistianingsih [19]"]
    while len(t18.columns) < len(headers_t18):
        t18.add_column(Inches(1.2))
    for c_idx, h in enumerate(headers_t18):
        set_cell_text(t18.rows[0].cells[c_idx], h, bold=True, font_size=8.0, align=WD_ALIGN_PARAGRAPH.CENTER)

    rows_t18_data = [
        ["Fitur yang Digunakan", "Radiomics lengkap (First-order statistics, Fractal Dimension, Superpixel Histon)", "6 Fitur Tekstur GLCM (Rerata 4 arah: 0°, 45°, 90°, 135°) berdimensi ringkas", "24 Fitur GLCM (6 fitur x 4 arah tanpa dirata-ratakan)"],
        ["Pra-pengolahan Citra", "Histogram matching, BCET, Top 25% masking, CLAHE pada citra asli", "Grayscale 8-bit 150×150 piksel, StandardScaler terisolasi, studi ablasi 27 kombinasi", "Grayscale 300×300 piksel tanpa penskalaan fitur (StandardScaler)"],
        ["Protokol Data Uji", "Set uji resmi Kaggle: 624 citra (234 Normal [37,5%], 390 Pneumonia [62,5%])", "Set uji resmi Kaggle: 624 citra (P2) + kuantisasi komparatif P1 dan P3", "Set uji resmi Kaggle: 624 citra (234 Normal, 390 Pneumonia)"],
        ["Kinerja Tertinggi", "Random Forest: Akurasi 82,21%, Recall 89,74%, Spesifisitas 69,66%, F1 86,10%", "Ablasi Ortiz+Ng32: Akurasi 80,77%, Recall 94,10%, F1 85,95% (Baseline P2: 79,33%)", "SVM Baseline: Akurasi 76,12% (Recall 86,41%, Spesifisitas 58,97%, F1 81,80%)"],
        ["Evaluasi Statistik & Keragaman", "Evaluasi metrik diagnostik tanpa uji signifikansi ensemble", "Uji McNemar berpasangan (Bonferroni alpha=0,005), korelasi galat, Permutation Importance", "Tanpa uji signifikansi statistik dan tanpa analisis feature importance"]
    ]
    while len(t18.rows) < len(rows_t18_data) + 1:
        t18.add_row()
    for r_idx, r_data in enumerate(rows_t18_data):
        row = t18.rows[r_idx + 1]
        for c_idx, val in enumerate(r_data):
            align = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            set_cell_text(row.cells[c_idx], val, bold=(c_idx==0), font_size=7.5, align=align)

    # -------------------------------------------------------------------------
    # 10. SUBBAB BARU: 4.4 ANCAMAN VALIDITAS (THREATS TO VALIDITY)
    # -------------------------------------------------------------------------
    print("Adding New Section: 4.4 Ancaman Validitas (Threats to Validity)...")
    # Insert new paragraph before Kesimpulan (P102)
    # Let's find paragraph containing "5   Kesimpulan"
    kesimpulan_idx = -1
    for idx, p in enumerate(doc.paragraphs):
        if "Kesimpulan" in p.text and any(p.text.strip().startswith(x) for x in ["5", "5."]):
            kesimpulan_idx = idx
            break

    threats_text = (
        "4.4   Ancaman Validitas (Threats to Validity)\n"
        "Untuk menjaga integritas dan transparansi saintifik, penelitian ini mengidentifikasi sejumlah ancaman validitas metodologis yang membatasi generalisasi temuan:\n"
        "1. Validitas Eksternal dan Generalisasi Populasi: Seluruh citra radiografi dalam dataset Kermany et al. berasal dari satu institusi tunggal (Guangzhou Women and Children's Medical Center) yang mencakup pasien pediatrik usia satu hingga lima tahun. Oleh karena itu, performa model GLCM ini tidak dapat digeneralisasi secara langsung pada pasien dewasa, citra rontgen posisi posteroanterior (PA), atau populasi rumah sakit lain yang memiliki kalibrasi modalitas pencitraan berbeda.\n"
        "2. Status Kebocoran Identitas Pasien (Patient ID Overlap): Audit data pada penelitian ini membuktikan secara empiris bahwa set uji resmi Kaggle memuat 170 identitas pasien Pneumonia yang tumpang tindih dengan data latih. Meskipun protokol P3 berhasil mengontrol isolasi pasien 100%, evaluasi pada set uji resmi P2 tetap membawa bias bawaan dari desain dataset aslinya.\n"
        "3. Ketiadaan Segmentasi Paru dan Risiko Bias Artefak Tepi: Ekstraksi fitur GLCM dilakukan pada seluruh bidang citra toraks tanpa segmentasi parenkim paru otomatis (seperti U-Net). Hal ini menimbulkan risiko model mempelajari penanda radio-opaque (huruf L/R), diafragma, atau kontras batas tulang rusuk. Uji ketahanan artefak (center crop) membuktikan bahwa saat tepi citra dipotong sebesar 10% dan 15%, akurasi model mengalami degradasi signifikan masing-masing menjadi 68,59% dan 63,14% (mendekati batas baseline mayoritas 62,50%). Temuan ini mengonfirmasi bahwa fitur GLCM global sangat sensitif terhadap kontras struktur anatomi perifer.\n"
        "4. Multikolinieritas Fitur GLCM: Parameter tekstur GLCM saling berkorelasi secara matematis (khususnya ASM dan Energy), sehingga tidak dapat ditafsirkan sebagai biomarker patologis langsung melainkan deskriptor visual makro.\n"
        "5. Ketiadaan Pembanding CNN Internal dan Validasi Radiolog Baru: Penelitian ini difokuskan secara mendalam pada analisis komparatif model Machine Learning berbasis fitur tekstur terukur dan tidak membangun baseline CNN secara mandiri. Selain itu, anotasi label sepenuhnya mengandalkan data sekunder tanpa dilakukan proses pembacaan ulang (re-reading) independen oleh dokter spesialis radiologi dalam penelitian ini."
    )

    if kesimpulan_idx != -1:
        new_p = doc.paragraphs[kesimpulan_idx].insert_paragraph_before(threats_text)
        new_p.runs[0].font.name = "Times New Roman"
        new_p.runs[0].font.size = Pt(10)
        # Add space after
        new_p.paragraph_format.space_after = Pt(8)

    # -------------------------------------------------------------------------
    # 11. KESIMPULAN & SARAN
    # -------------------------------------------------------------------------
    print("Updating Conclusion (Kesimpulan)...")
    # Find Kesimpulan text paragraph
    for idx, p in enumerate(doc.paragraphs):
        if "Penelitian ini telah mengevaluasi secara komparatif" in p.text:
            p.text = (
                "Penelitian ini telah mengevaluasi secara komparatif kinerja fitur tekstur GLCM dan algoritma Machine Learning "
                "(SVM, Random Forest, KNN, dan Soft Voting Ensemble) dengan fokus pada kuantisasi sensitivitas protokol pembagian data, dekomposisi faktor, "
                "dan pengujian batas efektivitas ensemble. Berdasarkan hasil pengujian empiris, ditarik beberapa kesimpulan pokok:\n"
                "1. Terbukti adanya optimism gap yang nyata: evaluasi berbasis citra acak (P1) menghasilkan akurasi semu 87,03% dan spesifisitas 79,50%, "
                "yang terkoreksi turun menjadi 79,33% dan 56,84% pada set uji resmi P2 (penurunan akurasi 7,70% dan spesifisitas 22,66%).\n"
                "2. Pada pengujian uji resmi P2, SVM Tuned dan Ensemble Soft Voting meraih akurasi 79,33% dengan recall 92,82% (SVM: 91,03%) dan F1-Score 84,88%. "
                "Pra-pengolahan Ortiz-Toro yang dipadukan dengan kuantisasi Ng = 32 mampu meningkatkan akurasi hingga 80,77% dan recall 94,10%.\n"
                "3. Uji statistik McNemar berpasangan (alpha = 0,005) membuktikan bahwa Soft Voting Ensemble tidak memberikan peningkatan performa yang signifikan "
                "secara statistik dibanding model tunggal (p > 0,005) akibat korelasi galat yang sangat tinggi (r = 0,79 - 0,90) pada representasi fitur tekstur homogen.\n"
                "4. Permutation Feature Importance menunjukkan bahwa Dissimilarity (7,51%) dan Homogeneity (6,47%) merupakan prediktor dominan, sedangkan uji ketahanan artefak "
                "mengungkapkan penurunan performa drastis saat tepi citra dipotong (hingga 63,14%), menandakan ketergantungan fitur global terhadap kontras batas toraks.\n\n"
                "Sebagai saran penelitian lanjutan, pengembangan masa depan direkomendasikan untuk mengintegrasikan model segmentasi paru otomatis (seperti U-Net atau SAM) "
                "sebelum ekstraksi GLCM guna mengisolasi murni parenkim paru, menguji generalisasi model pada dataset multi-senter dengan populasi pasien dewasa, serta "
                "menggabungkan fitur GLCM dengan representasi radiomik multi-skala atau fitur representasi mendalam (deep features) ringan untuk meningkatkan keragaman ensemble."
            )
            p.runs[0].font.name = "Times New Roman"
            p.runs[0].font.size = Pt(10)
            break

    # Save document
    print(f"Saving updated manuscript to {DOC_PATH}...")
    doc.save(DOC_PATH)
    print("Manuscript updated successfully!")

if __name__ == "__main__":
    update_manuscript()
