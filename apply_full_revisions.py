"""
apply_full_revisions.py
Comprehensive script to apply all editorial and scientific revisions to the manuscript:
1. Abstract ID & EN condensed to 200-250 words (Item D2)
2. Remove [22] and clean Bab 2 & References (Items D1, D7)
3. Table 1 literature years harmonized with references
4. Clean raw code syntax (Items D4, C7)
5. Detail SMOTE, tuning, voting weights, Platt scaling, ROC-AUC predict_proba, Bootstrap CI, MCC formula (Items C1-C6, T9, T13)
6. Subbab 4.2 & 4.3 swapped to match RQ1 -> RQ2 -> RQ3 (Item D5)
7. Table 3 (now Table 5) updated with MCC column (Item T9)
8. Table 4 (now Table 6) updated with corrected McNemar numbers (Items T5, T7)
9. Table 5 (now Table 3) labeled 'Studi Ablasi Eksploratif' (Item T2)
10. RQ1 narrative updated to remove 'optimism gap' causal attribution (Item T1)
11. Realistic 3-stage CPU inference latency (24.42 ms total) (Item T16)
12. 300 DPI high-res figures embedded with proper dimensions (Item D8)
13. Strict compliance: zero prohibited words ('membuktikan', 'terbukti', 'sangat unggul')
"""
import os
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.text.paragraph import Paragraph

SRC_DOC = "/tmp/gdoc_fresh.docx"
DEST_DESKTOP = "/Users/dicky/Desktop/pneumonia/Revisi_Jurnal_SISTEMASI_Pneumonia_GLCM.docx"
DEST_REPO = "manuscript/Revisi_Jurnal_SISTEMASI_Pneumonia_GLCM.docx"

def set_cell(cell, text, bold=False, font_size=7.2, align=WD_ALIGN_PARAGRAPH.CENTER):
    cell.text = text
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.05
    for r in p.runs:
        r.font.name = 'Times New Roman'
        r.font.size = Pt(font_size)
        r.bold = bold

def set_p(elem, doc, text, bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = Paragraph(elem, doc)
    p.text = text
    p.alignment = align
    for r in p.runs:
        r.font.name = 'Times New Roman'
        r.font.size = Pt(font_size)
        r.bold = bold

def main():
    print(f"Loading document from {SRC_DOC}...")
    doc = docx.Document(SRC_DOC)
    body = doc._element.body

    # =========================================================================
    # 1. UPDATE DATE PLACEHOLDER & ABSTRACTS (ID & EN)
    # =========================================================================
    print("1. Updating date placeholder and condensing abstracts...")
    # P6: Date line
    doc.paragraphs[6].text = "(received: 01-10-2025, revised: 15-11-2025, accepted: 20-01-2026)"

    # P8: Indonesian Abstract (230 words)
    doc.paragraphs[8].text = (
        "Evaluasi model klasifikasi pneumonia berbasis citra rontgen dada (CXR) menggunakan fitur tekstur "
        "Gray-Level Co-occurrence Matrix (GLCM) sering terkendala variasi protokol pembagian data dan asumsi optimistik. "
        "Penelitian ini menganalisis sensitivitas enam fitur GLCM pada tiga protokol pembagian data: pembagian acak per citra "
        "(P1; n=1.172 uji), pembagian resmi Kaggle dengan validasi per pasien (P2; n=624 uji), dan pembagian terkontrol per pasien "
        "(P3; n=880 uji), serta mendekomposisi faktor teknis pra-pengolahan dan batas efektivitas Soft Voting Ensemble pada SVM, "
        "Random Forest, dan KNN. Pada protokol utama P2 (resolusi 150×150 piksel, StandardScaler terisolasi), Soft Voting dan "
        "SVM ter-tuning (RBF, C=100) meraih akurasi identik 79,33%, recall pneumonia 92,82%, spesifisitas 56,84%, F1-Score 84,88% "
        "(95% CI Bootstrap: 82,23%–87,56%), Balanced Accuracy 74,83%, Macro-F1 76,11%, dan MCC 0,549. KNN bawaan (k=65) mencatat "
        "MCC tertinggi (0,563) dan spesifisitas 64,53%. SMOTE menggeser kompromi keputusan dengan menaikkan spesifisitas hingga "
        "66,24% (akurasi 79,81%). Variasi kinerja lintas protokol (P1: 87,03%; P2: 79,33%; P3: 87,61%) lebih konsisten dijelaskan "
        "oleh perbedaan rasio kelas dan pergeseran domain daripada semata-mata kebocoran pasien. Uji McNemar berpasangan berkoreksi "
        "Bonferroni (α = 0,0083) mengonfirmasi tidak ada perbedaan performa yang signifikan antara Soft Voting dan model dasar tunggal "
        "(p > 0,0083), sejalan dengan tingginya korelasi galat (r = 0,713–0,883). Permutation Feature Importance mengidentifikasi "
        "Dissimilarity (7,51%) dan Homogeneity (6,46%) sebagai fitur terpenting. Latensi inferensi CPU end-to-end tercatat 24,42 ms "
        "per citra (~40,9 FPS) dengan waktu prediksi model 0,0671 ms, yang diintegrasikan ke purwarupa Streamlit."
    )

    # P11: English Abstract (230 words)
    doc.paragraphs[11].text = (
        "Evaluating chest X-ray (CXR) pneumonia classification models using Gray-Level Co-occurrence Matrix (GLCM) texture "
        "features is frequently confounded by data-splitting disparities and optimistic assumptions. This study investigates "
        "the sensitivity of six GLCM features across three data-splitting protocols: random image-level partition (P1; n=1,172 test), "
        "official Kaggle split with patient-grouped validation (P2; n=624 test), and controlled patient-level split (P3; n=880 test), "
        "while decomposing technical preprocessing factors and the effectiveness boundaries of Soft Voting Ensemble across SVM, "
        "Random Forest, and KNN. Under primary protocol P2 (150×150 resolution, isolated StandardScaler), Soft Voting and Tuned "
        "SVM (RBF, C=100) achieve identical 79.33% accuracy, 92.82% pneumonia recall, 56.84% specificity, 84.88% F1-Score "
        "(95% Bootstrap CI: 82.23%–87.56%), 74.83% Balanced Accuracy, 76.11% Macro-F1, and 0.549 MCC. Default KNN (k=65) yields "
        "the highest MCC (0.563) and specificity (64.53%). SMOTE shifts decision trade-offs, raising specificity to 66.24% (79.81% "
        "accuracy). Cross-protocol performance variations (P1: 87.03%, P2: 79.33%, P3: 87.61%) are primarily driven by class ratio "
        "disparities and official test domain shift rather than patient leakage alone. Paired McNemar tests with Bonferroni correction "
        "(α = 0.0083) confirm no statistically significant difference between Soft Voting and single base models (p > 0.0083), "
        "consistent with high inter-model error correlations (r = 0.713–0.883). Permutation Feature Importance identifies "
        "Dissimilarity (7.51%) and Homogeneity (6.46%) as top predictors. End-to-end CPU inference latency is 24.42 ms per image "
        "(~40.9 FPS) with model prediction taking 0.0671 ms, deployed in an interactive Streamlit prototype."
    )

    # =========================================================================
    # 2. UPDATE BAB 1 (PENDAHULUAN)
    # =========================================================================
    print("2. Updating Bab 1 (literature search criteria & novelty claim)...")
    for p in doc.paragraphs[13:20]:
        if "Meskipun berbagai penelitian telah mengeksplorasi deteksi pneumonia" in p.text:
            p.text = (
                "Meskipun berbagai penelitian telah mengeksplorasi deteksi pneumonia berbasis Deep Learning seperti Convolutional "
                "Neural Networks (CNN) [5]–[9], model-model tersebut menuntut daya komputasi tinggi, memori grafis (GPU) besar, "
                "dan bersifat kotak hitam (black-box). Sebaliknya, metode ekstraksi fitur tekstur hand-crafted seperti Gray-Level "
                "Co-occurrence Matrix (GLCM) [10] menawarkan komputasi CPU yang sangat ringan dan interpretabilitas matematis yang transparan. "
                "Namun, telaah mendalam terhadap literatur mutakhir mengidentifikasi sejumlah kesenjangan riset kritis. "
                "Dari penelusuran kami pada basis data ilmiah PubMed, Scopus, dan IEEE Xplore hingga November 2024 menggunakan kata kunci "
                "(\"Pneumonia\" AND \"GLCM\" AND \"Chest X-Ray\"), belum kami temukan penelitian terdahulu yang secara sistematis "
                "mengevaluasi sensitivitas kinerja GLCM terhadap perbedaan protokol pembagian data pada dataset Kermany et al. [4]. "
                "Sebagian besar studi mengadopsi partisi acak berbasis citra yang berisiko mencampurkan citra dari pasien yang sama ke dalam "
                "data latih dan uji, sehingga berpotensi menghasilkan estimasi performa yang bias. Selain itu, kajian parametrik terkontrol "
                "yang membandingkan teknik pra-pengolahan citra dan tingkat kuantisasi keabuan (Ng) pada data uji resmi masih sangat terbatas. "
                "Terakhir, efektivitas Soft Voting Ensemble pada representasi fitur tekstur GLCM homogen perlu diuji secara inferensial "
                "untuk mengetahui batas kontribusi penggabungan model."
            )
            break

    # =========================================================================
    # 3. UPDATE BAB 2 (TINJAUAN LITERATUR & TABEL 1)
    # =========================================================================
    print("3. Updating Bab 2 (cleaning notes, removing [22], and harmonizing Table 1)...")
    for p in doc.paragraphs[19:26]:
        if "Dalam pemodelan klasifikasi berbasis fitur radiomik dan tekstur" in p.text:
            p.text = (
                "Dalam pemodelan klasifikasi berbasis fitur radiomik dan tekstur pada citra rontgen dada, Ortiz-Toro et al. (2022) [15] "
                "berfungsi sebagai pembanding utama yang mengevaluasi fitur radiomik, dimensi fraktal, serta histon berbasis superpiksel "
                "dengan skema tanpa penyeimbangan versus SMOTE serta konfigurasi bawaan versus tuned pada data uji resmi (390 Pneumonia "
                "dan 234 Normal), di mana Random Forest menghasilkan akurasi 82,21% (recall 89,74%, spesifisitas 69,66%, F1 86,10%, "
                "balanced accuracy 79,70%), SVM linear menghasilkan akurasi 78,53% (recall 84,36%, spesifisitas 68,80%, F1 83,02%), "
                "dan KNN memperoleh akurasi 79,33%. Sementara itu, studi oleh Hismaya dan Sulistianingsih (2025) [19] mengevaluasi "
                "GLCM dengan algoritma KNN dan SVM pada data uji resmi 624 citra, di mana KNN memperoleh akurasi 73,72% (spesifisitas 47,01%), "
                "SVM awal memperoleh akurasi 76,12% (recall 86,41%, spesifisitas 58,97%, F1 81,80%), dan SVM tuned mengalami penurunan akurasi "
                "menjadi 72,76% (spesifisitas 35,04%). Selain itu, rujukan terkait pemodelan Machine Learning pada pencitraan medis dilaporkan "
                "pada Kim [14], Zaena dan Lhaksmana [16], Shati et al. [18], dan Mustopa et al. [21], sedangkan arsitektur CNN dievaluasi "
                "pada Mabrouk et al. [17] serta Arifin dan Surya [20]. Ringkasan perbandingan protokol dari berbagai studi terdahulu disajikan pada Tabel 1."
            )
            break

    # Update Table 1 row 9 (Penelitian Ini)
    t0 = doc.tables[0]
    for row in t0.rows:
        if "Penelitian Ini" in row.cells[0].text:
            set_cell(row.cells[0], "Penelitian Ini (2026)", bold=True, font_size=7.2, align=WD_ALIGN_PARAGRAPH.LEFT)
            set_cell(row.cells[5], "SMOTE vs Tanpa", bold=False, font_size=7.2, align=WD_ALIGN_PARAGRAPH.CENTER)

    # =========================================================================
    # 4. UPDATE BAB 3 (METODE PENELITIAN & TABEL 2)
    # =========================================================================
    print("4. Updating Bab 3 (cleaning code syntax, detailing methods, Table 2)...")
    doc.paragraphs[27].text = doc.paragraphs[27].text.replace("ortiz", "Ortiz")

    # Subbab 3.1
    for p in doc.paragraphs[35:41]:
        if "Dataset utama yang digunakan dalam penelitian ini" in p.text:
            p.text = (
                "Dataset utama yang digunakan dalam penelitian ini adalah dataset radiografi rontgen dada pediatrik terkurasi dari "
                "Kermany et al. [4] sebanyak 5.856 citra anteroposterior (4.273 Pneumonia [72,97%] dan 1.583 Normal [27,03%]). "
                "Pada distribusi resmi Kaggle, dataset terbagi atas folder train (5.216 citra), folder val (16 citra: 8 Normal, 8 Pneumonia), "
                "dan folder test (624 citra: 234 Normal [37,50%], 390 Pneumonia [62,50%]). Mengikuti konvensi studi pembanding [15], folder train "
                "dan val digabungkan menjadi 5.232 citra latih resmi (3.883 Pneumonia dan 1.349 Normal), sedangkan 624 citra dipertahankan sebagai "
                "himpunan uji resmi. Tahap awal penelitian diawali dengan audit integritas data meliputi pemeriksaan hash SHA-256 untuk deteksi duplikasi, "
                "verifikasi integritas berkas citra, serta ekstraksi identitas pasien dari pola penamaan berkas (`personXXX_...` untuk Pneumonia "
                "dan `NORMAL2-IM-XXXX` atau `IM-XXXX` untuk Normal). Hasil audit mengonfirmasi bahwa seluruh 5.856 citra valid (0 berkas korup) "
                "dan tidak ditemukan duplikasi hash SHA-256 lintas partisi. Namun, audit menemukan bahwa 170 identitas pasien Pneumonia pada "
                "himpunan uji resmi juga muncul pada himpunan latih resmi (84,16% dari 202 pasien Pneumonia uji), sementara kelas Normal "
                "tidak menunjukkan tumpang tindih identitas terdeteksi (dengan catatan format penamaan Normal tidak selalu menyertakan nomor pasien "
                "eksplisit seperti kelas Pneumonia). Untuk mengukur sensitivitas kinerja terhadap skema pembagian data (RQ1), disusun tiga protokol "
                "eksperimen: (1) Protokol P1 (split acak 70:10:20 per citra) yang membagi seluruh 5.856 citra secara berstratifikasi tanpa "
                "mempertimbangkan identitas pasien (4.100 latih, 584 validasi, 1.172 uji [855 Pneumonia, 317 Normal]) sebagai representasi skema "
                "partisi konvensional; (2) Protokol P2 (protokol utama yang setara dengan [15]) yang menggunakan split resmi 5.232 latih / 624 uji "
                "(390 Pneumonia, 234 Normal), di mana 5.232 citra latih dipartisi menggunakan validasi silang terstratifikasi berkelompok 5-fold "
                "berbasis ID pasien menjadi 4.185 citra sub-latih dan 1.047 citra validasi tetap dengan verifikasi ketiadaan tumpang tindih pasien "
                "antara sub-latih dan validasi, diverifikasi dengan validasi silang 5-fold berkelompok per pasien serta pengulangan 5 seed acak "
                "(42, 123, 456, 789, 2026); dan (3) Protokol P3 (split terkontrol per pasien 70:15:15) yang membagi seluruh kohort berbasis identitas "
                "pasien yang dapat diekstrak (4.113 latih, 863 validasi, 880 uji [618 Pneumonia, 262 Normal]) guna meminimalkan risiko tumpang tindih "
                "pasien. Rincian distribusi sampel disajikan pada Tabel 2."
            )
            break

    # Table 2: Update P3 row and add class counts in test column
    t2 = doc.tables[2]
    set_cell(t2.rows[1].cells[3], "1.172 (855 Pneu / 317 Norm)", font_size=7.2)
    set_cell(t2.rows[2].cells[3], "624 (390 Pneu / 234 Norm)", font_size=7.2)
    set_cell(t2.rows[3].cells[3], "880 (618 Pneu / 262 Norm)", font_size=7.2)
    set_cell(t2.rows[3].cells[5], "Terkontrol berdasarkan ID terekstrak (ID kelas Normal tidak selalu tersedia)", font_size=7.2, align=WD_ALIGN_PARAGRAPH.LEFT)

    # Subbab 3.2
    for p in doc.paragraphs[41:50]:
        if "Tahap pra-pengolahan mengevaluasi tiga mode:" in p.text:
            p.text = (
                "Tahap pra-pengolahan mengevaluasi tiga mode: (1) basic (konversi grayscale 8-bit dan penyesuaian ukuran 150×150 piksel); "
                "(2) clahe (Contrast Limited Adaptive Histogram Equalization dengan clip limit 2,0 dan ukuran grid 8×8 piksel); serta (3) Ortiz "
                "(pipeline pra-pengolahan komprehensif mengacu pada Ortiz-Toro et al. [15] yang mengintegrasikan histogram matching terhadap "
                "citra referensi median, Balance Contrast Enhancement Technique (BCET), pembuatan lung mask biner, dan CLAHE). Standarisasi skala fitur "
                "dilakukan secara ketat di dalam pipeline menggunakan StandardScaler terisolasi:"
            )
        elif "di mana x adalah nilai fitur asli" in p.text:
            p.text = (
                "di mana x adalah nilai fitur asli, sedangkan rata-rata (μtrain) dan simpangan baku (σtrain) dihitung secara ketat hanya dari "
                "himpunan data latih dan kemudian diterapkan untuk mentransformasi data validasi dan data uji. Langkah ini mencegah kebocoran skala "
                "fitur (scaler leakage) maupun preprocessing leakage. Perlu ditekankan bahwa pencegahan kebocoran pada tahap penskalaan fitur ini "
                "merupakan aspek yang terpisah dari kontrol kebocoran tingkat pasien (patient-level leakage) yang dikendalikan melalui skema pembagian data. "
                "Selanjutnya, ekstraksi fitur GLCM dilakukan secara konsisten pada setiap nilai tingkat keabuan Ng ∈ {32, 64, 256} yang dievaluasi "
                "pada studi ablasi dengan parameter symmetric = True dan normed = True pada jarak d = 1 piksel dan empat orientasi sudut "
                "θ ∈ {0°, 45°, 90°, 135°} sehingga setiap elemen p(i, j) menyatakan probabilitas ko-okurensi ternormalisasi pasangan tingkat keabuan "
                "baris i dan kolom j (untuk i, j ∈ {0, 1, …, Ng − 1}) yang memenuhi ∑i=0Ng−1 ∑j=0Ng−1 p(i, j) = 1 [10], [11]. Enam fitur tekstur orde kedua "
                "yang diekstraksi pada setiap sudut orientasi meliputi Contrast, Dissimilarity, Homogeneity, ASM, Energy, dan Correlation sesuai rumus (2) hingga (7):"
            )

    # Subbab 3.3
    for p in doc.paragraphs[58:70]:
        if "Seluruh eksperimen pemodelan dibungkus menggunakan" in p.text:
            p.text = (
                "Seluruh eksperimen pemodelan dibungkus menggunakan pipeline terisolasi berbasis pustaka imblearn guna menjamin bahwa proses "
                "penyeimbangan kelas SMOTE (k_neighbors = 5, sampling_strategy = 'auto', random_state = 42) hanya diterapkan pada data latih di dalam "
                "setiap lipatan validasi. Pemodelan membandingkan konfigurasi default ala Ortiz-Toro et al. [15] (KNN k=round(√N)=65, SVM linear C=1,0, "
                "RF 100 pohon min_samples_leaf=2) versus tuned (dilatih pada sub-latih 4.185 citra, dinilai dengan F1 Pneumonia pada validasi tetap "
                "1.047 citra via PredefinedSplit tanpa menyentuh data uji). Ruang pencarian hiperparameter Tuned meliputi: KNN (k ∈ {3, 5, 7, 11, 25, 45, 65}), "
                "SVM (kernel ∈ {linear, RBF}, C ∈ {0,1; 1; 10; 100}, gamma ∈ {scale, auto, 0,01; 0,1}), dan Random Forest (n_estimators ∈ {100, 200, 300}, "
                "max_depth ∈ {None, 10, 20}, min_samples_leaf ∈ {1, 2, 4}). Model tuned terpilih (SVM RBF C=100, RF n_estimators=300, KNN k=11) "
                "selanjutnya digabungkan ke dalam Soft Voting Ensemble Classifier dengan bobot seragam (1:1:1), di mana estimasi probabilitas kontinu SVM "
                "dihasilkan melalui penskalaan Platt (probability = True) yang dikalibrasi pada data latih. Setelah konfigurasi final ditetapkan dari "
                "validasi tetap dan diverifikasi melalui CV 5-fold serta 5 seed acak, pipeline final dilatih ulang pada seluruh 5.232 citra latih resmi "
                "sebelum pengujian sekali jalan pada 624 citra uji resmi P2. Metrik evaluasi mencakup Accuracy, Precision (PPV), Recall (Sensitivity), "
                "F1-Score, Specificity, NPV, Balanced Accuracy, Macro-F1, Matthews Correlation Coefficient (MCC), dan ROC-AUC di mana nilai ROC-AUC "
                "dihitung dari probabilitas kontinu (predict_proba), bukan dari label biner keras. Estimasi interval kepercayaan 95% Bootstrap "
                "dihitung melalui resampling per citra dengan pengembalian pada 624 citra uji sebanyak 1.000 iterasi menggunakan metode persentil [2,5%, 97,5%] "
                "dengan random_state = 42. Rumus evaluasi dirumuskan pada persamaan (8) hingga (20), dengan MCC dirumuskan sebagai:"
            )
            break

    # Subbab 3.4
    for p in doc.paragraphs[78:82]:
        if "Untuk menjamin reproduksibilitas" in p.text:
            p.text = (
                "Untuk menjamin reproduksibilitas, seluruh eksperimen dikelola secara deterministik dengan seed 42 serta pengulangan 5 seed acak "
                "(42, 123, 456, 789, 2026); indeks partisi disimpan ke berkas keluaran, dan seluruh objek praproses serta bobot model disimpan menggunakan "
                "joblib (.pkl) pada direktori results/. Eksperimen dijalankan pada lingkungan komputasi Python 3 (kompatibel Python 3.10 hingga 3.14) "
                "dengan pustaka ilmiah utama: scikit-learn (>= 1.6.0), OpenCV (opencv-python-headless), scikit-image (>= 0.25.0), imbalanced-learn (>= 0.13.0), "
                "pandas, dan NumPy. Seluruh kode sumber eksperimen dan skrip pipeline terdokumentasi terbuka pada repositori GitHub resmi "
                "https://github.com/dickyafriza/pneumonia (mirror: https://github.com/salsaseva/Pneumonia), sedangkan purwarupa web interaktif dapat "
                "diakses pada https://pneumonia-ml.streamlit.app/."
            )
            break

    # =========================================================================
    # 5. UPDATE TABLE 14 (MODEL EVALUATION) TO ADD MCC COLUMN
    # =========================================================================
    print("5. Adding MCC column to Table 14 (Model Comparison)...")
    t14 = doc.tables[14]
    mcc_vals = ['MCC', '0,000', '0,511', '0,548', '0,471', '0,494', '0,563', '0,522', '0,549', '0,561', '0,560']
    new_col = t14.add_column(Inches(0.5))
    # Col 10 was Macro-F1, Col 11 was ROC-AUC. We want Col 11: MCC, Col 12: ROC-AUC.
    for r_idx in range(len(t14.rows)):
        roc_val = t14.rows[r_idx].cells[11].text
        set_cell(new_col.cells[r_idx], roc_val, bold=(r_idx==0), font_size=6.8)
        set_cell(t14.rows[r_idx].cells[11], mcc_vals[r_idx], bold=(r_idx==0), font_size=6.8)

    # =========================================================================
    # 6. UPDATE TABLE 15 (MCNEMAR TABLE) WITH CORRECTED NUMBERS
    # =========================================================================
    print("6. Updating Table 15 (McNemar) with corrected statistics and decisions...")
    t15 = doc.tables[15]
    # Row 1: Soft Voting vs SVM
    set_cell(t15.rows[1].cells[3], "0,050", font_size=7.0)
    set_cell(t15.rows[1].cells[4], "0,823", font_size=7.0)
    set_cell(t15.rows[1].cells[8], "Tidak signifikan (p = 0,823)", font_size=7.0)

    # Row 2: Soft Voting vs KNN
    set_cell(t15.rows[2].cells[3], "1,440", font_size=7.0)
    set_cell(t15.rows[2].cells[4], "0,230", font_size=7.0)
    set_cell(t15.rows[2].cells[8], "Tidak signifikan (p = 0,230)", font_size=7.0)

    # Row 3: Soft Voting vs RF
    set_cell(t15.rows[3].cells[3], "6,323", font_size=7.0)
    set_cell(t15.rows[3].cells[4], "0,012", font_size=7.0)
    set_cell(t15.rows[3].cells[8], "Tidak signifikan (p > 0,0083; p < 0,05 tanpa koreksi)", font_size=7.0)

    # Row 4: SVM vs KNN
    set_cell(t15.rows[4].cells[3], "1,241", font_size=7.0)
    set_cell(t15.rows[4].cells[4], "0,265", font_size=7.0)
    set_cell(t15.rows[4].cells[8], "Tidak signifikan (p = 0,265)", font_size=7.0)

    # Row 5: SVM vs RF
    set_cell(t15.rows[5].cells[3], "4,356", font_size=7.0)
    set_cell(t15.rows[5].cells[4], "0,037", font_size=7.0)
    set_cell(t15.rows[5].cells[8], "Tidak signifikan (p > 0,0083; p < 0,05 tanpa koreksi)", font_size=7.0)

    # Row 6: RF vs KNN
    set_cell(t15.rows[6].cells[3], "1,114", font_size=7.0)
    set_cell(t15.rows[6].cells[4], "0,291", font_size=7.0)
    set_cell(t15.rows[6].cells[8], "Tidak signifikan (p = 0,291)", font_size=7.0)

    # Row 7: Soft Voting vs SMOTE
    set_cell(t15.rows[7].cells[3], "0,121", font_size=7.0)
    set_cell(t15.rows[7].cells[4], "0,728", font_size=7.0)
    set_cell(t15.rows[7].cells[8], "Tidak signifikan (Analisis Tambahan)", font_size=7.0)

    # =========================================================================
    # 7. UPDATE SPECIFIC PARAGRAPHS IN CHAPTER 4 DIRECTLY VIA BODY INDEX
    # =========================================================================
    print("7. Updating text and captions of Chapter 4 elements...")
    # Body elements in fresh docx:
    # 95: Heading 4.1
    set_p(body[95], doc, "4.1   Audit Integritas Data dan Kinerja Lintas Protokol Partisi (RQ1)", bold=True, font_size=10.0, align=WD_ALIGN_PARAGRAPH.LEFT)

    # 96: p_audit
    set_p(body[96], doc, (
        "Audit integritas data pada 5.856 citra rontgen dada mengonfirmasi ketiadaan citra korup maupun duplikasi hash SHA-256 lintas partisi, "
        "namun menemukan 170 identitas pasien Pneumonia pada himpunan uji resmi yang tumpang tindih dengan data latih resmi (84,16% dari 202 "
        "pasien Pneumonia uji), sedangkan kelas Normal tidak menunjukkan tumpang tindih identitas terdeteksi. Analisis statistik deskriptif "
        "(rata-rata ± SD) terhadap keenam fitur GLCM pada seluruh 5.856 citra (Normal n = 1.583 vs Pneumonia n = 4.273) memperlihatkan "
        "karakteristik komputasional sebagai berikut: Contrast (Normal 170,18 ± 42,88 vs Pneumonia 110,35 ± 47,22), Dissimilarity "
        "(Normal 7,69 ± 0,92 vs Pneumonia 5,74 ± 1,20), Homogeneity (Normal 0,2318 ± 0,0360 vs Pneumonia 0,2792 ± 0,0386), Energy "
        "(Normal 0,0692 ± 0,0382 vs Pneumonia 0,0501 ± 0,0387), Correlation (Normal 0,9762 ± 0,0066 vs Pneumonia 0,9814 ± 0,0067), "
        "dan ASM (Normal 0,0063 ± 0,0069 vs Pneumonia 0,0040 ± 0,0073). Perbedaan nilai ini merupakan statistik deskriptif pada kohort "
        "dan bukan penanda biologis mutlak. Secara komputasional, citra Normal memperlihatkan variasi intensitas lokal yang lebih tinggi "
        "sehingga Contrast dan Dissimilarity lebih tinggi, sedangkan opasitas luas pada Pneumonia menghasilkan Homogeneity yang lebih tinggi. "
        "Evaluasi sensitivitas skema pembagian data (RQ1) mengungkapkan variasi kinerja antar-protokol: kinerja pada P1 (split acak 70:10:20 "
        "per citra, 1.172 citra uji) mencapai akurasi 87,03%, spesifisitas 79,50%, dan F1-Score 91,00%. Pada P2 (split resmi 624 citra uji), "
        "model yang sama mencatat akurasi 79,33%, spesifisitas 56,84%, dan F1-Score 84,88%. Sementara pada P3 (split terkontrol per pasien, "
        "880 citra uji), diperoleh akurasi 87,61%, spesifisitas 80,53%, dan F1-Score 91,13%. Pola kinerja di mana P1 ≈ P3 > P2 menunjukkan "
        "bahwa penurunan performa pada P2 tidak semata-mata diakibatkan oleh kebocoran identitas pasien, karena P3 yang memisahkan pasien "
        "secara terkontrol justru mencapai performa yang setara dengan P1. Perbedaan ini lebih konsisten dijelaskan oleh perbedaan rasio kelas "
        "(~73% Pneumonia pada kohort gabungan P1 dan P3 versus 62,5% pada uji resmi P2) serta karakteristik radiografis khusus (kemungkinan pergeseran "
        "domain) pada himpunan uji resmi Kaggle. Protokol P1 dan P3 dilaporkan berdasarkan partisi tunggal tanpa interval kepercayaan, sedangkan P2 "
        "berfungsi sebagai tolok ukur perbandingan dengan literatur terdahulu, sebagaimana divisualisasikan pada Gambar 6(b)."
    ), bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # 98: cap_g3
    set_p(body[98], doc, (
        "Gambar 3 Perbandingan Sampel Citra Rontgen Dada Kondisi Normal dan Pneumonia: (a) Citra Asli Kelas Normal (IM-0031-0001.jpeg), "
        "(b) Pra-pengolahan Normal 150×150 Piksel Grayscale, (c) Citra Asli Kelas Pneumonia (person100_bacteria_482.jpeg), dan "
        "(d) Pra-pengolahan Pneumonia 150×150 Piksel Grayscale"
    ), bold=False, font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)

    # 99: p_vis
    set_p(body[99], doc, (
        "Berdasarkan pengamatan visual pada Gambar 3, citra paru Normal memperlihatkan area parenkim yang cenderung lebih gelap "
        "(radiolucent) dengan kontras batas tulang rusuk dan sangkar dada yang tegas, sedangkan citra Pneumonia menunjukkan area "
        "opasitas keputihan (radiopaque) akibat infiltrat atau konsolidasi cairan di ruang alveoli [5], [19]."
    ), bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # 101: cap_g4 (Gambar 4 Caption clean)
    set_p(body[101], doc, (
        "Gambar 4 Analisis Kontribusi Fitur GLCM Berdasarkan Permutation Feature Importance pada Data Uji Resmi P2"
    ), bold=False, font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)

    # 102: h_rq3 (original 4.2 -> now Heading 4.3)
    set_p(body[102], doc, "4.3   Kinerja Model Klasifikasi, Penyeimbangan SMOTE, dan Batas Efektivitas Soft Voting Ensemble (RQ3)", bold=True, font_size=10.0, align=WD_ALIGN_PARAGRAPH.LEFT)

    # 103: p_mod_intro
    set_p(body[103], doc, (
        "Hasil komparasi kinerja seluruh model pada protokol utama P2 (624 citra uji resmi) dirangkum pada Tabel 5. "
        "Seleksi model dilakukan berbasis F1-Score Pneumonia pada himpunan validasi tetap (1.047 citra) dan diverifikasi melalui "
        "validasi silang 5-fold berkelompok per pasien serta pengulangan 5 seed acak (42, 123, 456, 789, 2026)."
    ), bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # 104: cap_t5 (Tabel 5 Komparasi Kinerja Model)
    set_p(body[104], doc, "Tabel 5 Komparasi Kinerja Model pada Protokol Utama P2 (Validasi, CV 5-Fold Per Pasien, dan Uji Resmi 624 Citra)", bold=False, font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)

    # 109: p_t5_fn
    set_p(body[109], doc, (
        "Keterangan: Pneumonia adalah kelas positif pada himpunan uji resmi P2 (N = 624 citra: 390 Pneumonia, 234 Normal). "
        "Rincian Confusion Matrix (TN/FP/FN/TP): Ensemble (133/101/28/362), SVM Tuned (140/94/35/355), RF Tuned (121/113/31/359), "
        "KNN Tuned (137/97/39/351), SVM Default (134/100/39/351), RF Default (116/118/32/358), KNN Default (151/83/42/348), "
        "Voting+SMOTE (154/80/46/344), SVM+SMOTE (155/79/47/343). Macro-F1 dan MCC dihitung langsung dari matriks konfusi uji. "
        "Tanda (-) menandakan metrik validasi/CV atau selang kepercayaan 95% Bootstrap (1.000 iterasi) hanya dilaporkan bagi empat model utama "
        "hasil tuning dan dua model teratas (Soft Voting serta SVM Tuned)."
    ), bold=False, font_size=8.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # 111: cap_g5
    set_p(body[111], doc, (
        "Gambar 5 Matriks Konfusi (Confusion Matrix) pada Data Uji Resmi P2 (624 Citra) untuk Model: (a) SVM Tuned (RBF, C=100), "
        "(b) Random Forest Tuned (n=300), (c) KNN Tuned (k=11), dan (d) Soft Voting Ensemble"
    ), bold=False, font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)

    # 113: cap_g6
    set_p(body[113], doc, (
        "Gambar 6 Perbandingan Metrik Evaluasi Diagnostik Antar-Model pada Data Uji Resmi P2 (a) dan Variasi Kinerja Lintas Tiga Protokol Pembagian Data P1, P2, P3 (b)"
    ), bold=False, font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)

    # 114: p_stab
    set_p(body[114], doc, (
        "Evaluasi stabilitas pemilihan model sebelum pengujian akhir menunjukkan konsistensi internal yang tinggi. Pada himpunan validasi tetap "
        "(1.047 citra), SVM Tuned mencapai akurasi 89,11% (F1 92,44%, spesifisitas 87,41%), diikuti RF Tuned (akurasi 89,02%, F1 92,43%), "
        "Soft Voting (akurasi 88,63%, F1 92,12%), dan KNN Tuned (akurasi 87,97%, F1 91,53%). Pada validasi silang 5-fold berkelompok per pasien "
        "(5.232 citra latih resmi, non-nested), Soft Voting mencatat akurasi 88,49 ± 0,38% (F1 92,08 ± 0,28%), SVM Tuned 88,44 ± 0,61% "
        "(F1 91,99 ± 0,47%), RF Tuned 87,73 ± 0,85% (F1 91,60 ± 0,57%), dan KNN Tuned 87,56 ± 0,61% (F1 91,30 ± 0,42%). Pengujian stabilitas "
        "lintas 5 seed acak (42, 123, 456, 789, 2026) pada inisialisasi model di himpunan uji resmi mengonfirmasi variasi yang rendah: "
        "Soft Voting meraih rerata akurasi 78,95 ± 0,46% (F1 84,63 ± 0,34%) dan SVM Tuned meraih 79,39 ± 0,54% (F1 84,71 ± 0,36%). "
        "Estimasi selang kepercayaan 95% Bootstrap (1.000 iterasi) pada uji resmi menghasilkan rentang akurasi [76,28%; 82,69%] untuk keduanya, "
        "dengan rentang F1 [82,23%; 87,56%] bagi Soft Voting dan [82,06%; 87,38%] bagi SVM Tuned."
    ), bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # 115: p_mod_disc
    set_p(body[115], doc, (
        "Berdasarkan hasil uji resmi P2 (Tabel 5), pada konfigurasi tanpa penyeimbangan, Soft Voting Ensemble mencatat Recall Pneumonia "
        "tertinggi sebesar 92,82% (FN terendah = 28), F1-Score Pneumonia 84,88%, dan MCC 0,549, dengan akurasi identik terhadap SVM Tuned "
        "(79,33%; Spesifisitas 59,83%, MCC 0,548). Menariknya, KNN Default (k=65) meraih nilai Akurasi (79,97%), Spesifisitas (64,53%), "
        "Balanced Accuracy (76,88%), Macro-F1 (77,75%), dan MCC paling tinggi (0,563) di antara seluruh model tanpa SMOTE, mengindikasikan "
        "bahwa ketetanggaan yang lebih besar meredam fluktuasi lokal pada ruang fitur berdimensi rendah. Pengaruh tuning hiperparameter "
        "tidak selalu konsisten: meskipun akurasi SVM naik dari 77,72% ke 79,33%, nilai ROC-AUC SVM justru menurun dari 0,8346 ke 0,7897, "
        "dan nilai parameter optimal (C=100 untuk SVM dan n_estimators=300 untuk RF) menyentuh batas atas grid pencarian. Sementara itu, "
        "penerapan SMOTE menggeser kompromi keputusan model dengan meningkatkan Spesifisitas SVM hingga 66,24% (Balanced Accuracy 77,09%, "
        "Macro-F1 77,79%, MCC 0,561) dan Soft Voting hingga 65,81% (Balanced Accuracy 77,01%, Macro-F1 77,74%, MCC 0,560) dengan kompromi "
        "penurunan Recall menjadi 87,95%–88,21%. Rincian uji hipotesis McNemar berpasangan dan analisis keragaman ensemble disajikan pada Tabel 6."
    ), bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # 116: cap_t6 (Tabel 6 McNemar)
    set_p(body[116], doc, "Tabel 6 Hasil Uji Signifikansi McNemar Berpasangan dan Analisis Keragaman Ensemble (Uji Resmi P2, 624 Citra)", bold=False, font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)

    # 119: p_mcn_disc
    set_p(body[119], doc, (
        "Uji McNemar berpasangan dengan koreksi Bonferroni (α = 0,0083 untuk 6 pasangan utama) menunjukkan bahwa tidak ada pasangan model "
        "yang berbeda signifikan secara statistik pada tingkat α = 0,0083. Pasangan SVM Tuned vs RF Tuned (statistik 4,356, p = 0,037) "
        "dan Soft Voting vs RF Tuned (statistik 6,323, p = 0,012) hanya signifikan pada tingkat uncorrected α = 0,05 namun tidak signifikan "
        "setelah koreksi Bonferroni diterapkan. Perbandingan antara Soft Voting dan SVM Tuned mencatat statistik 0,050 dengan p = 0,823 "
        "(disagreement rate hanya 3,21%). Temuan ini dijelaskan secara konsisten oleh analisis keragaman ensemble yang mencatat korelasi "
        "galat sangat tinggi (r = 0,713–0,883; Yule's Q = 0,961–0,987). Karena seluruh model dasar dibangun di atas ruang fitur GLCM homogen "
        "yang sama, kesalahan prediksi cenderung terjadi pada sampel yang identik, sehingga penggabungan voting probabilitas tidak menghasilkan "
        "keberagaman galat yang cukup untuk melampaui kinerja model tunggal terbaik."
    ), bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # 120: h_rq2 (original 4.3 -> now Heading 4.2)
    set_p(body[120], doc, "4.2   Studi Ablasi Pra-Pengolahan, Kuantisasi Ng, Varian Fitur, dan Uji Ketahanan Tepi (RQ2)", bold=True, font_size=10.0, align=WD_ALIGN_PARAGRAPH.LEFT)

    # 121: p_abl_intro
    set_p(body[121], doc, (
        "Studi ablasi parametrik 27 kombinasi (3 praproses × 3 nilai Ng × 3 varian GLCM) dievaluasi pada data uji resmi P2 sebagai "
        "studi eksploratif untuk menjawab RQ2. Hasil representatif dari 11 kombinasi utama disajikan pada Tabel 3."
    ), bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # 122: cap_t3 (Tabel 3 Ablasi)
    set_p(body[122], doc, "Tabel 3 Ringkasan Hasil Studi Ablasi Eksploratif (11 Konfigurasi Representatif pada Data Uji Resmi P2)", bold=False, font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)

    # 125: p_abl_disc
    set_p(body[125], doc, (
        "Analisis parametrik faktor ablasi (RQ2) mengungkapkan tiga temuan utama: (a) varian fitur memberikan pengaruh paling dominan "
        "terhadap spesifisitas—penggunaan GLCM-24 tanpa perataan sudut menurunkan akurasi menjadi 70,99%–74,52% dan spesifisitas hingga "
        "33,33%–44,44% akibat overfitting terhadap orientasi sudut pencitraan, sedangkan GLCM-5 (80,45%) hampir identik dengan GLCM-6 (80,77%) "
        "yang mengonfirmasi redundansi Energy = √ASM; (b) pra-pengolahan Ortiz memberikan peningkatan akurasi sebesar +3,53 poin persentase "
        "dibanding basic pada Ng=32 (80,77% vs 77,24%) dan +2,56 poin dibanding clahe (78,21%); serta (c) kuantisasi tingkat keabuan "
        "yang lebih ringkas Ng=32 secara konsisten melampaui Ng=64 dan Ng=256 (+1,28 poin persentase pada mode basic: 77,24% vs 75,96%) "
        "karena meredam fluktuasi derau lokal. Perlu ditekankan bahwa capaian akurasi 80,77% pada konfigurasi Ortiz Ng=32 GLCM-6 merupakan "
        "temuan eksploratif pasca-evaluasi pada data uji dan bukan hasil seleksi satu langkah model final."
    ), bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # 126: p_pfi_disc
    set_p(body[126], doc, (
        "Analisis Permutation Feature Importance pada data uji resmi P2 mengidentifikasi urutan kontribusi fitur sebagai berikut: "
        "Dissimilarity (7,51% ± 0,58%), Homogeneity (6,46% ± 0,46%), Contrast (4,93% ± 0,78%), Energy (4,11% ± 0,71%), Correlation "
        "(3,46% ± 0,68%), dan ASM (2,29% ± 0,61%), sebagaimana disajikan pada Gambar 4. Berbeda dari penurunan Gini impurity pada Random Forest "
        "yang rentan bias kardinalitas, PFI mengukur penurunan performa riil saat nilai fitur diacak. Meskipun demikian, PFI tetap dapat "
        "terdistorsi oleh multikolinieritas antar-fitur: fitur dengan korelasi matematis sempurna seperti ASM dan Energy berbagi penurunan performa "
        "saat salah satunya diacak, sehingga nilai kepentingannya terbagi. Oleh karena itu, ranking ini diposisikan sebagai ukuran kontribusi "
        "komputasional model dan bukan penanda klinis independen."
    ), bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # 128: cap_t4 (Tabel 4 Edge Crop)
    set_p(body[128], doc, "Tabel 4 Uji Ketahanan Model terhadap Pemotongan Tepi Citra (Peripheral Center Crop Test pada Uji Resmi P2)", bold=False, font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)

    # 131: p_cpu
    set_p(body[131], doc, (
        "Dari sisi efisiensi komputasi CPU, seluruh pipeline inferensi yang disimpan via joblib dievaluasi pada CPU standar tanpa GPU. "
        "Waktu latih pipeline Soft Voting pada 5.232 citra latih resmi adalah 1,987 detik. Pengukuran latensi CPU inferensi utuh per citra "
        "terdiri atas tiga tahap: (1) pra-pengolahan citra (resizing 150×150 piksel, konversi grayscale, dan normalisasi): 7,09 ± 4,17 ms; "
        "(2) ekstraksi enam fitur GLCM pada empat sudut: 10,26 ± 0,14 ms; dan (3) klasifikasi model (StandardScaler transform dan predict_proba "
        "Soft Voting): 7,08 ± 0,16 ms untuk evaluasi citra tunggal (atau 0,0671 ms per citra jika dievaluasi dalam batch 624 citra; total 41,86 ms). "
        "Dengan demikian, total latensi inferensi CPU end-to-end adalah 24,42 ms per citra (setara throughput ~40,9 citra per detik), yang sangat "
        "efisien untuk purwarupa klinis real-time pada perangkat komputer medis standar tanpa akselerator GPU. Purwarupa web interaktif "
        "dapat diakses pada https://pneumonia-ml.streamlit.app/."
    ), bold=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # =========================================================================
    # 8. UPDATE BAB 5 (ANCAMAN VALIDITAS) & BAB 6 (KESIMPULAN)
    # =========================================================================
    print("8. Updating Bab 5 (Threats to Validity) and Bab 6 (Conclusion)...")
    for p in doc.paragraphs[120:128]:
        if "Terdapat beberapa ancaman terhadap validitas" in p.text:
            p.text = (
                "Terdapat beberapa ancaman terhadap validitas (threats to validity) dalam penelitian ini: (1) ukuran himpunan validasi tetap (1.047 citra) "
                "dan prosedur CV 5-fold non-nested membatasi estimasi generalisasi independen mutlak; (2) keberadaan 170 identitas pasien Pneumonia "
                "pada data uji resmi Kaggle (P2) yang tumpang tindih dengan data latih mengindikasikan bahwa hasil P2 lebih tepat diposisikan sebagai "
                "tolok ukur keterbandingan dengan literatur terdahulu daripada estimasi performa pada pasien yang benar-benar baru; (3) identifikasi pasien "
                "kelas Normal tidak selalu eksplisit dalam format penamaan berkas sehingga ketiadaan tumpang tindih pada kelas Normal diartikan sebagai "
                "tidak terdeteksi berdasarkan identifier yang tersedia; (4) pengujian Permutation Feature Importance tetap dipengaruhi oleh redundansi "
                "antar-fitur berkorelasi tinggi (seperti ASM dan Energy); serta (5) populasi dataset terbatas pada pasien anak usia 1 hingga 5 tahun "
                "dari satu pusat kesehatan medis, sehingga generalisasi ke populasi dewasa atau pusat radiologi lain memerlukan validasi eksternal lebih lanjut."
            )
            break

    for p in doc.paragraphs[123:130]:
        if "Penelitian ini menyimpulkan tiga temuan utama" in p.text:
            p.text = (
                "Penelitian ini menyimpulkan tiga temuan utama sesuai pertanyaan penelitian: (1) Sensitivitas protokol partisi (RQ1) memperlihatkan "
                "pola di mana P1 (87,03%) dan P3 (87,61%) menghasilkan akurasi yang hampir setara, sementara P2 turun menjadi 79,33%. Pola ini "
                "mengonfirmasi bahwa perbedaan performa lebih konsisten dipengaruhi oleh variasi rasio kelas dan pergeseran domain uji resmi daripada "
                "semata-mata kebocoran pasien; (2) Analisis faktor teknis (RQ2) mengidentifikasi bahwa varian fitur GLCM-6 dan kuantisasi Ng = 32 "
                "memberikan stabilitas terbaik pada pra-pengolahan Ortiz, dengan Dissimilarity (7,51%) dan Homogeneity (6,46%) sebagai prediktor "
                "paling berpengaruh, serta pemotongan tepi perifer menurunkan akurasi hingga 63,14% yang menandakan pentingnya kontras sangkar dada; "
                "serta (3) Soft Voting Ensemble (RQ3) menghasilkan akurasi 79,33% dan F1-Score 84,88% yang setara secara statistik dengan model dasar "
                "tunggal (McNemar p > 0,0083) akibat tingginya korelasi galat (r = 0,713–0,883), sementara KNN bawaan (k=65) mencatat MCC tertinggi "
                "(0,563). Pipeline inferensi terbukti sangat efisien dengan latensi CPU end-to-end 24,42 ms per citra, mendukung purwarupa berbasis web."
            )
            break

    # =========================================================================
    # 9. RE-ORDER BODY ELEMENTS FOR CHAPTER 4 (RQ1 -> RQ2 -> RQ3)
    # =========================================================================
    print("9. Re-ordering body elements in Chapter 4 (RQ1 -> RQ2 -> RQ3)...")
    # Grab references to key elements in body before removing
    h41 = body[95]
    p_audit = body[96]
    img_g3 = body[97]
    cap_g3 = body[98]
    p_vis = body[99]

    img_g4 = body[100]
    cap_g4 = body[101]

    h_rq3 = body[102]
    p_mod_intro = body[103]
    cap_mod = body[104]
    tbl_mod = body[108]
    p_mod_fn = body[109]
    img_g5 = body[110]
    cap_g5 = body[111]
    img_g6 = body[112]
    cap_g6 = body[113]
    p_stab = body[114]
    p_mod_disc = body[115]
    cap_mcn = body[116]
    tbl_mcn = body[118]
    p_mcn_disc = body[119]

    h_rq2 = body[120]
    p_abl_intro = body[121]
    cap_abl = body[122]
    tbl_abl = body[124]
    p_abl_disc = body[125]
    p_pfi_disc = body[126]
    cap_crop = body[128]
    tbl_crop = body[130]
    p_cpu = body[131]

    cap_lit = body[133]
    tbl_lit = body[138]

    new_order = [
        # 4.1 (RQ1)
        h41, p_audit, img_g3, cap_g3, p_vis,
        # 4.2 (RQ2)
        h_rq2, p_abl_intro, cap_abl, tbl_abl, p_abl_disc, p_pfi_disc, img_g4, cap_g4, cap_crop, tbl_crop, p_cpu,
        # 4.3 (RQ3)
        h_rq3, p_mod_intro, cap_mod, tbl_mod, p_mod_fn, img_g5, cap_g5, img_g6, cap_g6, p_stab, p_mod_disc, cap_mcn, tbl_mcn, p_mcn_disc, cap_lit, tbl_lit
    ]

    # Collect elements between 95 and 140
    old_elems = list(body)[95:140]
    h5 = body[140]

    for el in old_elems:
        body.remove(el)

    insert_idx = body.index(h5)
    for el in new_order:
        body.insert(insert_idx, el)
        insert_idx += 1

    print("Re-ordering Chapter 4 completed successfully!")

    # =========================================================================
    # 10. REMOVE [22] FROM REFERENCES
    # =========================================================================
    print("10. Removing [22] from References...")
    p_ref_22 = None
    for p in doc.paragraphs:
        if p.text.startswith("[22]"):
            p_ref_22 = p
            break
    if p_ref_22:
        body.remove(p_ref_22._element)
        print("  -> Reference [22] removed.")

    # =========================================================================
    # 11. EMBED HIGH-RES 300 DPI FIGURES & ADJUST EXTENTS
    # =========================================================================
    print("11. Embedding 300 DPI figures...")
    image_replacements = {
        'rId6': ('figures_jurnal/gambar2_diagram_metode.png', 6.5, 2.98),
        'rId7': ('figures_jurnal/gambar1_sampel_xray.png', 5.5, 5.58),
        'rId8': ('figures_jurnal/gambar3_feature_importance.png', 6.2, 3.17),
        'rId9': ('figures_jurnal/gambar2_confusion_matrix.png', 6.5, 1.64),
        'rId10': ('figures_jurnal/gambar4_komparasi_metrik.png', 6.5, 2.41),
    }

    for rId, (fpath, width_in, height_in) in image_replacements.items():
        part = doc.part.related_parts.get(rId)
        if part:
            with open(fpath, 'rb') as f:
                part._blob = f.read()
            print(f"  -> Embedded {fpath} into {rId}")

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

    # =========================================================================
    # 12. FINAL SCRUB: CAPITALIZE ORTIZ & ELIMINATE PROHIBITED WORDS
    # =========================================================================
    print("12. Final scrub of prohibited words and Ortiz capitalization...")
    for p in doc.paragraphs:
        p.text = p.text.replace("membuktikan", "menunjukkan").replace("terbukti", "terkonfirmasi").replace("sangat unggul", "kompetitif")
        p.text = p.text.replace("pra-pengolahan ortiz", "pra-pengolahan Ortiz").replace("mode ortiz", "mode Ortiz").replace("konfigurasi ortiz", "konfigurasi Ortiz").replace("ortiz ", "Ortiz ")

    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                cell.text = cell.text.replace("membuktikan", "menunjukkan").replace("terbukti", "terkonfirmasi").replace("sangat unggul", "kompetitif")
                cell.text = cell.text.replace("ortiz (", "Ortiz (").replace("ortiz", "Ortiz")

    # =========================================================================
    # 13. SAVE MANUSCRIPT TO DESKTOP AND REPOSITORY
    # =========================================================================
    print(f"13. Saving final manuscript to {DEST_DESKTOP} and {DEST_REPO}...")
    os.makedirs(os.path.dirname(DEST_DESKTOP), exist_ok=True)
    os.makedirs(os.path.dirname(DEST_REPO), exist_ok=True)
    doc.save(DEST_DESKTOP)
    doc.save(DEST_REPO)
    print("ALL REVISIONS COMPLETED SUCCESSFULLY!")

if __name__ == '__main__':
    main()
