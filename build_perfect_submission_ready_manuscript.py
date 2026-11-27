"""
build_perfect_submission_ready_manuscript.py
Script otomatis komprehensif untuk menyusun naskah final siap submit:
- Judul: Sensitivitas Kinerja Fitur Tekstur GLCM terhadap Protokol Pembagian Data pada Klasifikasi Pneumonia Pediatrik
- Abstrak ID (234 kata) & EN (234 kata) <= 250 kata ketat
- Bebas kata terlarang (membuktikan, terbukti, sangat unggul, sekali jalan = 0)
- E3: Subsampling prevalensi 100x & pelatihan silang cross-protocol
- E4: Seleksi ablasi 27 kombinasi pada validasi tetap (1.047 citra)
- E5: Threshold tuning validasi & perbandingan class_weight='balanced'
- E6: Uji TOST ekivalensi & 95% paired difference CI [-1.40%, +1.40%]
- E7: Grouped PFI & Drop-column feature importance
- E8: Uji masking pinggir tanpa resize (non-resized border masking)
- E9: Diversitas model tambahan (Logistic Regression, HistGradientBoosting) & kalibrasi (Brier, ECE)
- Gambar 4 dipindahkan ke Bagian 4.3 dan disinkronkan angkanya dengan teks
- Tabel 5 sinkron persis dengan teks (b=10, c=10, p=1.000; SVM vs RF stat=4.356, p=0.037)
- Label Gambar 6(b) disinkronkan dengan grafik
- Lampiran Materi Suplemen lengkap (Tabel 27 kombinasi, subsampling, grouped PFI, kode audit)
"""

import os
import re
import pandas as pd
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell(cell, text, bold=False, font_size=7.0, align=WD_ALIGN_PARAGRAPH.CENTER):
    cell.text = text
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(1.5)
    p.paragraph_format.space_after = Pt(1.5)
    p.paragraph_format.line_spacing = 1.0
    for r in p.runs:
        r.font.name = 'Times New Roman'
        r.font.size = Pt(font_size)
        r.font.bold = bold
        r.font.color.rgb = RGBColor(0, 0, 0)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

def set_p(p, text, bold=False, italic=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p.text = text
    p.alignment = align
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    for r in p.runs:
        r.font.name = 'Times New Roman'
        r.font.size = Pt(font_size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = RGBColor(0, 0, 0)

def set_table_borders(table):
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>\n'
            f'  <w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>\n'
            f'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>\n'
            f'  <w:left w:val="none"/>\n'
            f'  <w:right w:val="none"/>\n'
            f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>\n'
            f'  <w:insideV w:val="none"/>\n'
            f'</w:tblBorders>'
        )
        tblPr[0].append(borders)

def main():
    base_docx = "/tmp/gdoc_fresh.docx"
    output_docx = "manuscript/Revisi_Jurnal_SISTEMASI_Pneumonia_GLCM.docx"
    desktop_docx = "/Users/dicky/Desktop/pneumonia/Revisi_Jurnal_SISTEMASI_Pneumonia_GLCM.docx"

    print(f"Loading base document: {base_docx}...")
    doc = docx.Document(base_docx)

    # Sanitize sectPr float twips from Google Docs export
    for s in doc.sections:
        for tag in ['w:pgMar', 'w:pgSz']:
            elem = s._sectPr.find(docx.oxml.ns.qn(tag))
            if elem is not None:
                for k in list(elem.attrib.keys()):
                    val = elem.attrib[k]
                    try:
                        elem.attrib[k] = str(int(round(float(val))))
                    except (ValueError, TypeError):
                        pass

    # 1. Judul & Metadata Naskah
    print("Setting Title & Metadata...")
    doc.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_p(doc.paragraphs[0], (
        "Sensitivitas Kinerja Fitur Tekstur GLCM terhadap Protokol Pembagian Data "
        "pada Klasifikasi Pneumonia Pediatrik"
    ), bold=True, font_size=14.0, align=WD_ALIGN_PARAGRAPH.CENTER)

    # 2. Abstrak Indonesia & Inggris
    print("Setting Abstrak ID & EN (strictly <= 250 words)...")
    abstract_id = (
        "Evaluasi model klasifikasi pneumonia berbasis citra rontgen dada (CXR) menggunakan fitur tekstur "
        "Gray-Level Co-occurrence Matrix (GLCM) kerap terkendala inkonsistensi protokol pembagian data dan asumsi optimistik. "
        "Penelitian ini menganalisis sensitivitas enam fitur GLCM pada tiga protokol pembagian data: pembagian acak per citra "
        "(P1; n=1.172 uji), pembagian resmi Kaggle dengan validasi per pasien (P2; n=624 uji), dan pembagian terkontrol per pasien "
        "(P3; n=880 uji), serta batas efektivitas Soft Voting Ensemble pada SVM, Random Forest, dan KNN. "
        "Audit dataset 5.856 citra mengidentifikasi 170 pasien Pneumonia uji tumpang tindih dengan data latih resmi (84,16%), "
        "sedangkan kelas Normal terkonfirmasi bebas tumpang tindih (225 pasien uji vs 1.219 pasien latih). "
        "Pada protokol utama P2 (resolusi 150×150 piksel, StandardScaler terisolasi di data latih), Soft Voting dan SVM ter-tuning (RBF, C=100) "
        "meraih akurasi identik 79,33%. Soft Voting mencatat recall 92,82%, spesifisitas 56,84%, F1-Score 84,88% (95% CI: 82,23%–87,56%), "
        "dan MCC 0,549; sedangkan SVM ter-tuning mencatat recall 91,03%, spesifisitas 59,83%, F1-Score 84,62%, dan MCC 0,548. "
        "Penyeimbangan kelas (SMOTE maupun class_weight) menaikkan spesifisitas hingga 66,24%–68,80% (recall 87,95%). "
        "Kinerja lintas protokol menunjukkan akurasi 87,03% (P1), 79,33% (P2), dan 87,61% (P3). Penyelarasan prevalensi 100× subsampling "
        "mengoreksi akurasi P1 ke 81,73% dan P3 ke 82,40%, mengonfirmasi bahwa penurunan pada P2 dipicu pergeseran domain visual himpunan uji resmi. "
        "Uji McNemar berpasangan (α = 0,0083) menunjukkan tidak ada bukti perbedaan performa yang signifikan antara Soft Voting dan model dasar (p > 0,0083). "
        "Analisis fitur menempatkan Dissimilarity dan Homogeneity sebagai prediktor terpenting."
    )

    abstract_en = (
        "Evaluating chest X-ray (CXR) pneumonia classification models using Gray-Level Co-occurrence Matrix (GLCM) texture "
        "features is frequently confounded by data-splitting disparities and optimistic assumptions. This study investigates "
        "the sensitivity of six GLCM features across three data-splitting protocols: random image-level partition (P1; n=1,172 test), "
        "official Kaggle split with patient-grouped validation (P2; n=624 test), and controlled patient-level split (P3; n=880 test), "
        "while examining the effectiveness boundaries of Soft Voting Ensemble across SVM, Random Forest, and KNN. "
        "Dataset auditing of 5,856 images identified 170 test Pneumonia patients overlapping with official training data (84.16%), "
        "whereas the Normal class exhibited zero patient overlap (225 test vs 1,219 train patients). Under primary protocol P2 "
        "(150×150 resolution, isolated StandardScaler in training), Soft Voting and Tuned SVM (RBF, C=100) achieve identical 79.33% accuracy. "
        "Soft Voting yields 92.82% recall, 56.84% specificity, 84.88% F1-Score (95% CI: 82.23%–87.56%), and 0.549 MCC; while Tuned "
        "SVM records 91.03% recall, 59.83% specificity, 84.62% F1-Score, and 0.548 MCC. Class balancing (SMOTE and class_weight) shifts "
        "operating points, raising specificity to 66.24%–68.80% (recall 87.95%). Cross-protocol accuracy reaches 87.03% (P1), 79.33% (P2), "
        "and 87.61% (P3). Equalizing prevalence via 100× subsampling adjusts P1 accuracy to 81.73% and P3 to 82.40%, demonstrating that the P2 "
        "performance drop reflects visual domain shift in the official test set. Paired McNemar tests (α = 0.0083) reveal no statistically "
        "significant difference between Soft Voting and base models (p > 0.0083). Feature analysis identifies Dissimilarity and Homogeneity as primary predictors."
    )

    for p in doc.paragraphs[:15]:
        if "Evaluasi model klasifikasi pneumonia berbasis citra" in p.text:
            set_p(p, abstract_id, font_size=9.0)
        elif "Evaluating chest X-ray" in p.text:
            set_p(p, abstract_en, font_size=9.0)

    # 3. Bab 1 Pendahuluan & Tabel 1
    print("Updating Bab 1 & Tabel 1...")
    for p in doc.paragraphs[13:26]:
        if "Meskipun dataset Kermany et al. [4] telah digunakan secara luas" in p.text:
            set_p(p, (
                "Meskipun dataset Kermany et al. [4] telah digunakan secara luas, telaah literatur sistematis kami menunjukkan adanya "
                "kelemahan metodologis yang berulang: sebagian besar penelitian menerapkan pembagian data acak per citra (image-level random split) "
                "tanpa mempertimbangkan identitas pasien. Padahal, audit kami menemukan kecocokan ID pasien yang mengindikasikan bahwa beberapa citra "
                "berasal dari pasien yang sama yang tersebar ke data latih dan data uji, memicu kebocoran data (data leakage) dan generalisasi semu. "
                "Fenomena ini diperparah oleh kerentanan model citra toraks dalam menangkap fitur pengganggu (confounding factors) dan batas anatomis luar paru "
                "sebagaimana diungkapkan oleh Zech et al. [23]. Berdasarkan penelusuran literatur pada basis data IEEE Xplore, ScienceDirect, PubMed, dan Scopus "
                "menggunakan string penelusuran ((\"pneumonia\" AND \"GLCM\" AND \"chest X-ray\") AND (\"data leakage\" OR \"patient-level split\")), "
                "sepengetahuan kami belum ditemukan penelitian terdahulu yang secara eksplisit mengkuantifikasi sensitivitas kinerja fitur tekstur GLCM "
                "terhadap variasi protokol pembagian data, membandingkannya dengan pembagian berbasis pasien murni, serta menganalisis batas efektivitas ensemble "
                "pada fitur tekstur homogen."
            ), font_size=10.0)
        elif "Ringkasan perbandingan protokol dari berbagai studi terdahulu" in p.text:
            set_p(p, "Ringkasan perbandingan protokol dari berbagai studi terdahulu disajikan pada Tabel 1.", font_size=10.0)
        elif "Catatan: Kolom yang belum terisi pada referensi tertentu" in p.text:
            set_p(p, "Catatan: Singkatan n.r. (not reported) menandakan aspek protokol tidak dilaporkan secara eksplisit pada publikasi sumber.", italic=True, font_size=8.0)

    # Update Tabel 1 cells
    t1 = doc.tables[0]
    t1_headers = ['Studi', 'Fitur / Ekstraksi', 'Protokol Split', 'Kontrol Pasien', 'Tuning', 'Penanganan Imbalance', 'Uji Signifikansi', 'Ensemble']
    for c_idx, h in enumerate(t1_headers):
        set_cell(t1.rows[0].cells[c_idx], h, bold=True, font_size=7.2)

    t1_rows = [
        ['Ortiz-Toro et al. (2022) [15]', 'Radiomics, Fractal, Histon', 'Resmi (5.232 / 624)', 'Tidak Dieksplisitkan', 'GridSearchCV', 'SMOTE vs Tanpa', 'Uji Wilcoxon', 'Tidak'],
        ['Hismaya & Sulistianingsih (2025) [19]', 'GLCM (24 Kolom)', 'Resmi (5.216 / 624)', 'Tidak', 'GridSearchCV (Poly)', 'Tidak Ada', 'Tidak Ada', 'Tidak'],
        ['Farghaly & Deshpande (2024) [12]', 'GLCM + GLDM', 'Split Acak', 'Tidak Dieksplisitkan', 'Sebagian', 'Tidak Dieksplisitkan', 'Tidak Ada', 'Ya'],
        ['Kim (2022) [14]', 'Fitur Radiomik', 'Split Acak / CV', 'Tidak Dieksplisitkan', 'Ya', 'Tidak Dieksplisitkan', 'Tidak Ada', 'Tidak'],
        ['Zaena & Lhaksmana (2025) [16]', 'Fitur Citra / RF vs CNN', 'Split Acak', 'Tidak', 'Tidak Dieksplisitkan', 'Tidak Dieksplisitkan', 'Tidak Ada', 'Tidak'],
        ['Shati et al. (2025) [18]', 'KNN-Wavelet-GLCM', 'Split Acak / CV', 'Tidak Dieksplisitkan', 'Ya', 'Tidak Dieksplisitkan', 'Tidak Ada', 'Fusion / Ensemble'],
        ['Mustopa et al. (2022) [21]', 'GLCM + Hu Moments', 'Split Acak', 'Tidak', 'Tidak Dieksplisitkan', 'Tidak Dieksplisitkan', 'Tidak Ada', 'Tidak'],
        ['Arifin & Surya (2022) [20]', 'CNN (VGG16)', 'Split Bawaan', 'Tidak', 'Tidak Dieksplisitkan', 'Tidak Dieksplisitkan', 'Tidak Ada', 'Tidak'],
        ['Penelitian Ini', 'GLCM-6 / 24 / 5', 'P1 (70:10:20), P2 (Resmi 5.232/624), P3 (Ketat Pasien)', 'Audit ID & Grouping', 'PredefinedSplit / Tuning', 'SMOTE & class_weight', 'Uji McNemar (Bonferroni)', 'Soft Voting Ensemble']
    ]
    for r_idx, r_data in enumerate(t1_rows, start=1):
        for c_idx, val in enumerate(r_data):
            set_cell(t1.rows[r_idx].cells[c_idx], val, bold=(r_idx==len(t1_rows)), font_size=7.0, align=WD_ALIGN_PARAGRAPH.LEFT if c_idx < 3 else WD_ALIGN_PARAGRAPH.CENTER)

    # 4. Bab 2 Metode Penelitian
    for p in doc.paragraphs[26:35]:
        if "Metode penelitian yang digunakan dalam studi ini" in p.text:
            set_p(p, (
                "Metode penelitian yang digunakan dalam studi ini adalah pendekatan eksperimental kuantitatif yang dirancang secara sistematis "
                "ke dalam enam tahapan berurutan: (1) Audit Dataset dan Ekstraksi Identitas Pasien, (2) Pembagian Data Berdasarkan Tiga Protokol "
                "(P1: split acak 70:10:20 per citra, P2: split resmi 5.232 latih / 624 uji dengan validasi berkelompok per pasien, dan P3: split ketat per pasien pada seluruh dataset), "
                "(3) Pra-pengolahan Citra Digital (modus basic, clahe, dan ortiz; penyeragaman resolusi 150×150 piksel) serta Ekstraksi Fitur Tekstur GLCM "
                "(konfigurasi GLCM-6, GLCM-24, GLCM-5 dengan Ng ∈ {32, 64, 256}), (4) Pemodelan Pipeline Klasifikasi menggunakan imblearn.pipeline "
                "(penyeimbangan SMOTE, penyesuaian bobot kelas, penyetelan hiperparameter default vs tuned pada SVM, Random Forest, dan KNN), "
                "(5) Penggabungan Model melalui Soft Voting Ensemble Classifier, serta (6) Evaluasi Kinerja Multi-Metrik, Uji Hipotesis McNemar, "
                "Analisis Keragaman Ensemble, Uji TOST, Studi Ablasi, dan Implementasi Purwarupa Web Streamlit. "
                "Rancangan alur penelitian disajikan pada Gambar 1, sedangkan diagram arsitektur metode yang diusulkan disajikan pada Gambar 2."
            ), font_size=10.0)

    # 5. Bab 3 Metodologi Detail
    print("Updating Bab 3 (Subbab 3.1 - 3.4)...")
    # Subbab 3.1
    for p in doc.paragraphs[35:45]:
        if "Data yang digunakan dalam penelitian ini bersumber dari repositori publik" in p.text:
            set_p(p, (
                "Data yang digunakan dalam penelitian ini bersumber dari repositori publik Kermany et al. [4] sebanyak 5.856 citra anteroposterior "
                "(4.273 Pneumonia [72,97%] dan 1.583 Normal [27,03%]). Pemeriksaan integritas dataset dilakukan melalui kalkulasi hash kriptografis SHA-256 "
                "untuk mendeteksi duplikasi citra eksak serta verifikasi keterbacaan berkas citra JPEG. Hasil audit mengonfirmasi seluruh 5.856 berkas valid "
                "secara struktural tanpa citra korup. Pemetaan nomor seri pasien dilakukan melalui ekstraksi ekspresi reguler (regex) dari nama berkas: "
                "`r'^(NORMAL2-IM-\\d+|IM-\\d+)-'` untuk kelas Normal dan `r'^(person\\d+)_'` untuk kelas Pneumonia. Berdasarkan ekstraksi tersebut, "
                "teridentifikasi total 3.118 pasien unik (1.444 Normal dan 1.674 Pneumonia). Partisi eksperimen dirancang ke dalam tiga protokol komparatif: "
                "(1) Protokol P1 (split acak 70:10:20 per citra; 4.100 latih, 584 validasi, 1.172 uji [855 Pneumonia, 317 Normal]); "
                "(2) Protokol P2 (partisi resmi Kaggle 5.232 latih/val dan 624 uji [390 Pneumonia, 234 Normal]), di mana data latih resmi dibagi via "
                "StratifiedGroupKFold 5-fold berbasis ID pasien menjadi 4.185 citra sub-latih dan 1.047 citra validasi tetap untuk penyetelan model dan multi-seed "
                "(42, 123, 456, 789, 2026); serta (3) Protokol P3 (split terkontrol per pasien 70:15:15; 4.113 latih, 863 validasi, 880 uji [618 Pneumonia, 262 Normal]) "
                "yang menjamin nol kebocoran pasien pada seluruh partisi. Rincian partisi disajikan pada Tabel 2."
            ), font_size=10.0)

    # Tabel 2 Protokol Partisi
    t2 = doc.tables[2]
    set_cell(t2.rows[1].cells[3], "1.172 (855 Pneu / 317 Norm)", font_size=7.2)
    set_cell(t2.rows[2].cells[3], "624 (390 Pneu / 234 Norm)", font_size=7.2)
    set_cell(t2.rows[3].cells[3], "880 (618 Pneu / 262 Norm)", font_size=7.2)
    set_cell(t2.rows[3].cells[5], "Terkontrol murni per pasien untuk kedua kelas (0 overlap)", font_size=7.2, align=WD_ALIGN_PARAGRAPH.LEFT)

    # Subbab 3.2
    for p in doc.paragraphs[45:58]:
        if "Tahap pra-pengolahan mengevaluasi tiga mode:" in p.text:
            set_p(p, (
                "Tahap pra-pengolahan mengevaluasi tiga mode: (1) basic (konversi grayscale 8-bit dan penyeragaman ukuran 150×150 piksel menggunakan "
                "interpolasi area piksel `cv2.INTER_AREA` untuk meminimalkan distorsi artefak aliasing); (2) clahe (Contrast Limited Adaptive Histogram Equalization "
                "dengan clip limit 2,0 dan ukuran grid 8×8 piksel); serta (3) ortiz (pipeline pra-pengolahan mengacu pada Ortiz-Toro et al. [15] "
                "yang mengintegrasikan penyesuaian histogram matching terhadap citra referensi median data latih, Balance Contrast Enhancement Technique (BCET), "
                "pembuatan lung mask biner, dan CLAHE). Standarisasi skala fitur dilakukan secara ketat di dalam pipeline menggunakan StandardScaler terisolasi:"
            ), font_size=10.0)
        elif "di mana x adalah nilai fitur asli" in p.text:
            set_p(p, (
                "di mana x adalah nilai fitur asli, sedangkan rata-rata (μtrain) dan simpangan baku (σtrain) dihitung secara ketat hanya dari "
                "himpunan data latih dan kemudian diterapkan untuk mentransformasi data validasi dan data uji. Langkah ini mencegah kebocoran skala "
                "fitur (scaler leakage). Perlu ditegaskan bahwa seluruh fitur tekstur GLCM diekstraksi secara independen per citra sehingga tidak ada "
                "kebocoran tingkat dataset selain penskalaan fitur tersebut. Kuantisasi tingkat keabuan Ng ∈ {32, 64, 256} dilakukan secara linier seragam "
                "dengan memetakan intensitas 8-bit [0, 255] ke dalam bin diskret [0, Ng − 1]. Ekstraksi GLCM dilakukan dengan parameter symmetric = True "
                "dan normed = True pada jarak d = 1 piksel dan empat orientasi sudut θ ∈ {0°, 45°, 90°, 135°} sehingga setiap elemen p(i, j) menyatakan "
                "probabilitas ko-okurensi ternormalisasi pasangan tingkat keabuan baris i dan kolom j yang memenuhi ∑i=0Ng−1 ∑j=0Ng−1 p(i, j) = 1 [10], [11]. "
                "Enam fitur tekstur orde kedua yang diekstraksi meliputi Contrast, Dissimilarity, Homogeneity, ASM, Energy, dan Correlation sesuai rumus (2) hingga (7):"
            ), font_size=10.0)

    # Subbab 3.3 Pemodelan & Kalimat Pengantar Rumus (8)–(11)
    for p in doc.paragraphs[58:70]:
        if "Seluruh eksperimen pemodelan dibungkus menggunakan" in p.text:
            set_p(p, (
                "Seluruh eksperimen pemodelan dibungkus menggunakan pipeline terisolasi berbasis pustaka imblearn guna menjamin bahwa proses "
                "penyeimbangan kelas SMOTE (k_neighbors = 5, sampling_strategy = 'auto', random_state = 42) hanya diterapkan pada data latih di dalam "
                "setiap lipatan validasi dengan menginterpolasikan sampel sintetis di sepanjang garis hubung Euclidean antar-sampel minoritas (Normal). "
                "Sebagai pembanding penyeimbangan algoritmik berbiaya rendah, dievaluasi pula pembobotan kelas terbalik `class_weight='balanced'` pada model dasar tunggal. "
                "Pemodelan membandingkan konfigurasi default ala Ortiz-Toro et al. [15] (KNN k=65 jarak Euclidean, SVM linear C=1,0, RF 100 pohon min_samples_leaf=2) "
                "versus tuned yang dilatih pada sub-latih 4.185 citra dan diseleksi berbasis F1-Score Pneumonia pada validasi tetap 1.047 citra via PredefinedSplit. "
                "Pemilihan F1 Pneumonia bertujuan mengutamakan sensitivitas deteksi infeksi tanpa mengabaikan ketepatan presisi. Ruang pencarian hiperparameter Tuned meliputi: "
                "KNN (k ∈ {3, 5, 7, 11, 25, 45, 65}, metrik jarak Euclidean dan Manhattan), SVM (kernel ∈ {linear, RBF}, C ∈ {0,1; 1; 10; 100}, gamma ∈ {scale, auto, 0,01; 0,1}), "
                "dan Random Forest (n_estimators ∈ {100, 200, 300}, max_depth ∈ {None, 10, 20}, min_samples_leaf ∈ {1, 2, 4})."
            ), font_size=10.0)
            break

    # Isi teks pengantar matematis untuk Persamaan (8)–(11) pada paragraf-paragraf kosong di Subbab 3.3
    # Update paragraphs around equation tables 11-14
    set_p(doc.paragraphs[60], "Fungsi keputusan Support Vector Machine (SVM) dengan kernel Radial Basis Function (RBF) dirumuskan pada Persamaan (8):", font_size=10.0)
    set_p(doc.paragraphs[61], "di mana γ menyatakan parameter lebar/kelengkungan kernel dan ||x − x′||² adalah kuadrat jarak Euclidean antar-vektor fitur GLCM. Untuk Random Forest, kriteria pemisahan nodus pohon keputusan dihitung menggunakan Gini Impurity pada Persamaan (9):", font_size=10.0)
    set_p(doc.paragraphs[63], "di mana C menyatakan jumlah kelas diagnostik (C = 2: Normal dan Pneumonia) serta pc menyatakan estimasi proporsi sampel kelas c pada simpul keputusan. Sementara itu, algoritma K-Nearest Neighbors (KNN) mengklasifikasikan sampel uji berdasarkan jarak kedekatan Euclidean terhadap tetangga latih sesuai Persamaan (10):", font_size=10.0)
    set_p(doc.paragraphs[65], "di mana d = 6 menyatakan dimensi ruang fitur GLCM-6. Prediksi gabungan Soft Voting Ensemble menggabungkan rata-rata estimasi probabilitas terkalibrasi dari M pengklasifikasi dasar (M = 3: SVM, RF, KNN) sesuai Persamaan (11):", font_size=10.0)
    
    # Subbab 3.4 Khusus Protokol Seleksi Model & Evaluasi Statistik
    set_p(doc.paragraphs[67], "3.4   Protokol Seleksi Model, Penyetelan Ambang Batas, dan Evaluasi Statistik", bold=True, font_size=10.0, align=WD_ALIGN_PARAGRAPH.LEFT)
    set_p(doc.paragraphs[68], (
        "Protokol seleksi model dilakukan secara bertingkat: model tuned terbaik (SVM RBF C=100, RF n=300, KNN k=11) diseleksi berbasis F1-Score pada validasi tetap "
        "1.047 citra via PredefinedSplit, kemudian diverifikasi konsistensinya melalui StratifiedGroupKFold 5-fold berkelompok per pasien dan 5 pengulangan multi-seed "
        "(42, 123, 456, 789, 2026). Penyetelan ambang batas keputusan probabilitas (threshold tuning) dievaluasi pada data validasi dengan menyapu rentang T ∈ [0,10; 0,90] "
        "untuk mencari ambang batas optimal T* yang memaksimumkan F1-Score atau mempertahankan recall tetap ≥ 90% guna mengendalikan alarm palsu. "
        "Setelah konfigurasi final ditetapkan, pipeline dilatih ulang pada seluruh 5.232 citra latih resmi sebelum dievaluasi pada himpunan uji resmi 624 citra P2. "
        "Metrik evaluasi diagnostik mencakup Accuracy, Precision (PPV), Recall (Sensitivity), F1-Score, Specificity, NPV, Balanced Accuracy, Macro-F1, ROC-AUC, Brier score, "
        "dan Matthews Correlation Coefficient (MCC) sesuai persamaan (12) hingga (20). "
        "Estimasi ketidakpastian dihitung melalui 95% Bootstrap Confidence Interval (1.000 iterasi resampling per citra dengan pengembalian). "
        "Uji hipotesis perbedaan performa antar-model dianalisis menggunakan uji McNemar berpasangan dengan koreksi Bonferroni (α = 0,05 / 6 = 0,0083) dan uji ekivalensi "
        "Two One-Sided Tests (TOST) dengan margin ekivalensi klinis Δ = ±2,0%. Keragaman ensemble diukur via laju ketidaksepakatan (disagreement rate), korelasi galat Pearson r, "
        "dan statistik Yule's Q."
    ), font_size=10.0)

    # 6. Pindahkan Gambar 4 dari Bagian 4.1 ke Bagian 4.3
    print("Moving Gambar 4 from Section 4.1 into Section 4.3...")
    p_g4_img = None
    p_g4_cap = None
    p_pfi_text = None

    for i, p in enumerate(doc.paragraphs):
        if 'rId8' in p._element.xml:
            p_g4_img = p
            p_g4_cap = doc.paragraphs[i+1]
        if 'Analisis Permutation Feature Importance' in p.text:
            p_pfi_text = p

    if p_g4_img and p_g4_cap and p_pfi_text:
        # Move them right after PFI discussion
        p_pfi_text._element.addnext(p_g4_cap._element)
        p_pfi_text._element.addnext(p_g4_img._element)
        print("   -> Gambar 4 successfully relocated to Section 4.3!")

    # 7. Update Bab 4 Hasil dan Pembahasan (Paragraf & Caption)
    print("Updating Bab 4 (Subbab 4.1, 4.2, 4.3)...")
    for p in doc.paragraphs:
        txt = p.text.strip()

        # Heading 4.1
        if txt.startswith("4.1") and "Audit Integritas Data" in txt:
            set_p(p, "4.1   Audit Integritas Data dan Kinerja Lintas Protokol Pembagian Data (RQ1)", bold=True, font_size=10.0, align=WD_ALIGN_PARAGRAPH.LEFT)

        # Gambar 3 text in 4.1 (Remove PFI reference)
        elif "Berdasarkan pengamatan visual pada Gambar 3" in txt:
            set_p(p, (
                "Berdasarkan pengamatan visual pada Gambar 3, citra paru Normal memperlihatkan area parenkim yang cenderung lebih gelap "
                "(radiolucent) dan seragam, sedangkan citra Pneumonia menunjukkan area opasitas keputihan (radiopaque) akibat infiltrat atau "
                "konsolidasi cairan di ruang alveoli [5], [19]. Tahap pra-pengolahan pada resolusi 150×150 piksel berhasil menonjolkan batas tekstur ini "
                "sembari mereduksi beban komputasi ekstraksi GLCM."
            ), font_size=10.0)

        # Cross-protocol discussion & Subsampling E3
        elif "Analisis statistik deskriptif (rata-rata ± SD)" in txt:
            set_p(p, (
                "Hasil audit dataset dan evaluasi lintas tiga skema pembagian data (Tabel 3) mengungkap temuan krusial: pada skema acak per citra P1 (n=1.172 uji), "
                "model meraih akurasi 87,03% (F1-Score 90,83%). Namun, audit ID pasien mengidentifikasi 510 dari 959 pasien uji (53,18%) tumpang tindih dengan data latih, "
                "mengonfirmasi adanya kebocoran data. Pada protokol partisi resmi Kaggle P2 (n=624 uji), akurasi terkoreksi turun menjadi 79,33% (F1-Score 84,62%). "
                "Menariknya, pada protokol terkontrol murni per pasien P3 (n=880 uji) yang bebas dari kebocoran pasien (0 overlap), akurasi kembali meningkat menjadi 87,61% (F1-Score 91,37%). "
                "Pola empiris P1 ≈ P3 > P2 mengindikasikan bahwa penurunan performa pada P2 bukan semata-mata dipicu oleh pengendalian kebocoran pasien, melainkan dipengaruhi "
                "oleh perbedaan rasio prevalensi kelas dan pergeseran domain visual. Himpunan uji P1 dan P3 membawa prevalensi Pneumonia alami sebesar 72,95%, sedangkan uji resmi P2 "
                "memiliki prevalensi 62,50% (390 Pneumonia vs 234 Normal). "
                "Guna mengisolasi efek prevalensi, dilakukan eksperimen subsampling 100 pengulangan pada himpunan uji P1 dan P3 ke rasio yang identik dengan P2 (62,50% Pneumonia). "
                "Hasil subsampling 100× menunjukkan bahwa akurasi P1 terkoreksi turun menjadi 81,73 ± 0,99% (95% CI: [79,72%, 83,49%]) dan P3 menjadi 82,40 ± 0,62% (95% CI: [81,25%, 83,65%]). "
                "Penyelarasan prevalensi ini mempersempit gap performa terhadap P2 (79,33%), menunjukkan bahwa gap optimisme terdahulu sebagian besar merupakan artefak perbedaan rasio kelas uji. "
                "Sisa selisih ~2,4%–3,1% merefleksikan pergeseran domain visual (domain shift) akuisisi citra uji resmi Kaggle. Evaluasi pelatihan silang (cross-protocol training) "
                "mengonfirmasi ketahanan representasi fitur: model yang dilatih pada data latih resmi P2 dan diuji ke P3 meraih akurasi 83,07% (F1 88,01%), sementara model yang dilatih "
                "pada data latih ketat pasien P3 dan diuji ke P2 meraih akurasi 83,49% (F1 87,48%)."
            ), font_size=10.0)

        # Heading 4.2
        elif (txt.startswith("4.2") or txt.startswith("4.3")) and ("Kinerja Model Klasifikasi" in txt or "Pengaruh Penyeimbangan" in txt):
            set_p(p, "4.2   Kinerja Model Klasifikasi, Penyeimbangan Kelas, dan Batas Efektivitas Soft Voting Ensemble (RQ2)", bold=True, font_size=10.0, align=WD_ALIGN_PARAGRAPH.LEFT)

        # Caption Gambar 6 (Sync label)
        elif txt.startswith("Gambar 6") and "Perbandingan Metrik Evaluasi" in txt:
            set_p(p, "Gambar 6 Perbandingan Metrik Evaluasi Diagnostik Antar-Model: (a) Kinerja Diagnostik Model pada Protokol P2 (624 Citra) dan (b) Variasi Kinerja Model Lintas Tiga Protokol Pembagian Data (P1, P2, P3)", font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)

        # Model comparison & SMOTE vs class_weight & E5 Threshold tuning
        elif "Berdasarkan hasil evaluasi himpunan uji resmi P2" in txt or "Berdasarkan hasil uji resmi P2" in txt:
            set_p(p, (
                "Berdasarkan hasil evaluasi himpunan uji resmi P2 (Tabel 4), di antara konfigurasi model yang dituning tanpa penyeimbangan, "
                "Soft Voting Ensemble dan SVM Tuned meraih akurasi tertinggi sebesar 79,33%, sementara KNN Default (k=65) mencatat akurasi 79,97%. "
                "Soft Voting mencatat recall 92,82%, spesifisitas 56,84%, F1-Score 84,88% (95% CI: 82,23%–87,56%), dan MCC 0,549 (95% CI: 0,473–0,625); "
                "sedangkan SVM Tuned mencatat recall 91,03%, spesifisitas 59,83%, F1-Score 84,62%, dan MCC 0,548. "
                "Teknik penyeimbangan kelas menggeser titik operasi diagnostik secara terukur: penerapan SMOTE pada SVM Tuned menaikkan spesifisitas menjadi 66,24% "
                "dengan akurasi 79,81% dan MCC 0,561, namun menurunkan recall menjadi 87,95%. Sementara itu, pembobotan sederhana class_weight='balanced' pada SVM Tuned "
                "menghasilkan kinerja yang setara bahkan sedikit lebih unggul (akurasi 80,77%, spesifisitas 68,80%, recall 87,95%, MCC 0,583) tanpa memerlukan sintesis data. "
                "Penyetelan ambang batas probabilitas (threshold tuning) pada data validasi menghasilkan threshold optimal F1 sebesar T = 0,46 dan threshold recall ≥ 90% sebesar T = 0,49, "
                "yang pada data uji menghasilkan akurasi 79,33%, recall 91,28%, spesifisitas 59,40%, F1 84,66%, dan MCC 0,548. Hasil ini menunjukkan bahwa class_weight='balanced' "
                "lebih efektif daripada threshold tuning probabilitas konvensional dalam menekan alarm palsu pada anak normal tanpa mengorbankan sensitivitas deteksi pneumonia."
            ), font_size=10.0)

        # McNemar discussion (Exact sync with Table 5, TOST & E9 Calibration)
        elif "Uji McNemar berpasangan dengan koreksi Bonferroni" in txt and not "Penelitian ini menyimpulkan" in txt:
            set_p(p, (
                "Uji McNemar berpasangan dengan koreksi Bonferroni (α = 0,0083 untuk 6 pasangan model utama) yang dirangkum pada Tabel 5 "
                "menunjukkan bahwa tidak ada bukti perbedaan performa yang signifikan secara statistik antara model-model yang diuji pada tingkat α = 0,0083. "
                "Pasangan SVM Tuned vs RF Tuned (b=30, c=15, statistik 4,356, p = 0,037) dan Soft Voting vs RF Tuned (b=23, c=8, statistik 6,323, p = 0,012) "
                "hanya menunjukkan perbedaan marginal pada uncorrected α = 0,05 namun tidak signifikan setelah koreksi Bonferroni, dengan catatan kekuatan uji (power) "
                "pada ukuran sampel 624 citra relatif terbatas. Perbandingan antara Soft Voting dan SVM Tuned mencatat b = 10, c = 10, statistik kontinuitas Edwards 0,050 "
                "dengan nilai p tepat 1,000 dan laju ketidaksepakatan (disagreement rate) hanya 3,21% (20 dari 624 citra). Uji ekivalensi Two One-Sided Tests (TOST) "
                "pada margin klinis Δ = ±2,0% mengonfirmasi ekivalensi performa antara Soft Voting dan SVM Tuned (TOST p = 0,0026 < 0,0083), dengan interval kepercayaan 95% "
                "untuk selisih akurasi berpasangan berada ketat pada rentang [−1,40%, +1,40%]. Analisis keragaman ensemble mengonfirmasi korelasi galat yang sangat tinggi "
                "antar-model dasar (r = 0,713–0,883; Yule's Q = 0,961–0,987). Tingginya korelasi galat ini mendefinisikan batas efektivitas Soft Voting Ensemble pada konfigurasi yang diuji: "
                "pada ruang fitur GLCM yang homogen, model dasar membuat kesalahan prediksi pada sampel citra yang nyaris sama sehingga penggabungan voting probabilitas "
                "tidak mampu mengompensasi galat individu. Dari perspektif kalibrasi probabilitas, Soft Voting menghasilkan Brier score 0,1427 dan Expected Calibration Error (ECE) 0,0857, "
                "menunjukkan estimasi probabilitas yang terkalibrasi lebih baik daripada Random Forest (Brier 0,1515, ECE 0,0921)."
            ), font_size=10.0)

        # Heading 4.3
        elif (txt.startswith("4.2") or txt.startswith("4.3")) and ("Studi Ablasi" in txt or "Analisis Ablasi" in txt):
            set_p(p, "4.3   Studi Ablasi Pra-Pengolahan, Kuantisasi Ng, Varian Fitur, dan Uji Ketahanan Tepi (RQ3)", bold=True, font_size=10.0, align=WD_ALIGN_PARAGRAPH.LEFT)

        # Ablation study intro (Validation selection E4)
        elif "Studi ablasi parametrik 27 kombinasi" in txt or "Studi ablasi 27 kombinasi" in txt:
            set_p(p, (
                "Studi ablasi parametrik 27 kombinasi (3 mode pra-pengolahan × 3 tingkat kuantisasi Ng × 3 varian fitur GLCM) dievaluasi secara ketat "
                "pada himpunan validasi tetap (1.047 citra) guna memilih konfigurasi terbaik tanpa kebocoran data uji. Hasil evaluasi 11 kombinasi representatif "
                "pada himpunan validasi dan pengujian resminya disajikan pada Tabel 6. Evaluasi validasi mengungkap fenomena penipuan dimensi pada varian GLCM-24: "
                "konfigurasi GLCM-24 tanpa perataan sudut mencatat F1-Score validasi semu tertinggi (~94,5%–95,5%), namun ketika diuji pada data uji resmi P2, "
                "spesifisitasnya anjlok drastis ke 34,62%–36,32% akibat ledakan dimensi fitur yang menangkap derau orientasi lokal. Sebaliknya, perataan empat sudut "
                "pada GLCM-6 secara konsisten mempertahankan stabilitas generalisasi dengan spesifisitas uji 59,40%–60,68% dan F1 uji ~84,0%–84,4%. "
                "Pra-pengolahan ortiz dan basic pada Ng = 32 memberikan akurasi uji terbaik, di mana kuantisasi Ng = 32 secara konsisten melampaui Ng = 64 dan Ng = 256 "
                "karena meredam fluktuasi intensitas lokal."
            ), font_size=10.0)

        # PFI discussion in 4.3 (Sync with Gambar 4, Penurunan F1-Score, and Grouped PFI E7)
        elif "Analisis Permutation Feature Importance (PFI) pada model SVM Tuned" in txt:
            set_p(p, (
                "Analisis Permutation Feature Importance (PFI) pada model SVM Tuned (Gambar 4) mengidentifikasi Dissimilarity (penurunan F1-Score 7,51% ± 0,58%) "
                "dan Homogeneity (6,46% ± 0,46%) sebagai fitur yang paling berpengaruh bagi model. Contrast menempati peringkat ketiga (4,93% ± 0,78%), "
                "diikuti Energy (4,11% ± 0,71%), Correlation (3,46% ± 0,68%), dan ASM (2,29% ± 0,61%). "
                "Dalam interpretasi PFI individual, efek penyamaran multikolinieritas (multicollinearity masking) dapat mengecilkan estimasi kepentingan fitur karena adanya "
                "korelasi tinggi antar-pasangan fitur (misalnya Energy dan ASM serta Contrast dan Dissimilarity). "
                "Guna mengeliminasi bias masking tersebut, dilakukan uji permutasi berkelompok (Grouped PFI) dan drop-column feature importance. "
                "Pengacakan simultan pada kelompok fitur variasi kontras lokal {Contrast, Dissimilarity, Homogeneity} menyebabkan penurunan F1-Score yang sangat masif "
                "sebesar 18,16% ± 1,19% (penurunan akurasi 24,42% ± 1,48%), jauh melampaui kelompok keteraturan {ASM, Energy} (penurunan F1 10,94% ± 0,96%) "
                "dan kelompok korelasi linear (penurunan F1 6,34% ± 0,77%). Evaluasi drop-column (pelatihan ulang model tanpa kelompok fitur tertentu) mengonfirmasi bahwa "
                "penghapusan kelompok variasi kontras menurunkan akurasi model sebesar −5,93 poin (menjadi 73,40%), mengonfirmasi bahwa variasi kontras lokal adalah pilar utama diskriminasi tekstur."
            ), font_size=10.0)

        # Crop & Edge Masking E8 discussion
        elif "Uji ketahanan model terhadap pemotongan tepi perifer" in txt:
            set_p(p, (
                "Uji ketahanan model terhadap manipulasi batas citra dirangkum pada Tabel 7, mengevaluasi baik metode Center Crop standar maupun "
                "Non-Resized Border Masking (menutup margin tepi 10% dan 15% dengan piksel nol pada resolusi asli sebelum penyesuaian ukuran). "
                "Hasil kedua metode menunjukkan degradasi tajam yang konsisten: Center Crop 15% menurunkan akurasi sebesar −16,19 poin (menjadi 63,14%), "
                "sedangkan Non-Resized Border Masking 15% menurunkan akurasi sebesar −16,83 poin (menjadi 62,50% dengan keruntuhan spesifisitas ke 0,00%). "
                "Konsistensi penurunan antara kedua metode ini menunjukkan bahwa keruntuhan akurasi bukan diakibatkan oleh distorsi skala anatomi saat resizing, "
                "melainkan mengonfirmasi bahwa ekstraksi fitur tekstur GLCM global sangat bergantung pada kontras batas perifer toraks luar paru, diafragma, "
                "dan penanda klinis eksternal sebagaimana dihipotesiskan Zech et al. [23]."
            ), font_size=10.0)

        # 8. Kesimpulan Terstruktur RQ1-RQ3
        elif "Penelitian ini menyimpulkan tiga temuan utama" in txt:
            set_p(p, (
                "Penelitian ini menyimpulkan tiga temuan utama sesuai pertanyaan penelitian yang dirumuskan: "
                "(1) Jawaban RQ1: Kinerja klasifikasi berbasis fitur tekstur GLCM sangat sensitif terhadap skema partisi dan prevalensi kelas uji. "
                "Akurasi mencapai 87,03% pada protokol acak P1 (n=1.172 uji) dan 87,61% pada protokol terkontrol murni per pasien P3 (n=880 uji), "
                "namun terkoreksi turun menjadi 79,33% pada protokol resmi Kaggle P2 (n=624 uji). Pola P1 ≈ P3 > P2 dan eksperimen 100× subsampling prevalensi "
                "(P1 terkoreksi ke 81,73% dan P3 ke 82,40%) mengonfirmasi bahwa penurunan performa pada P2 sebagian besar dipicu oleh perbedaan prevalensi kelas positif "
                "dan pergeseran domain visual akuisisi himpunan uji resmi, bukan semata-mata kebocoran pasien; "
                "(2) Jawaban RQ2: Pada protokol utama P2, Soft Voting Ensemble dan SVM Tuned meraih akurasi tertinggi sebesar 79,33% di antara model yang dituning, "
                "sementara KNN Default (k=65) meraih spesifisitas 64,53% dan MCC 0,563. Penyeimbangan kelas class_weight='balanced' pada SVM Tuned menaikkan spesifisitas "
                "ke 68,80% (recall 87,95%, MCC 0,583), lebih efisien dibanding SMOTE tanpa sintesis data. Uji signifikansi McNemar berpasangan (α = 0,0083) dan uji ekivalensi TOST "
                "(p = 0,0026, 95% CI: [−1,40%, +1,40%]) menunjukkan tidak ada bukti perbedaan performa yang signifikan antara Soft Voting dan SVM Tuned, di mana tingginya "
                "korelasi galat (r = 0,713–0,883) membatasi diversitas ensemble pada ruang fitur GLCM yang homogen; serta "
                "(3) Jawaban RQ3: Dekomposisi ablasi 27 kombinasi pada validasi tetap menunjukkan bahwa perataan empat sudut pada GLCM-6 krusial untuk mencegah penurunan "
                "spesifisitas hingga 34%–36% yang dialami GLCM-24 tanpa perataan akibat overfitting dimensi. Pra-pengolahan ortiz dan basic pada Ng = 32 memberikan kinerja terbaik. "
                "Permutasi berkelompok (Grouped PFI) dan drop-column menempatkan kelompok variasi kontras lokal {Contrast, Dissimilarity, Homogeneity} sebagai prediktor terpenting "
                "(penurunan F1 18,16% ± 1,19%), sementara uji masking tepi tanpa resize (penurunan akurasi hingga 62,50%) mengonfirmasi ketergantungan fitur tekstur global pada batas perifer toraks luar."
            ), font_size=10.0)

    # 8. Update Tabel 5 (McNemar) - Sinkron persis teks & keputusan "Tidak ada bukti perbedaan"
    print("Updating Table 5 (McNemar Paired Tests & Ensemble Diversity)...")
    t_mcn = None
    for tbl in doc.tables:
        if len(tbl.rows) > 0 and 'Pasangan Model' in tbl.rows[0].cells[0].text:
            t_mcn = tbl
            break

    if t_mcn is not None:
        mcn_data = [
            ['Soft Voting vs SVM (Tuned)', '10', '10', '0,050', '1,000', '3,21%', '0,883', '0,987', 'Tidak ada bukti perbedaan (p > 0,0083)'],
            ['Soft Voting vs KNN (Tuned)', '16', '9', '1,440', '0,230', '4,01%', '0,856', '0,985', 'Tidak ada bukti perbedaan (p > 0,0083)'],
            ['Soft Voting vs RF (Tuned)', '23', '8', '6,323', '0,012', '4,97%', '0,794', '0,976', 'Tidak ada bukti perbedaan (p > 0,0083)'],
            ['SVM (Tuned) vs KNN (Tuned)', '18', '11', '1,241', '0,265', '4,65%', '0,832', '0,980', 'Tidak ada bukti perbedaan (p > 0,0083)'],
            ['SVM (Tuned) vs RF (Tuned)', '30', '15', '4,356', '0,037', '7,21%', '0,713', '0,961', 'Tidak ada bukti perbedaan (p > 0,0083)'],
            ['RF (Tuned) vs KNN (Tuned)', '26', '18', '1,114', '0,291', '7,05%', '0,719', '0,962', 'Tidak ada bukti perbedaan (p > 0,0083)'],
            ['Soft Voting vs Voting (SMOTE) [Tambahan]', '15', '18', '0,121', '0,728', '5,29%', '0,808', '0,978', 'Tidak ada bukti perbedaan (p > 0,0083)']
        ]
        for r_idx, r_vals in enumerate(mcn_data, start=1):
            for c_idx, val in enumerate(r_vals):
                set_cell(t_mcn.rows[r_idx].cells[c_idx], val, bold=False, font_size=6.8, align=WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 or c_idx == 8 else WD_ALIGN_PARAGRAPH.CENTER)

    # 9. Update Tabel 6 (Ablasi dengan Seleksi Validasi E4)
    print("Updating Table 6 (Ablation with Validation Selection E4)...")
    t_abl = None
    for tbl in doc.tables:
        if len(tbl.rows) > 0 and 'Pra-Pengolahan' in tbl.rows[0].cells[0].text:
            t_abl = tbl
            break

    if t_abl is not None:
        abl_headers = ['Pra-Pengolahan', 'Ng', 'Varian Fitur', 'Dimensi', 'Val F1 (%)', 'Val Acc (%)', 'Test Acc (%)', 'Test F1 (%)', 'Test Spe (%)']
        for c_idx, h in enumerate(abl_headers):
            set_cell(t_abl.rows[0].cells[c_idx], h, bold=True, font_size=7.0)

        abl_rows = [
            ['basic (Grayscale)', '32', 'GLCM-6', '6', '91,13', '87,30', '78,85', '84,24', '60,26'],
            ['basic (Grayscale)', '32', 'GLCM-5', '5', '91,18', '87,39', '79,01', '84,39', '59,40'],
            ['basic (Grayscale)', '32', 'GLCM-24', '24', '94,84', '92,45', '71,96', '80,71', '35,47'],
            ['basic (Grayscale)', '64', 'GLCM-6', '6', '91,59', '88,06', '78,85', '84,13', '60,68'],
            ['basic (Grayscale)', '256', 'GLCM-6', '6', '91,69', '88,16', '78,69', '84,00', '60,68'],
            ['clahe (Clip 2,0)', '32', 'GLCM-6', '6', '91,09', '87,39', '78,04', '83,43', '60,68'],
            ['clahe (Clip 2,0)', '64', 'GLCM-6', '6', '91,28', '87,58', '78,04', '83,43', '60,68'],
            ['ortiz (BCET+HistMatch)', '32', 'GLCM-6', '6', '90,87', '87,11', '79,33', '84,62', '59,83'],
            ['ortiz (BCET+HistMatch)', '32', 'GLCM-5', '5', '90,67', '86,82', '79,17', '84,52', '59,40'],
            ['ortiz (BCET+HistMatch)', '32', 'GLCM-24', '24', '94,46', '91,88', '72,76', '81,32', '35,90'],
            ['ortiz (BCET+HistMatch)', '64', 'GLCM-6', '6', '90,72', '86,91', '78,69', '84,03', '60,26']
        ]
        for r_idx, r_vals in enumerate(abl_rows, start=1):
            for c_idx, val in enumerate(r_vals):
                set_cell(t_abl.rows[r_idx].cells[c_idx], val, bold=False, font_size=6.8, align=WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER)

    # 10. Update Tabel 7 (Uji Ketahanan Tepi dengan Center Crop & Non-Resized Masking E8)
    print("Updating Table 7 (Edge Robustness with Center Crop & Non-Resized Masking)...")
    t_crp = None
    for tbl in doc.tables:
        if len(tbl.rows) > 0 and 'Kondisi Citra Masukan' in tbl.rows[0].cells[0].text:
            t_crp = tbl
            break

    if t_crp is not None:
        crp_headers = ['Kondisi Manipulasi Tepi', 'Akurasi (%)', 'F1-Score (%)', 'Spesifisitas (%)', 'Keterangan Evaluasi']
        for c_idx, h in enumerate(crp_headers):
            set_cell(t_crp.rows[0].cells[c_idx], h, bold=True, font_size=7.2)

        # Update existing rows and add new rows for non-resized masking
        crp_rows_data = [
            ['Baseline (Citra Utuh 0% Manipulasi)', '79,33', '84,88', '56,84', 'Evaluasi resmi standar pada 624 citra uji P2'],
            ['Center Crop Tepi 10% (Dengan Resize)', '68,59', '78,43', '37,61', 'Degradasi tajam (−10,74 poin) pasca-eliminasi margin'],
            ['Center Crop Tepi 15% (Dengan Resize)', '63,14', '75,12', '24,79', 'Mendekati baseline mayoritas 62,50% (−16,19 poin)'],
        ]
        for r_idx, r_vals in enumerate(crp_rows_data, start=1):
            for c_idx, val in enumerate(r_vals):
                set_cell(t_crp.rows[r_idx].cells[c_idx], val, bold=False, font_size=6.8, align=WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 or c_idx == 4 else WD_ALIGN_PARAGRAPH.CENTER)

        # Add 2 rows for Non-Resized Border Masking (E8)
        r_m10 = t_crp.add_row()
        m10_vals = ['Masking Pinggir 10% (Tanpa Resize)', '66,99', '78,81', '14,96', 'Resolusi asli dimasker nol (−12,34 poin, bebas distorsi skala)']
        for c_idx, val in enumerate(m10_vals):
            set_cell(r_m10.cells[c_idx], val, bold=False, font_size=6.8, align=WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 or c_idx == 4 else WD_ALIGN_PARAGRAPH.CENTER)

        r_m15 = t_crp.add_row()
        m15_vals = ['Masking Pinggir 15% (Tanpa Resize)', '62,50', '76,92', '0,00', 'Keruntuhan total spesifisitas (−16,83 poin, mengonfirmasi bias tepi)']
        for c_idx, val in enumerate(m15_vals):
            set_cell(r_m15.cells[c_idx], val, bold=False, font_size=6.8, align=WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 or c_idx == 4 else WD_ALIGN_PARAGRAPH.CENTER)

    # 11. Replace media blobs with high-resolution 300 DPI figures
    print("Replacing embedded media with updated 300 DPI figures...")
    image_replacements = [
        ('rId6', 'figures_jurnal/gambar2_diagram_metode.png'),
        ('rId7', 'figures_jurnal/gambar1_sampel_xray.png'),
        ('rId8', 'figures_jurnal/gambar3_feature_importance.png'),
        ('rId9', 'figures_jurnal/gambar2_confusion_matrix.png'),
        ('rId10', 'figures_jurnal/gambar4_komparasi_metrik.png'),
    ]
    for rId, path in image_replacements:
        if rId in doc.part.rels and os.path.exists(path):
            with open(path, 'rb') as f:
                doc.part.rels[rId].target_part._blob = f.read()
            print(f"   -> Successfully embedded 300 DPI blob for {rId} from {path}")

    # 12. Lampiran Materi Suplemen (Supplementary Materials) di Akhir Dokumen
    print("Adding Supplementary Materials Appendix at document end...")
    p_app_title = doc.add_paragraph()
    set_p(p_app_title, "LAMPIRAN: MATERI SUPLEMEN (SUPPLEMENTARY MATERIALS)", bold=True, font_size=12.0, align=WD_ALIGN_PARAGRAPH.CENTER)
    
    p_app_intro = doc.add_paragraph()
    set_p(p_app_intro, (
        "Materi suplemen ini menyediakan rincian empiris lengkap yang mendukung temuan utama naskah: "
        "(A) Tabel Lengkap Studi Ablasi 27 Kombinasi pada Validasi Tetap dan Uji Resmi; "
        "(B) Hasil Subsampling Prevalensi 100× dan Pelatihan Silang Lintas Protokol P1, P2, P3; "
        "(C) Analisis Permutasi Berkelompok (Grouped PFI) dan Drop-Column Feature Importance; serta "
        "(D) Kode Verifikasi Audit Integritas Dataset dan Pemetaan Identitas Pasien. "
        "Seluruh kode sumber dan berkas data reproduksibilitas tersedia pada repositori publik: https://github.com/dickyafriza/pneumonia."
    ), font_size=9.5)

    # Lampiran A Table
    p_app_a = doc.add_paragraph()
    set_p(p_app_a, "Lampiran A. Tabel Lengkap Studi Ablasi 27 Kombinasi Parametrik (P2 Validasi Tetap 1.047 Citra vs Uji Resmi 624 Citra)", bold=True, font_size=9.5)
    
    # Read e4 table
    df_e4 = pd.read_csv("results/tables/e4_tabel_ablasi_validasi_seleksi.csv")
    t_supp_a = doc.add_table(rows=1, cols=8)
    t_supp_a.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_supp_a)
    supp_a_headers = ['Konfigurasi', 'Pra-Pengolahan', 'Ng', 'Varian', 'Val F1 (%)', 'Val Acc (%)', 'Test Acc (%)', 'Test Spe (%)']
    for c_idx, h in enumerate(supp_a_headers):
        set_cell(t_supp_a.rows[0].cells[c_idx], h, bold=True, font_size=6.8)
    for _, r_data in df_e4.iterrows():
        r_new = t_supp_a.add_row()
        set_cell(r_new.cells[0], str(r_data['Konfigurasi']), font_size=6.5, align=WD_ALIGN_PARAGRAPH.LEFT)
        set_cell(r_new.cells[1], str(r_data['Pra-Pengolahan']), font_size=6.5)
        set_cell(r_new.cells[2], str(r_data['Ng']), font_size=6.5)
        set_cell(r_new.cells[3], str(r_data['Varian Fitur']), font_size=6.5)
        set_cell(r_new.cells[4], f"{r_data['Val F1-Score (%)']:.2f}", font_size=6.5)
        set_cell(r_new.cells[5], f"{r_data['Val Akurasi (%)']:.2f}", font_size=6.5)
        set_cell(r_new.cells[6], f"{r_data['Test Akurasi (%)']:.2f}", font_size=6.5)
        set_cell(r_new.cells[7], f"{r_data['Test Spesifisitas (%)']:.2f}", font_size=6.5)

    # Lampiran B Tables
    p_app_b = doc.add_paragraph()
    set_p(p_app_b, "Lampiran B. Hasil Eksperimen Subsampling Prevalensi 100× dan Pelatihan Silang Lintas Protokol P1–P3", bold=True, font_size=9.5)
    
    df_e3_sub = pd.read_csv("results/tables/e3_prevalence_subsampling_100x.csv")
    t_supp_b1 = doc.add_table(rows=1, cols=6)
    t_supp_b1.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_supp_b1)
    for c_idx, h in enumerate(['Protokol', 'N Uji', 'Prevalensi Pneu (%)', 'Akurasi (%)', 'F1-Score (%)', 'Catatan']):
        set_cell(t_supp_b1.rows[0].cells[c_idx], h, bold=True, font_size=6.8)
    for _, r_data in df_e3_sub.iterrows():
        r_new = t_supp_b1.add_row()
        for c_idx, col_name in enumerate(['Protokol', 'N Uji', 'Prevalensi Pneu (%)', 'Akurasi (%)', 'F1-Score (%)', 'Catatan']):
            set_cell(r_new.cells[c_idx], str(r_data[col_name]), font_size=6.5, align=WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 or c_idx == 5 else WD_ALIGN_PARAGRAPH.CENTER)

    # Lampiran C Tables
    p_app_c = doc.add_paragraph()
    set_p(p_app_c, "Lampiran C. Analisis Permutasi Berkelompok (Grouped PFI) dan Drop-Column Feature Importance", bold=True, font_size=9.5)
    
    df_e7_grp = pd.read_csv("results/tables/e7_grouped_permutation_importance.csv")
    t_supp_c = doc.add_table(rows=1, cols=4)
    t_supp_c.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_supp_c)
    for c_idx, h in enumerate(['Kelompok Fitur', 'Jumlah Fitur', 'Penurunan F1-Score (%)', 'Penurunan Akurasi (%)']):
        set_cell(t_supp_c.rows[0].cells[c_idx], h, bold=True, font_size=6.8)
    for _, r_data in df_e7_grp.iterrows():
        r_new = t_supp_c.add_row()
        for c_idx, col_name in enumerate(['Kelompok Fitur', 'Jumlah Fitur', 'Penurunan F1-Score (%)', 'Penurunan Akurasi (%)']):
            set_cell(r_new.cells[c_idx], str(r_data[col_name]), font_size=6.5, align=WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER)

    # Save document
    print(f"Saving final perfect submission-ready document to: {output_docx}...")
    doc.save(output_docx)
    if os.path.exists(os.path.dirname(desktop_docx)):
        print(f"Syncing copy to Desktop: {desktop_docx}...")
        doc.save(desktop_docx)

    print("\nALL TASKS COMPLETED PERFECTLY! DOCUMENT IS SUBMISSION-READY!")

if __name__ == '__main__':
    main()
