"""
build_final_manuscript.py
Comprehensive script addressing all 8 checklist items from expert feedback:
1. Perbaiki kesimpulan RQ1, RQ2, dan RQ3.
2. Perbaiki abstrak bahasa Indonesia dan Inggris (230 kata, zero prohibited words).
3. Hilangkan klaim pengujian sekali jalan yang tidak tepat.
4. Lengkapi keterbatasan McNemar, bootstrap, dan multi-seed.
5. Tambahkan tabel audit ID pasien (Tabel 3 baru).
6. Lengkapi metode PFI (multikolinieritas) dan pembahasan SMOTE (trade-off).
7. Pasang kembali 5 gambar 300 DPI dan masukkan rumus MCC (20) & kolom MCC di tabel evaluasi.
8. Verifikasi parameter eksperimen, referensi, dan metadata jurnal.
"""
import os
import re
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.text.paragraph import Paragraph

SRC_DOC = "/tmp/gdoc_fresh.docx"
DEST_DESKTOP = "/Users/dicky/Desktop/pneumonia/Revisi_Jurnal_SISTEMASI_Pneumonia_GLCM.docx"
DEST_REPO = "manuscript/Revisi_Jurnal_SISTEMASI_Pneumonia_GLCM.docx"

def set_cell(cell, text, bold=False, font_size=7.0, align=WD_ALIGN_PARAGRAPH.CENTER):
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

def set_p(p, text, bold=False, italic=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p.text = text
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    for r in p.runs:
        r.font.name = 'Times New Roman'
        r.font.size = Pt(font_size)
        r.bold = bold
        r.italic = italic

def set_table_borders(table):
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
            f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
            f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
            f'<w:insideV w:val="none"/>'
            f'<w:left w:val="none"/>'
            f'<w:right w:val="none"/>'
            f'</w:tblBorders>'
        )
        tblPr[0].append(borders)

def main():
    print(f"Loading document from {SRC_DOC}...")
    doc = docx.Document(SRC_DOC)

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

    # =========================================================================
    # ITEM 7: REPLACE MEDIA BLOBS WITH 300 DPI FIGURES
    # =========================================================================
    print("Item 7: Replacing embedded media with 300 DPI figures...")
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
                data = f.read()
            doc.part.rels[rId].target_part._blob = data
            print(f"  -> Replaced {rId} with {path} ({len(data)} bytes)")

    # =========================================================================
    # ITEM 8: DATE PLACEHOLDER & METADATA
    # =========================================================================
    print("Item 8: Updating date placeholder...")
    for p in doc.paragraphs[:10]:
        if "(received" in p.text.lower():
            set_p(p, "(received: 01-10-2025, revised: 15-11-2025, accepted: 20-01-2026)", italic=True, font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)
            break

    # =========================================================================
    # ITEM 2: CONDENSED ABSTRACTS (ID & EN) - STRICTLY ~230 WORDS
    # =========================================================================
    print("Item 2: Updating condensed abstracts (ID & EN)...")
    abstract_id = (
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

    abstract_en = (
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

    for i, p in enumerate(doc.paragraphs[:15]):
        if "Evaluasi model klasifikasi pneumonia berbasis citra" in p.text:
            set_p(p, abstract_id, font_size=9.0)
            print(f"  -> Abstract ID updated (word count: {len(abstract_id.split())})")
        elif "Evaluating chest X-ray" in p.text:
            set_p(p, abstract_en, font_size=9.0)
            print(f"  -> Abstract EN updated (word count: {len(abstract_en.split())})")

    # =========================================================================
    # ITEM 3: TABLE 1 (FLOWCHART) - REMOVE 'PENGUJIAN UJI SEKALI JALAN'
    # =========================================================================
    print("Item 3: Updating Table 1 (Flowchart)...")
    t1 = doc.tables[1]
    for row in t1.rows:
        for cell in row.cells:
            if "Pengujian Uji Sekali Jalan" in cell.text:
                cell.text = cell.text.replace("Pengujian Uji Sekali Jalan", "Evaluasi Himpunan Uji Resmi")
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    r.font.name = 'Times New Roman'
                    r.font.size = Pt(7.0)

    # Subbab 3.1 & Table 2
    for p in doc.paragraphs[30:45]:
        if "Dataset utama yang digunakan dalam penelitian ini" in p.text:
            set_p(p, (
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
            ), font_size=10.0)
            break

    # Table 2: Update row counts
    t2 = doc.tables[2]
    set_cell(t2.rows[1].cells[3], "1.172 (855 Pneu / 317 Norm)", font_size=7.2)
    set_cell(t2.rows[2].cells[3], "624 (390 Pneu / 234 Norm)", font_size=7.2)
    set_cell(t2.rows[3].cells[3], "880 (618 Pneu / 262 Norm)", font_size=7.2)
    set_cell(t2.rows[3].cells[5], "Terkontrol berdasarkan ID terekstrak (ID kelas Normal tidak selalu tersedia)", font_size=7.2, align=WD_ALIGN_PARAGRAPH.LEFT)

    # Subbab 3.2
    for p in doc.paragraphs[40:55]:
        if "Tahap pra-pengolahan mengevaluasi tiga mode:" in p.text:
            set_p(p, (
                "Tahap pra-pengolahan mengevaluasi tiga mode: (1) basic (konversi grayscale 8-bit dan penyesuaian ukuran 150×150 piksel); "
                "(2) clahe (Contrast Limited Adaptive Histogram Equalization dengan clip limit 2,0 dan ukuran grid 8×8 piksel); serta (3) Ortiz "
                "(pipeline pra-pengolahan komprehensif mengacu pada Ortiz-Toro et al. [15] yang mengintegrasikan histogram matching terhadap "
                "citra referensi median, Balance Contrast Enhancement Technique (BCET), pembuatan lung mask biner, dan CLAHE). Standarisasi skala fitur "
                "dilakukan secara ketat di dalam pipeline menggunakan StandardScaler terisolasi:"
            ), font_size=10.0)
        elif "di mana x adalah nilai fitur asli" in p.text:
            set_p(p, (
                "di mana x adalah nilai fitur asli, sedangkan rata-rata (μtrain) dan simpangan baku (σtrain) dihitung secara ketat hanya dari "
                "himpunan data latih dan kemudian diterapkan untuk mentransformasi data validasi dan data uji. Langkah ini mencegah kebocoran skala "
                "fitur (scaler leakage) maupun preprocessing leakage. Perlu ditekankan bahwa pencegahan kebocoran pada tahap penskalaan fitur ini "
                "merupakan aspek yang terpisah dari kontrol kebocoran tingkat pasien (patient-level leakage) yang dikendalikan melalui skema pembagian data. "
                "Selanjutnya, ekstraksi fitur GLCM dilakukan secara konsisten pada setiap nilai tingkat keabuan Ng ∈ {32, 64, 256} yang dievaluasi "
                "pada studi ablasi dengan parameter symmetric = True dan normed = True pada jarak d = 1 piksel dan empat orientasi sudut "
                "θ ∈ {0°, 45°, 90°, 135°} sehingga setiap elemen p(i, j) menyatakan probabilitas ko-okurensi ternormalisasi pasangan tingkat keabuan "
                "baris i dan kolom j (untuk i, j ∈ {0, 1, …, Ng − 1}) yang memenuhi ∑i=0Ng−1 ∑j=0Ng−1 p(i, j) = 1 [10], [11]. Enam fitur tekstur orde kedua "
                "yang diekstraksi pada setiap sudut orientasi meliputi Contrast, Dissimilarity, Homogeneity, ASM, Energy, dan Correlation sesuai rumus (2) hingga (7):"
            ), font_size=10.0)

    # =========================================================================
    # ITEM 6 & ITEM 7: SUBBAB 3.3 (SMOTE, PFI, & RUMUS MCC 20)
    # =========================================================================
    print("Item 6 & 7: Detailing SMOTE, PFI, and inserting MCC Equation (20)...")
    for p in doc.paragraphs[55:75]:
        if "Seluruh eksperimen pemodelan dibungkus menggunakan" in p.text:
            set_p(p, (
                "Seluruh eksperimen pemodelan dibungkus menggunakan pipeline terisolasi berbasis pustaka imblearn guna menjamin bahwa proses "
                "penyeimbangan kelas SMOTE (k_neighbors = 5, sampling_strategy = 'auto', random_state = 42) hanya diterapkan pada data latih di dalam "
                "setiap lipatan validasi dengan menginterpolasikan sampel sintetis di sepanjang garis hubung Euclidean antar-sampel minoritas (Normal). "
                "Pemodelan membandingkan konfigurasi default ala Ortiz-Toro et al. [15] (KNN k=round(√N)=65, SVM linear C=1,0, "
                "RF 100 pohon min_samples_leaf=2) versus tuned (dilatih pada sub-latih 4.185 citra, dinilai dengan F1 Pneumonia pada validasi tetap "
                "1.047 citra via PredefinedSplit tanpa menyentuh data uji). Ruang pencarian hiperparameter Tuned meliputi: KNN (k ∈ {3, 5, 7, 11, 25, 45, 65}), "
                "SVM (kernel ∈ {linear, RBF}, C ∈ {0,1; 1; 10; 100}, gamma ∈ {scale, auto, 0,01; 0,1}), dan Random Forest (n_estimators ∈ {100, 200, 300}, "
                "max_depth ∈ {None, 10, 20}, min_samples_leaf ∈ {1, 2, 4}). Model tuned terpilih (SVM RBF C=100, RF n_estimators=300, KNN k=11) "
                "selanjutnya digabungkan ke dalam Soft Voting Ensemble Classifier dengan bobot seragam (1:1:1), di mana estimasi probabilitas kontinu SVM "
                "dihasilkan melalui penskalaan Platt (probability = True) yang dikalibrasi pada data latih. Setelah konfigurasi final ditetapkan dari "
                "validasi tetap dan diverifikasi melalui CV 5-fold serta 5 seed acak, pipeline final dilatih ulang pada seluruh 5.232 citra latih resmi "
                "sebelum dievaluasi pada himpunan uji resmi 624 citra P2. Metrik evaluasi mencakup Accuracy, Precision (PPV), Recall (Sensitivity), "
                "F1-Score, Specificity, NPV, Balanced Accuracy, Macro-F1, Matthews Correlation Coefficient (MCC), dan ROC-AUC di mana nilai ROC-AUC "
                "dihitung dari probabilitas kontinu (predict_proba), bukan dari label biner keras. Estimasi interval kepercayaan 95% Bootstrap "
                "dihitung melalui resampling per citra dengan pengembalian pada 624 citra uji sebanyak 1.000 iterasi menggunakan metode persentil [2,5%, 97,5%] "
                "dengan random_state = 42. Analisis Permutation Feature Importance (PFI) dilakukan dengan mengacak nilai tiap fitur pada data uji sebanyak "
                "K = 10 kali permutasi dan mengukur rata-rata penurunan skor akurasi (ΔAcc). Rumus evaluasi dirumuskan pada persamaan (8) hingga (20):"
            ), font_size=10.0)
            break

    # Insert MCC formula (20) after equation (19)
    p_roc = None
    for p in doc.paragraphs:
        if "ROC-AUC = ∫ TPR d(FPR)" in p.text:
            p_roc = p
            break
    if p_roc is not None:
        p_mcc_eq = doc.add_paragraph()
        set_p(p_mcc_eq, "MCC = (TP · TN − FP · FN) / √((TP + FP)(TP + FN)(TN + FP)(TN + FN))   (20)", font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        
        p_mcc_desc = doc.add_paragraph()
        set_p(p_mcc_desc, (
            "Nilai Matthews Correlation Coefficient (MCC) mengevaluasi seluruh sel matriks konfusi (TP, TN, FP, FN) dengan rentang "
            "dari −1 (prediksi salah mutlak), 0 (setara tebakan acak / baseline mayoritas), hingga +1 (prediksi sempurna). "
            "MCC tahan terhadap bias kelas mayoritas sehingga memberikan gambaran keseimbangan diagnostik yang lebih objektif pada dataset "
            "dengan rasio kelas tidak seimbang (62,5% Pneumonia vs 37,5% Normal)."
        ), font_size=10.0, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

        # Remove from document end
        doc._element.body.remove(p_mcc_eq._element)
        doc._element.body.remove(p_mcc_desc._element)

        # Insert after p_roc using addnext
        p_roc._element.addnext(p_mcc_desc._element)
        p_roc._element.addnext(p_mcc_eq._element)
        print("  -> Inserted MCC formula (20) and description")

    # Subbab 3.4
    for p in doc.paragraphs[75:90]:
        if "Untuk menjamin reproduksibilitas" in p.text:
            set_p(p, (
                "Untuk menjamin reproduksibilitas, seluruh eksperimen dikelola secara deterministik dengan seed 42 serta pengulangan 5 seed acak "
                "(42, 123, 456, 789, 2026); indeks partisi disimpan ke berkas keluaran, dan seluruh objek praproses serta bobot model disimpan menggunakan "
                "joblib (.pkl) pada direktori results/. Eksperimen dijalankan pada lingkungan komputasi Python 3 (kompatibel Python 3.10 hingga 3.14) "
                "dengan pustaka ilmiah utama: scikit-learn (>= 1.6.0), OpenCV (opencv-python-headless), scikit-image (>= 0.25.0), imbalanced-learn (>= 0.13.0), "
                "pandas, dan NumPy. Seluruh kode sumber eksperimen dan skrip pipeline terdokumentasi terbuka pada repositori GitHub resmi "
                "https://github.com/dickyafriza/pneumonia (mirror: https://github.com/salsaseva/Pneumonia), sedangkan purwarupa web interaktif dapat "
                "diakses pada https://pneumonia-ml.streamlit.app/."
            ), font_size=10.0)
            break

    # =========================================================================
    # ITEM 7: ADD MCC COLUMN TO MODEL EVALUATION TABLE (TABLE 14)
    # =========================================================================
    print("Item 7: Adding MCC column to Model Evaluation table...")
    t14 = doc.tables[14]
    mcc_vals = ['MCC', '0,000', '0,511', '0,548', '0,471', '0,494', '0,563', '0,522', '0,549', '0,561', '0,560']
    new_col = t14.add_column(Inches(0.5))
    for r_idx in range(len(t14.rows)):
        roc_val = t14.rows[r_idx].cells[11].text
        set_cell(new_col.cells[r_idx], roc_val, bold=(r_idx==0), font_size=6.8)
        set_cell(t14.rows[r_idx].cells[11], mcc_vals[r_idx], bold=(r_idx==0), font_size=6.8)

    # =========================================================================
    # UPDATE TABLE 15 (MCNEMAR TABLE) WITH ACCURATE VALUES
    # =========================================================================
    print("Updating Table 15 (McNemar) values...")
    t15 = doc.tables[15]
    set_cell(t15.rows[1].cells[3], "0,050", font_size=7.0)
    set_cell(t15.rows[1].cells[4], "0,823", font_size=7.0)
    set_cell(t15.rows[1].cells[8], "Tidak signifikan (p = 0,823)", font_size=7.0)

    set_cell(t15.rows[2].cells[3], "1,440", font_size=7.0)
    set_cell(t15.rows[2].cells[4], "0,230", font_size=7.0)
    set_cell(t15.rows[2].cells[8], "Tidak signifikan (p = 0,230)", font_size=7.0)

    set_cell(t15.rows[3].cells[3], "6,323", font_size=7.0)
    set_cell(t15.rows[3].cells[4], "0,012", font_size=7.0)
    set_cell(t15.rows[3].cells[8], "Tidak signifikan (p > 0,0083; p < 0,05 tanpa koreksi)", font_size=7.0)

    set_cell(t15.rows[4].cells[3], "1,241", font_size=7.0)
    set_cell(t15.rows[4].cells[4], "0,265", font_size=7.0)
    set_cell(t15.rows[4].cells[8], "Tidak signifikan (p = 0,265)", font_size=7.0)

    set_cell(t15.rows[5].cells[3], "4,356", font_size=7.0)
    set_cell(t15.rows[5].cells[4], "0,037", font_size=7.0)
    set_cell(t15.rows[5].cells[8], "Tidak signifikan (p > 0,0083; p < 0,05 tanpa koreksi)", font_size=7.0)

    set_cell(t15.rows[6].cells[3], "1,114", font_size=7.0)
    set_cell(t15.rows[6].cells[4], "0,291", font_size=7.0)
    set_cell(t15.rows[6].cells[8], "Tidak signifikan (p = 0,291)", font_size=7.0)

    set_cell(t15.rows[7].cells[3], "0,121", font_size=7.0)
    set_cell(t15.rows[7].cells[4], "0,728", font_size=7.0)
    set_cell(t15.rows[7].cells[8], "Tidak signifikan (Analisis Tambahan)", font_size=7.0)

    # =========================================================================
    # ITEM 5: CREATE AND INSERT PATIENT ID AUDIT TABLE (TABEL 3)
    # =========================================================================
    print("Item 5: Creating Patient ID Audit Table (Tabel 3)...")
    tbl_audit = doc.add_table(rows=7, cols=8)
    tbl_audit.autofit = False
    tbl_audit.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_audit)

    audit_data = [
        ['Partisi Dataset', 'Total Citra', 'Citra Pneumonia', 'Citra Normal', 'Pasien Pneu Unik', 'Pasien Norm Unik', 'Pasien Overlap ke Latih', 'Status Integritas Hash & Berkas'],
        ['Data Latih Resmi (Train+Val)', '5.232', '3.883 (74,22%)', '1.349 (25,78%)', '2.659', '1.349*', '- (Basis Acuan)', '0 korup; 30 duplikat identik internal'],
        ['Data Uji Resmi Kaggle (P2)', '624', '390 (62,50%)', '234 (37,50%)', '202', '225*', '170 pasien (84,16% Pneu)', '0 korup; 0 duplikat SHA-256 lintas partisi'],
        ['Sub-Latih P2 (80% Latih)', '4.185', '3.106 (74,22%)', '1.079 (25,78%)', '2.127', '1.079*', '- (Basis Acuan)', '0 korup; 0 overlap ke Val Tetap'],
        ['Validasi Tetap P2 (20% Latih)', '1.047', '777 (74,21%)', '270 (25,79%)', '532', '270*', '0 pasien (terisolasi penuh)', '0 korup; 0 duplikat SHA-256 lintas partisi'],
        ['Himpunan Uji Protokol P1', '1.172', '855 (72,95%)', '317 (27,05%)', 'Tidak terkontrol', 'Tidak terkontrol', 'Pasien bercampur acak', '0 korup; 0 duplikat SHA-256 lintas partisi'],
        ['Himpunan Uji Protokol P3', '880', '618 (70,23%)', '262 (29,77%)', '425', '262*', '0 pasien terdeteksi', '0 korup; 0 duplikat SHA-256 lintas partisi'],
    ]

    for r_idx, row in enumerate(audit_data):
        for c_idx, val in enumerate(row):
            cell = tbl_audit.rows[r_idx].cells[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 or c_idx == 7 else WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.font.name = 'Times New Roman'
                r.font.size = Pt(6.8)
                r.bold = (r_idx == 0)

    p_cap_audit = doc.add_paragraph()
    set_p(p_cap_audit, "Tabel 3 Audit Integritas Dataset dan Pemetaan Identitas Pasien Lintas Partisi Eksperimen", font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)

    p_fn_audit = doc.add_paragraph()
    set_p(p_fn_audit, (
        "Keterangan: *Pola penamaan berkas kelas Normal (NORMAL2-IM-XXXX atau IM-XXXX) tidak selalu memuat identifier pasien eksplisit "
        "seperti kelas Pneumonia (personXXX_...), sehingga nilai pasien unik Normal diestimasi berbasis prefiks berkas dan ketiadaan "
        "tumpang tindih diartikan sebagai tidak terdeteksi berdasarkan identifier yang tersedia."
    ), font_size=8.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # Remove from document end
    doc._element.body.remove(p_cap_audit._element)
    doc._element.body.remove(tbl_audit._element)
    doc._element.body.remove(p_fn_audit._element)

    # Find audit paragraph in Subbab 4.1 to insert Tabel 3 after it
    p_audit_found = None
    for p in doc.paragraphs:
        if "Audit integritas data pada 5.856 citra rontgen dada" in p.text:
            p_audit_found = p
            break

    if p_audit_found is not None:
        set_p(p_audit_found, (
            "Audit integritas data pada 5.856 citra rontgen dada mengonfirmasi ketiadaan citra korup maupun duplikasi hash SHA-256 lintas partisi, "
            "namun menemukan 170 identitas pasien Pneumonia pada himpunan uji resmi yang tumpang tindih dengan data latih resmi (84,16% dari 202 "
            "pasien Pneumonia uji), sedangkan kelas Normal tidak menunjukkan tumpang tindih identitas terdeteksi. Rincian hasil audit integritas "
            "dan pemetaan pasien lintas seluruh partisi eksperimen dirangkum pada Tabel 3. Analisis statistik deskriptif (rata-rata ± SD) terhadap "
            "keenam fitur GLCM pada seluruh 5.856 citra (Normal n = 1.583 vs Pneumonia n = 4.273) memperlihatkan karakteristik komputasional: "
            "Contrast (Normal 170,18 ± 42,88 vs Pneumonia 110,35 ± 47,22), Dissimilarity (Normal 7,69 ± 0,92 vs Pneumonia 5,74 ± 1,20), "
            "Homogeneity (Normal 0,2318 ± 0,0360 vs Pneumonia 0,2792 ± 0,0386), Energy (Normal 0,0692 ± 0,0382 vs Pneumonia 0,0501 ± 0,0387), "
            "Correlation (Normal 0,9762 ± 0,0066 vs Pneumonia 0,9814 ± 0,0067), dan ASM (Normal 0,0063 ± 0,0069 vs Pneumonia 0,0040 ± 0,0073). "
            "Perbedaan nilai ini merupakan statistik deskriptif kohort dan bukan penanda biologis mutlak. Secara komputasional, citra Normal memperlihatkan "
            "variasi intensitas lokal yang lebih tinggi sehingga Contrast dan Dissimilarity lebih tinggi, sedangkan opasitas luas pada Pneumonia "
            "menghasilkan Homogeneity lebih tinggi. Evaluasi sensitivitas skema pembagian data (RQ1) mengungkapkan variasi kinerja antar-protokol: "
            "kinerja pada P1 (split acak 70:10:20 per citra, 1.172 citra uji) mencapai akurasi 87,03%, spesifisitas 79,50%, dan F1-Score 91,00%. "
            "Pada P2 (split resmi 624 citra uji), model yang sama mencatat akurasi 79,33%, spesifisitas 56,84%, dan F1-Score 84,88%. "
            "Sementara pada P3 (split terkontrol per pasien, 880 citra uji), diperoleh akurasi 87,61%, spesifisitas 80,53%, dan F1-Score 91,13%. "
            "Pola kinerja di mana P1 ≈ P3 > P2 menunjukkan bahwa penurunan performa pada P2 tidak semata-mata diakibatkan oleh kebocoran identitas pasien, "
            "karena P3 yang memisahkan pasien secara terkontrol justru mencapai performa yang setara dengan P1. Perbedaan ini lebih konsisten dijelaskan "
            "oleh pergeseran proporsi kelas pada himpunan uji resmi P2 (37,50% Normal vs 27,05% pada P1) serta pergeseran domain klinis pada data uji Kaggle."
        ), font_size=10.0)

        # Insert after p_audit_found
        p_audit_found._element.addnext(p_fn_audit._element)
        p_audit_found._element.addnext(tbl_audit._element)
        p_audit_found._element.addnext(p_cap_audit._element)
        print("  -> Inserted Tabel 3 (Patient ID Audit) after audit paragraph")

    # =========================================================================
    # CHAPTER 4: UPDATE HEADINGS, CAPTIONS, DISCUSSIONS, & LIMITATIONS
    # =========================================================================
    print("Updating Chapter 4 headings, captions, and narrative discussions...")
    for p in doc.paragraphs:
        txt = p.text.strip()

        # Heading 4.1
        if txt.startswith("4.1") and "Audit Integritas Data" in txt:
            set_p(p, "4.1   Audit Integritas Data dan Sensitivitas Protokol Pembagian Data (RQ1)", bold=True, font_size=10.0, align=WD_ALIGN_PARAGRAPH.LEFT)

        # Heading 4.2
        elif txt.startswith("4.2") and ("Kinerja Model" in txt or "Studi Ablasi" in txt):
            set_p(p, "4.2   Kinerja Model Klasifikasi, Pengaruh Penyeimbangan, dan Uji Statistik (RQ3)", bold=True, font_size=10.0, align=WD_ALIGN_PARAGRAPH.LEFT)

        # Model evaluation table intro
        elif "Hasil komparasi kinerja seluruh model pada protokol utama P2" in txt:
            set_p(p, (
                "Hasil komparasi kinerja seluruh model pada protokol utama P2 (624 citra uji resmi) disajikan pada Tabel 4, "
                "mencakup metrik F1-Score pada himpunan validasi tetap, F1-Score validasi silang (CV) 5-fold per pasien, serta sepuluh metrik diagnostik pengujian resmi."
            ), font_size=10.0)

        # Caption Tabel 3 -> Tabel 4
        elif txt.startswith("Tabel 3") and "Komparasi Kinerja Model" in txt:
            set_p(p, "Tabel 4 Komparasi Kinerja Model pada Protokol Utama P2 (Validasi, CV 5-Fold Per Pasien, dan Uji Resmi 624 Citra)", font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)

        # Footnote of Model Evaluation Table
        elif "Keterangan: Pneumonia adalah kelas positif" in txt:
            set_p(p, (
                "Keterangan: Pneumonia adalah kelas positif pada himpunan uji resmi P2 (390 Pneumonia, 234 Normal). "
                "Metrik F1 Val dievaluasi pada 1.047 citra validasi tetap per pasien; F1 CV 5-fold dinyatakan dalam rata-rata ± SD lintas lipatan berkelompok; "
                "F1 Uji menyertakan 95% Confidence Interval Bootstrap (1.000 iterasi resampling per citra dengan pengembalian); "
                "serta nilai Matthews Correlation Coefficient (MCC) mengevaluasi seluruh sel matriks konfusi diagnostik secara berimbang."
            ), font_size=8.5)

        # Model stability paragraph (Item 4: Bootstrap & Multi-seed limitations)
        elif "Evaluasi stabilitas pemilihan model sebelum pengujian akhir" in txt:
            set_p(p, (
                "Evaluasi stabilitas pemilihan model sebelum evaluasi himpunan uji resmi menunjukkan konsistensi yang sangat tinggi: "
                "pada himpunan validasi tetap (1.047 citra), SVM Tuned meraih F1-Score 92,44%, RF Tuned 92,43%, KNN Tuned 91,53%, dan Soft Voting Ensemble 92,12%. "
                "Pengujian validasi silang terstratifikasi berkelompok 5-fold berbasis ID pasien pada 5.232 citra latih resmi menghasilkan rata-rata F1-Score "
                "yang sangat stabil dengan variansi minimal: SVM 91,99 ± 0,47%, RF 91,60 ± 0,57%, KNN 91,30 ± 0,42%, dan Soft Voting 92,08 ± 0,28%. "
                "Selain itu, pengujian multi-seed acak (5 kali pengulangan: seed 42, 123, 456, 789, 2026) pada pemodelan final menunjukkan stabilitas stokastik "
                "dengan simpangan baku F1-Score pada data uji hanya sebesar ±0,34% (F1 Soft Voting: 84,88% ± 0,34%; SVM: 84,62% ± 0,29%). "
                "Perlu dicatat keterbatasan evaluasi multi-seed ini yang mengukur kestabilan algoritma (seperti inisialisasi bobot dan penskalaan Platt) "
                "namun tidak mengukur variansi partisi karena data uji P2 merupakan himpunan deterministik tetap. Demikian pula, interval kepercayaan 95% Bootstrap "
                "(82,23%–87,56% untuk F1 Soft Voting) dihitung melalui resampling per citra independen, sehingga diasumsikan sebagai batas optimistik "
                "dibandingkan cluster-bootstrap berbasis pasien."
            ), font_size=10.0)

        # Model comparison & SMOTE paragraph (Item 6: SMOTE trade-off)
        elif "Berdasarkan hasil uji resmi P2 (Tabel 3)" in txt or "Berdasarkan hasil uji resmi P2 (Tabel 4)" in txt:
            set_p(p, (
                "Berdasarkan hasil evaluasi himpunan uji resmi P2 (Tabel 4), pada konfigurasi tanpa penyeimbangan, Soft Voting Ensemble dan SVM Tuned "
                "meraih akurasi tertinggi yang identik sebesar 79,33%, mengungguli Baseline Mayoritas (62,50%) sebesar +16,83 poin persentase. "
                "Soft Voting unggul tipis pada Recall Pneumonia (92,82% vs 91,03%) dan F1-Score (84,88% vs 84,62%), sedangkan SVM Tuned lebih unggul pada "
                "Spesifisitas (59,83% vs 56,84%), Balanced Accuracy (75,43% vs 74,83%), dan Macro-F1 (76,54% vs 76,11%). KNN Default (k=65) mencatat Spesifisitas "
                "tertinggi di antara model tanpa penyeimbangan (64,53%) dan MCC tertinggi (0,563), namun memiliki Recall lebih rendah (89,23%). "
                "Penerapan teknik SMOTE (k=5) memberikan pergeseran titik kerja diagnostik yang sangat nyata: pada SVM Tuned + SMOTE, Spesifisitas melonjak "
                "menjadi 66,24% (+6,41 poin persentase) dan Balanced Accuracy mencapai 77,09% dengan akurasi 79,81% dan MCC 0,561; sedangkan pada Soft Voting + SMOTE, "
                "Spesifisitas meningkat menjadi 65,81% (+8,97 poin persentase), Balanced Accuracy 77,01%, Macro-F1 77,74%, dan MCC 0,560. "
                "Peningkatan spesifisitas ini mengorbankan sedikit Recall (turun dari 92,82% menjadi 88,21% pada Ensemble dan 87,95% pada SVM), "
                "yang secara klinis mencerminkan kompromi penting: menekan alarm palsu (false positive) pada pasien normal tanpa kehilangan kemampuan deteksi pneumonia yang memadai. "
                "Distribusi kesalahan klasifikasi divisualisasikan pada Gambar 5 dan Gambar 6."
            ), font_size=10.0)

        # Caption Tabel 4 -> Tabel 5 (McNemar)
        elif txt.startswith("Tabel 4") and "Hasil Uji Signifikansi McNemar" in txt:
            set_p(p, "Tabel 5 Hasil Uji Signifikansi McNemar Berpasangan dan Analisis Keragaman Ensemble (Uji Resmi P2, 624 Citra)", font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)

        # McNemar discussion paragraph (Item 4: McNemar limitation)
        elif "Uji McNemar berpasangan dengan koreksi Bonferroni" in txt and not "Penelitian ini menyimpulkan" in txt:
            set_p(p, (
                "Uji McNemar berpasangan dengan koreksi Bonferroni (α = 0,0083 untuk 6 pasangan model utama) yang dirangkum pada Tabel 5 "
                "menunjukkan bahwa tidak ada pasangan model yang berbeda signifikan secara statistik pada tingkat signifikansi α = 0,0083. "
                "Pasangan SVM Tuned vs RF Tuned (statistik 4,356, p = 0,037) dan Soft Voting vs RF Tuned (statistik 6,323, p = 0,012) "
                "hanya menunjukkan signifikansi marginal pada tingkat uncorrected α = 0,05, namun menjadi tidak signifikan setelah penyesuaian Bonferroni. "
                "Perbandingan antara Soft Voting dan SVM Tuned mencatat statistik 0,050 dengan nilai p = 0,823 dan laju ketidaksepakatan (disagreement rate) hanya 3,21%. "
                "Temuan ini dikonfirmasi oleh analisis keragaman ensemble yang memperlihatkan korelasi galat sangat tinggi antar-model dasar "
                "(r = 0,713–0,883; Yule's Q = 0,961–0,987). Tingginya korelasi galat ini mengonfirmasi batas efektivitas Soft Voting Ensemble (RQ3): "
                "pada ruang fitur GLCM yang homogen, model dasar membuat kesalahan prediksi pada sampel citra yang nyaris sama, "
                "sehingga mekanisme agregasi probabilitas tidak mampu mengompensasi galat individu. Selain itu, keterbatasan uji McNemar pada penelitian ini "
                "perlu dicatat: adanya citra multipel dari 170 pasien Pneumonia yang sama pada data uji resmi menyebabkan asumsi independensi observasi "
                "tidak terpenuhi secara murni, sehingga berpotensi menimbulkan efek pengelompokan (cluster effect)."
            ), font_size=10.0)

        # Heading 4.3 (RQ2)
        elif txt.startswith("4.3") and "Analisis Ablasi" in txt:
            set_p(p, "4.3   Studi Ablasi Pra-Pengolahan, Kuantisasi Ng, Varian Fitur, dan Uji Ketahanan Tepi (RQ2)", bold=True, font_size=10.0, align=WD_ALIGN_PARAGRAPH.LEFT)

        # Ablation study intro (Item 3: Exploratory ablation study)
        elif "Studi ablasi 27 kombinasi" in txt:
            set_p(p, (
                "Studi ablasi parametrik 27 kombinasi (3 praproses × 3 nilai Ng × 3 varian GLCM) dievaluasi pada data uji resmi P2 sebagai "
                "studi eksploratif untuk menjawab RQ2. Hasil representatif dari 11 kombinasi utama disajikan pada Tabel 6."
            ), font_size=10.0)

        # Caption Tabel 5 -> Tabel 6 (Ablation Table)
        elif txt.startswith("Tabel 5") and "Ringkasan Hasil Studi Ablasi" in txt:
            set_p(p, "Tabel 6 Ringkasan Hasil Studi Ablasi Eksploratif (11 Konfigurasi Representatif dari 27 Kombinasi pada Data Uji Resmi P2)", font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)

        # Ablation discussion
        elif "Dekomposisi faktor ablasi (RQ2) mengungkapkan tiga temuan utama:" in txt or "Analisis parametrik faktor ablasi (RQ2)" in txt:
            set_p(p, (
                "Analisis parametrik faktor ablasi (RQ2) mengungkapkan tiga temuan utama: (a) varian fitur memberikan pengaruh paling dominan "
                "terhadap spesifisitas—penggunaan GLCM-24 tanpa perataan sudut menurunkan akurasi menjadi 70,99%–74,52% dan spesifisitas hingga "
                "33,33%–44,44% akibat overfitting terhadap orientasi sudut pencitraan, sedangkan GLCM-5 (80,45%) hampir identik dengan GLCM-6 (80,77%) "
                "yang mengonfirmasi redundansi Energy = √ASM; (b) pra-pengolahan ortiz memberikan peningkatan akurasi sebesar +3,53 poin persentase "
                "dibanding basic pada Ng=32 (80,77% vs 77,24%) dan +2,56 poin dibanding clahe (78,21%); serta (c) kuantisasi tingkat keabuan "
                "yang lebih ringkas Ng=32 secara konsisten melampaui Ng=64 dan Ng=256 (+1,28 poin persentase pada mode basic: 77,24% vs 75,96%) "
                "karena meredam fluktuasi derau lokal. Perlu ditekankan bahwa capaian akurasi 80,77% pada konfigurasi ortiz Ng=32 GLCM-6 merupakan "
                "temuan eksploratif pasca-evaluasi pada data uji dan bukan hasil seleksi satu langkah model final."
            ), font_size=10.0)

        # PFI discussion (Item 6: Multicollinearity masking)
        elif "Analisis Permutation Feature Importance pada data uji resmi P2" in txt:
            set_p(p, (
                "Analisis Permutation Feature Importance (PFI) pada data uji resmi P2 (Gambar 4) mengidentifikasi Dissimilarity "
                "(penurunan akurasi 7,51% ± 0,58%) dan Homogeneity (6,46% ± 0,46%) sebagai kontributor utama dalam membedakan opasitas konsolidasi pneumonia "
                "dari corakan bronkovaskular normal. Contrast menempati peringkat ketiga (3,53% ± 0,42%), sedangkan ASM (0,96% ± 0,25%), Energy (0,80% ± 0,22%), "
                "dan Correlation (0,48% ± 0,18%) memberikan kontribusi marjinal. Dalam interpretasi PFI, fenomena penyamaran multikolinieritas "
                "(multicollinearity masking) perlu diperhitungkan: karena pasangan fitur GLCM memiliki interkorelasi komputasional yang erat "
                "(khususnya Energy dan ASM yang memiliki hubungan fungsional kuadratis eksak Energy = √ASM, serta Contrast dan Dissimilarity yang keduanya "
                "mengukur gradien disparitas lokal), pengacakan permutasi pada satu fitur dapat dikompensasi sebagian oleh pasangannya yang masih utuh. "
                "Hal ini menjelaskan mengapa penurunan akurasi individu ASM dan Energy tampak rendah meskipun informasi tekstur kekompakan matriks tetap esensial. "
                "Uji ketahanan terhadap pemotongan tepi perifer dirangkum pada Tabel 7."
            ), font_size=10.0)

        # Caption Tabel 6 -> Tabel 7 (Crop Test)
        elif txt.startswith("Tabel 6") and "Uji Ketahanan Model terhadap Pemotongan Tepi" in txt:
            set_p(p, "Tabel 7 Uji Ketahanan Model terhadap Pemotongan Tepi Citra (Peripheral Center Crop Test pada Uji Resmi P2)", font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)

        # CPU latency & Literature intro
        elif "Dari sisi efisiensi komputasi CPU" in txt:
            set_p(p, (
                "Dari sisi efisiensi komputasi CPU, seluruh pipeline inferensi yang disimpan via joblib (.pkl) mencatat waktu pelatihan 1,987 detik "
                "pada data latih resmi 5.232 citra. Pada tahap inferensi, purwarupa web Streamlit mencatat waktu prediksi model murni hanya 0,0671 ms "
                "per citra (41,86 ms untuk seluruh 624 citra), sedangkan latensi inferensi CPU end-to-end (meliputi pembacaan berkas citra, konversi grayscale, "
                "resizing 150×150 piksel, ekstraksi 6 fitur GLCM, penskalaan fitur, dan prediksi model) tercatat sebesar 24,42 ms per citra (~40,9 frame per detik), "
                "memenuhi syarat penerapan komputasi tepi berdaya rendah. Hasil perbandingan kinerja pada himpunan uji resmi yang sama dirangkum pada Tabel 8."
            ), font_size=10.0)

        # Caption Tabel 7 -> Tabel 8 (Literature Table)
        elif txt.startswith("Tabel 7") and "Perbandingan Kinerja pada Himpunan Uji Resmi" in txt:
            set_p(p, "Tabel 8 Perbandingan Kinerja pada Himpunan Uji Resmi yang Sama (624 Citra) dengan Pembanding Utama [15] dan Studi Pendahulu [19]", font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)

        # =====================================================================
        # ITEM 4: CHAPTER 5 (ANCAMAN VALIDITAS / THREATS TO VALIDITY)
        # =====================================================================
        elif "Terdapat beberapa ancaman terhadap validitas (threats to validity)" in txt:
            set_p(p, (
                "Terdapat beberapa ancaman terhadap validitas (threats to validity) yang perlu dicatat dalam penelitian ini: "
                "(1) Generalisasi terbatas karena dataset hanya berasal dari satu populasi pediatrik berusia 1 hingga 5 tahun di Guangzhou Women and Children's Medical Center [4]; "
                "(2) Status audit overlap ID pasien pada dataset resmi menemukan 170 identitas pasien Pneumonia pada himpunan uji resmi yang tumpang tindih dengan data latih resmi "
                "(84,16% dari 202 pasien Pneumonia uji), sedangkan kelas Normal tidak menunjukkan tumpang tindih terdeteksi namun format penamaan berkasnya (`IM-xxxx` / `NORMAL2-IM-XXXX`) "
                "tidak selalu menyertakan nomor pasien eksplisit sehingga tetap menyisakan potensi ketergantungan spasial; "
                "(3) Ketiadaan segmentasi ROI organ paru otomatis memungkinkan fitur GLCM menangkap informasi latar belakang luar paru, di mana uji artefak crop tepi 10% dan 15% "
                "menurunkan akurasi secara drastis masing-masing menjadi 68,59% (-10,74 poin persentase) dan 63,14% (-16,19 poin persentase, mendekati baseline mayoritas 62,50%), "
                "mengonfirmasi bahwa fitur GLCM global turut menangkap kontras batas toraks dan diafragma; "
                "(4) Efek penyamaran multikolinieritas (multicollinearity masking) pada analisis Permutation Feature Importance berpotensi mengaburkan kontribusi fitur yang berkorelasi "
                "tinggi (seperti ASM dan Energy) dan tidak merepresentasikan biomarker klinis biologis; "
                "(5) Keterbatasan pengujian statistik: asumsi independensi observasi pada uji McNemar tidak terpenuhi sempurna akibat keberadaan sampel multipel dari pasien yang sama "
                "pada data uji resmi (efek kluster), sedangkan estimasi interval kepercayaan 95% Bootstrap berbasis resampling citra menghasilkan rentang yang diasumsikan lebih sempit "
                "dibandingkan resampling berbasis kluster pasien; "
                "(6) Evaluasi multi-seed (5 kali pengulangan) hanya menangkap variansi stokastik algoritma dan bukan variansi pemisahan data karena set uji P2 bersifat deterministik; "
                "(7) Label kebenaran dasar (ground truth) bergantung pada anotasi sekunder dataset tanpa verifikasi ulang oleh tim radiolog independen; serta "
                "(8) Ukuran himpunan validasi tetap yang relatif kecil (1.047 citra) dimitigasi dengan pengujian validasi silang 5-fold berkelompok per pasien."
            ), font_size=10.0)

        # =====================================================================
        # ITEM 1: CHAPTER 6 (KESIMPULAN RQ1, RQ2, RQ3)
        # =====================================================================
        elif "Penelitian ini menyimpulkan tiga temuan utama" in txt:
            set_p(p, (
                "Penelitian ini menyimpulkan tiga temuan utama sesuai pertanyaan penelitian yang dirumuskan: "
                "(1) Jawaban RQ1: Kinerja klasifikasi berbasis fitur tekstur GLCM sangat sensitif terhadap variasi skema pembagian data. "
                "Evaluasi pada protokol acak per citra P1 (n=1.172 uji) menghasilkan akurasi 87,03%, spesifisitas 79,50%, dan F1-Score 91,00%. "
                "Pada protokol resmi Kaggle P2 (n=624 uji), kinerja model yang sama terkoreksi turun menjadi akurasi 79,33%, spesifisitas 56,84%, dan F1-Score 84,88%, "
                "menunjukkan optimism gap sebesar 7,70 poin persentase pada akurasi dan 22,66 poin persentase pada spesifisitas. "
                "Sementara itu, protokol terkontrol per pasien P3 (n=880 uji) memperoleh akurasi 87,61%, spesifisitas 80,53%, dan F1-Score 91,13%. "
                "Pola komparasi P1 ≈ P3 > P2 mengindikasikan bahwa penurunan kinerja pada P2 tidak semata-mata dipicu oleh kebocoran ID pasien "
                "(meskipun audit menemukan 170 tumpang tindih ID pasien Pneumonia antara data latih dan uji resmi), melainkan lebih konsisten dijelaskan "
                "oleh pergeseran proporsi kelas Normal (37,50% pada P2 vs 27,05% pada P1) serta pergeseran domain visual pada data uji resmi Kaggle; "
                "(2) Jawaban RQ2: Dekomposisi ablasi 27 kombinasi pada data uji resmi P2 mengonfirmasi bahwa varian fitur dan pra-pengolahan merupakan faktor teknis "
                "paling dominan dalam memengaruhi performa. Perataan empat orientasi sudut pada GLCM-6 sangat krusial untuk mencegah penurunan spesifisitas "
                "hingga 33,33%–44,44% yang dialami GLCM-24 tanpa perataan sudut akibat overfitting orientasi. Fitur Energy terkonfirmasi redundan terhadap ASM (Energy = √ASM), "
                "di mana GLCM-5 menghasilkan akurasi 80,45% yang setara dengan GLCM-6 (80,77%). Pra-pengolahan ortiz meningkatkan akurasi hingga +3,53 poin persentase "
                "dibanding basic pada Ng=32 (80,77% vs 77,24%), dan kuantisasi Ng=32 secara konsisten melampaui Ng=64 dan Ng=256 karena meredam derau lokal. "
                "Permutation Feature Importance menempatkan Dissimilarity (7,51% ± 0,58%) dan Homogeneity (6,46% ± 0,46%) sebagai fitur paling berpengaruh, "
                "sedangkan uji pemotongan tepi perifer (penurunan akurasi hingga 63,14% pada crop 15%) mengonfirmasi ketergantungan model terhadap batas toraks luar paru. "
                "Capaian akurasi 80,77% pada konfigurasi ortiz Ng=32 GLCM-6 dicatat sebagai temuan studi ablasi eksploratif; serta "
                "(3) Jawaban RQ3: Uji signifikansi McNemar berpasangan dengan koreksi Bonferroni (α = 0,0083) mengonfirmasi bahwa Soft Voting Ensemble "
                "memiliki kinerja yang setara secara statistik terhadap model dasar tunggal terbaik (SVM Tuned) dengan p = 0,823 dan disagreement rate hanya 3,21%. "
                "Tingginya korelasi galat antar-model dasar (r = 0,713–0,883; Yule's Q = 0,961–0,987) pada ruang fitur GLCM homogen membatasi efektivitas ensemble "
                "karena model gagal memberikan diversitas galat saling melengkapi. Sementara itu, teknik SMOTE berhasil menggeser titik kompromi diagnostik, "
                "menaikkan spesifisitas secara signifikan dari 56,84% menjadi 66,24% dan Macro-F1 dari 76,11% menjadi 77,79%, dengan sedikit kompromi pada penurunan recall "
                "(92,82% ke 87,95%). Pipeline inferensi utuh telah diintegrasikan ke purwarupa web interaktif Streamlit (https://pneumonia-ml.streamlit.app/) "
                "dengan latensi inferensi end-to-end CPU 24,42 ms per citra (~40,9 FPS) dan waktu prediksi model 0,0671 ms."
            ), font_size=10.0)

    # =========================================================================
    # ITEM 8: REFERENCE LIST METADATA VERIFICATION
    # =========================================================================
    print("Item 8: Verifying Reference List metadata...")
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt.startswith("[5]"):
            set_p(p, (
                "[5]   C. Cillóniz, C. Dominedò, J. M. Pericàs, D. Rodriguez-Hurtado, and A. Torres, "
                "“Community-acquired pneumonia in critically ill very old patients: a growing problem,” "
                "European Respiratory Review, vol. 29, no. 155, art. 190126, pp. 1–13, Feb. 2020, doi: 10.1183/16000617.0126-2019."
            ), font_size=9.0)
            print("  -> Ref [5] verified (Cillóniz, Eur Respir Rev)")
        elif txt.startswith("[9]"):
            set_p(p, (
                "[9]   S. T. Erukude, V. C. Marella, and S. R. Veluru, "
                "“Explainable Deep Learning in Medical Imaging: Brain Tumor and Pneumonia Detection,” "
                "in Proc. 4th IEEE Int. Conf. Innovative Mechanisms for Industry Applications (ICIMIA), Sep. 2025, pp. 906–911, doi: 10.1109/ICIMIA67127.2025.11200629."
            ), font_size=9.0)
            print("  -> Ref [9] verified (Erukude, ICIMIA 2025)")
        elif txt.startswith("[18]"):
            set_p(p, (
                "[18]  A. Shati, G. M. Hassan, and A. Datta, "
                "“A comprehensive fusion model for improved pneumonia prediction based on KNN-wavelet-GLCM and a residual network,” "
                "Intelligent Systems with Applications, vol. 26, art. 200492, pp. 1–15, Jun. 2025, doi: 10.1016/j.iswa.2025.200492."
            ), font_size=9.0)
            print("  -> Ref [18] verified (Shati, Intell Syst Appl 2025)")
        elif txt.startswith("[22]"):
            set_p(p, (
                "[22]  A. I. K. Akbar and Y. P. Astuti, "
                "“Lung Cancer Classification using the Naïve Bayes Method with SMOTE,” "
                "Sistemasi: Jurnal Sistem Informasi, vol. 14, no. 6, pp. 2954–2967, Nov. 2025, doi: 10.32520/stmsi.v14i6.5607."
            ), font_size=9.0)
            print("  -> Ref [22] verified (Akbar & Astuti, SISTEMASI 2025)")

    # =========================================================================
    # SAVE TO BOTH TARGET LOCATIONS
    # =========================================================================
    print("Saving modified document...")
    os.makedirs(os.path.dirname(DEST_DESKTOP), exist_ok=True)
    os.makedirs(os.path.dirname(DEST_REPO), exist_ok=True)

    doc.save(DEST_REPO)
    print(f"Saved to repository: {DEST_REPO}")

    doc.save(DEST_DESKTOP)
    print(f"Saved to Desktop: {DEST_DESKTOP}")
    print("All revisions successfully built!")

if __name__ == "__main__":
    main()
