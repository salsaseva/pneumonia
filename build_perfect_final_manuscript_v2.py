import os
import sys
import shutil
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
import pandas as pd
import zipfile

def set_cell(cell, text, bold=False, font_size=7.0, align=WD_ALIGN_PARAGRAPH.CENTER, color=None):
    cell.text = text
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(1.5)
    p.paragraph_format.space_after = Pt(1.5)
    p.paragraph_format.line_spacing = 1.0
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    if len(p.runs) > 0:
        run = p.runs[0]
        run.font.name = 'Times New Roman'
        run.font.size = Pt(font_size)
        run.font.bold = bold
        if color:
            run.font.color.rgb = color

def set_p(p, text, bold=False, italic=False, font_size=10.0, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=4.0):
    p.text = text
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.05
    for r in p.runs:
        r.font.name = 'Times New Roman'
        r.font.size = Pt(font_size)
        r.font.bold = bold
        r.font.italic = italic

def set_table_borders(table):
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>\n'
            f'  <w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>\n'
            f'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>\n'
            f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>\n'
            f'  <w:insideV w:val="none"/>\n'
            f'  <w:left w:val="none"/>\n'
            f'  <w:right w:val="none"/>\n'
            f'</w:tblBorders>'
        )
        tblPr[0].append(borders)

def build_manuscript():
    base_docx = "/tmp/gdoc_fresh.docx"
    output_docx = "manuscript/Revisi_Jurnal_SISTEMASI_Pneumonia_GLCM.docx"
    desktop_docx = "/Users/dicky/Desktop/pneumonia/Revisi_Jurnal_SISTEMASI_Pneumonia_GLCM.docx"
    
    print(f"Loading fresh base template from: {base_docx}...")
    doc = docx.Document(base_docx)
    
    # 0. Sanitize sectPr twips
    for s in doc.sections:
        for tag in ['w:pgMar', 'w:pgSz']:
            elem = s._sectPr.find(qn(tag))
            if elem is not None:
                for k in list(elem.attrib.keys()):
                    val = elem.attrib[k]
                    try:
                        elem.attrib[k] = str(int(round(float(val))))
                    except (ValueError, TypeError):
                        pass

    # 1. Judul & Penulis
    print("[1/10] Updating Title and Author Metadata...")
    doc.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_p(doc.paragraphs[0], (
        "Sensitivitas Kinerja Fitur Tekstur GLCM terhadap Protokol Pembagian Data "
        "pada Klasifikasi Pneumonia Pediatrik"
    ), bold=True, font_size=14.0, align=WD_ALIGN_PARAGRAPH.CENTER)
    
    set_p(doc.paragraphs[1], (
        "Performance Sensitivity of GLCM Texture Features Across Data Splitting Protocols "
        "in Pediatric Pneumonia Classification"
    ), italic=True, font_size=11.0, align=WD_ALIGN_PARAGRAPH.CENTER)

    # 2. Abstrak Indonesia & Inggris (Strictly <= 250 words)
    print("[2/10] Setting Abstrak ID (247 kata) & EN (249 kata)...")
    abstract_id = (
        "Evaluasi klasifikasi pneumonia berbasis citra rontgen dada (CXR) menggunakan fitur tekstur "
        "Gray-Level Co-occurrence Matrix (GLCM) kerap terkendala sensitivitas protokol partisi data dan optimisme semu. "
        "Penelitian ini mengevaluasi sensitivitas enam fitur GLCM lintas tiga protokol: acak per citra "
        "(P1; n=1.172 uji), resmi Kaggle dengan validasi per pasien (P2; 5.232 latih / 624 uji), dan kontrol pasien "
        "murni (P3; n=880 uji) pada 5.856 citra pediatrik. Audit dataset mendeteksi 170 dari 202 pasien Pneumonia "
        "data uji resmi tumpang tindih dengan data latih (84,16%), sedangkan kelas Normal teridentifikasi bebas tumpang tindih. "
        "Seluruh eksperimen menerapkan pipeline terisolasi (StandardScaler fit di data latih) guna mencegah kebocoran data. "
        "Pada uji resmi P2, Soft Voting Ensemble dan Support Vector Machine (SVM) Tuned meraih akurasi 79,33% "
        "(F1 84,88% vs 84,62%, spesifisitas 56,84% vs 59,83%, MCC 0,549 vs 0,548). Penyeimbangan kelas "
        "class_weight='balanced' pada SVM Tuned menaikkan spesifisitas ke 68,80% (akurasi 80,77%, MCC 0,583), "
        "lebih efisien dibanding SMOTE tanpa memicu artefak sintetis. Uji McNemar (α = 0,0083) dan ekivalensi TOST "
        "(p = 0,0026, margin ±2%) mengonfirmasi tidak ada perbedaan performa antara Soft Voting dan SVM akibat tingginya "
        "korelasi galat (r = 0,713–0,883). Subsampling 100× menyelaraskan akurasi P1 ke 81,73% dan P3 ke 82,40%, "
        "mengonfirmasi gap P2 dipicu oleh prevalensi dan pergeseran domain visual akuisisi. "
        "Ablasi 27 kombinasi menunjukkan GLCM-6 mencegah kolaps spesifisitas (34%–36%) akibat overfitting dimensi pada GLCM-24. "
        "Permutasi berkelompok (Grouped PFI) menempatkan kelompok variasi kontras sebagai prediktor utama "
        "(penurunan F1 18,16%), sedangkan masking tepi tanpa resize (akurasi turun ke 62,50%) mengonfirmasi "
        "ketergantungan model pada batas toraks luar."
    )

    abstract_en = (
        "Evaluating chest X-ray (CXR) pneumonia classification using Gray-Level Co-occurrence Matrix "
        "(GLCM) texture features is frequently confounded by protocol sensitivity and optimistic bias. "
        "This study evaluates six GLCM features across three protocols: random image partition (P1; n=1,172 test), "
        "official Kaggle split with patient validation (P2; 5,232 train / 624 test), and patient-controlled "
        "split (P3; n=880 test) across 5,856 pediatric radiographs. Dataset auditing identified 170 of 202 Pneumonia "
        "test patients overlapping with training data (84.16%), whereas Normal test patients were identified without overlap. "
        "All workflows enforce strict isolated pipelines (StandardScaler fitted exclusively on training folds) to prevent leakage. "
        "Under protocol P2, Soft Voting Ensemble and Tuned Support Vector Machine (SVM) achieved identical accuracy of 79.33% "
        "(F1 84.88% vs 84.62%, specificity 56.84% vs 59.83%, MCC 0.549 vs 0.548). Cost-sensitive class_weight='balanced' "
        "in Tuned SVM elevated specificity to 68.80% (accuracy 80.77%, MCC 0.583), outperforming SMOTE without synthetic artifacts. "
        "Paired McNemar testing (α = 0.0083) and TOST equivalence (p = 0.0026, margin ±2%) confirmed no statistically "
        "significant difference between Soft Voting and Tuned SVM due to substantial error correlation (r = 0.713–0.883). "
        "Equalizing prevalence via 100× subsampling adjusted P1 accuracy to 81.73% and P3 to 82.40%, confirming that the P2 "
        "gap reflects prevalence disparity and visual domain shift. Ablation across 27 configurations revealed "
        "that GLCM-6 prevents severe specificity collapse (34%–36%) seen in un-averaged GLCM-24 due to dimensionality overfitting. "
        "Grouped Permutation Feature Importance identified local contrast as the primary predictor (18.16% F1 drop), "
        "while non-resized edge masking (accuracy drop to 62.50%) confirmed reliance on external thoracic boundaries."
    )

    for p in doc.paragraphs[:15]:
        if "Evaluasi model klasifikasi pneumonia berbasis citra" in p.text:
            set_p(p, abstract_id, font_size=9.0)
        elif "Evaluating chest X-ray" in p.text:
            set_p(p, abstract_en, font_size=9.0)

    # 3. Bab 1 Pendahuluan, Bab 2 Tinjauan Literatur, & Tabel 1
    print("[3/10] Updating Bab 1 Pendahuluan, Bab 2 Tinjauan Literatur, and Tabel 1...")
    for p in doc.paragraphs[13:26]:
        if "Meskipun dataset Kermany et al. [4] telah digunakan secara luas" in p.text:
            set_p(p, (
                "Meskipun dataset Kermany et al. [4] telah digunakan secara luas, telaah literatur sistematis kami menunjukkan "
                "kelemahan metodologis yang berulang: sebagian besar penelitian menerapkan pembagian data acak per citra (image-level random split) "
                "tanpa mempertimbangkan identitas pasien. Padahal, audit kami menemukan kecocokan ID pasien yang mengindikasikan bahwa beberapa citra "
                "berasal dari pasien yang sama yang tersebar ke data latih dan data uji, memicu kebocoran data (data leakage) dan estimasi performa semu. "
                "Fenomena ini diperparah oleh kerentanan model citra toraks dalam menangkap fitur pengganggu (confounding features) dari anotasi tepi "
                "sebagaimana diungkapkan oleh Zech et al. [22]. Di sisi klinis, penegakan diagnosis pneumonia pediatrik bertumpu pada pemeriksaan "
                "radiografi rontgen dada (Chest X-Ray / CXR) anteroposterior guna mengidentifikasi konsolidasi maupun infiltrat paru berdasarkan "
                "pedoman standar interpretasi radiologis World Health Organization (WHO) oleh Cherian et al. [5]. Upaya otomatisasi skrining berbasis "
                "kecerdasan artifisial berkembang pesat [6], [7], [8], [9]. Berdasarkan penelusuran literatur pada basis data PubMed, IEEE Xplore, "
                "dan Scopus menggunakan string penelusuran ((\"pneumonia\" AND \"GLCM\" AND \"chest X-ray\") AND (\"split\" OR \"leakage\" OR \"ensemble\")), "
                "sepengetahuan penulis, penelitian komparatif yang secara sistematis mengevaluasi sensitivitas model GLCM tradisional "
                "terhadap variasi protokol pembagian data, membandingkannya dengan pembagian berbasis pasien, serta menguji efektivitas ensemble "
                "pada fitur tekstur homogen masih sangat terbatas."
            ), font_size=10.0)
        elif "Dalam pemodelan klasifikasi berbasis fitur radiomik" in p.text:
            set_p(p, (
                "Dalam pemodelan klasifikasi berbasis fitur radiomik dan tekstur pada citra rontgen dada, Ortiz-Toro et al. [15] berfungsi sebagai "
                "pembanding utama yang mengevaluasi fitur radiomics (statistik orde pertama dan tekstur radiomik), dimensi fraktal, serta histon "
                "berbasis superpiksel dengan pra-pengolahan kontras adaptif tanpa penyeimbangan versus SMOTE pada data uji resmi (390 Pneumonia dan 234 Normal), "
                "di mana Random Forest menghasilkan akurasi 82,21% (recall 89,74%, spesifisitas 69,66%, F1 86,10%, balanced accuracy 79,70%), "
                "SVM linear menghasilkan akurasi 78,53% (recall 84,36%, spesifisitas 68,80%, F1 83,02%), dan KNN memperoleh akurasi 79,33%. "
                "Sementara itu, studi pendahulu oleh Hismaya dan Sulistianingsih [19] mengevaluasi GLCM 24 fitur (4 arah sudut tanpa perataan) "
                "pada data uji resmi 624 citra, di mana KNN memperoleh akurasi 73,72% (spesifisitas 47,01%), SVM awal memperoleh akurasi 76,12% "
                "((138+337)/624 citra; recall 86,41%, spesifisitas 58,97%, F1 81,80%), dan SVM tuned mengalami penurunan akurasi pada data uji menjadi 72,76% "
                "(spesifisitas 35,04%) akibat overfitting dimensionalitas. Selain itu, rujukan mengenai pemodelan Machine Learning dan CNN pada pencitraan toraks "
                "juga dilaporkan pada [14], [16], [17], [18], [20], [21]. Ringkasan perbandingan protokol dari berbagai studi terdahulu disajikan pada Tabel 1."
            ), font_size=10.0)
        elif "Catatan: Kolom yang belum terisi pada referensi tertentu" in p.text or "Catatan: Singkatan n.r." in p.text:
            set_p(p, "Catatan: Singkatan n.r. (not reported) menandakan aspek protokol tidak dilaporkan secara eksplisit pada publikasi sumber langsung.", italic=True, font_size=8.0)

    # Update Tabel 1 cells
    t1 = doc.tables[0]
    t1_headers = ['Studi', 'Fitur / Ekstraksi', 'Protokol Split', 'Kontrol Pasien', 'Tuning', 'Penanganan Imbalance', 'Uji Statistik', 'Akurasi Uji (%)']
    for c_idx, h in enumerate(t1_headers):
        set_cell(t1.rows[0].cells[c_idx], h, bold=True, font_size=7.2)
    
    t1_rows = [
        ['Ortiz-Toro et al. (2022) [15]', 'Radiomics, Fractal, Histon', 'Resmi (5.232 / 624)', 'Tidak Dieksplisitkan', 'GridSearch', 'Tidak', 'Wilcoxon', '82,21% (RF Resmi)'],
        ['Hismaya & Sulistianingsih (2025) [19]', 'GLCM (24 Kolom)', 'Resmi (5.216 / 624)', 'Tidak', 'GridSearchCV', 'Tidak', 'n.r.', '76,12% (SVM Resmi)'],
        ['Farghaly & Deshpande (2024) [12]', 'GLCM + GLDM', 'Split Acak', 'Tidak Dieksplisitkan', 'Sebagian', 'Tidak', 'n.r.', '96,20% (Klaim)'],
        ['Kim (2022) [14]', 'Fitur Radiomik', 'Split Acak / CV', 'Tidak Dieksplisitkan', 'Ya', 'Tidak Dieksplisitkan', 'Uji Korelasi', '85,00% (Klaim)'],
        ['Zaena & Lhaksmana (2025) [16]', 'Fitur Citra / RF vs CNN', 'Split Acak', 'Tidak', 'Tidak Dieksplisitkan', 'Tidak', 'n.r.', '91,00% (Klaim)'],
        ['Shati et al. (2025) [18]', 'KNN-Wavelet-GLCM', 'Split Acak / CV', 'Tidak Dieksplisitkan', 'Ya', 'Tidak Dieksplisitkan', 'n.r.', '98,46% (Klaim)'],
        ['Mustopa et al. (2022) [21]', 'GLCM + Hu Moments', 'Split Acak', 'Tidak', 'Tidak Dieksplisitkan', 'SMOTE', 'n.r.', '90,00% (Klaim)'],
        ['Arifin & Surya (2022) [20]', 'CNN (VGG16)', 'Split Bawaan', 'Tidak', 'Tidak Dieksplisitkan', 'Tidak', 'n.r.', '88,00% (Resmi)'],
        ['Penelitian Ini', 'GLCM-6 / 24 / 5', 'P1 (70:10:20), P2 (Resmi 5.232/624), P3 (Ketat Pasien)', 'Audit ID SHA-256 (3.118 Pasien)', 'GridSearchCV 5-Fold Grouped', 'class_weight & SMOTE', 'McNemar + TOST', '79,33% (80,77% cw)']
    ]
    for r_idx, r_data in enumerate(t1_rows, start=1):
        for c_idx, val in enumerate(r_data):
            set_cell(t1.rows[r_idx].cells[c_idx], val, bold=(r_idx==len(t1_rows)), font_size=7.0, align=WD_ALIGN_PARAGRAPH.CENTER)

    # 4. Bab 3 Metode Penelitian & Gambar 1 Flowchart
    print("[4/10] Embedding Gambar 1 Flowchart and Removing Old Makeshift Table...")
    for p in doc.paragraphs[26:35]:
        if "Metode penelitian yang digunakan dalam studi ini" in p.text:
            set_p(p, (
                "Metode penelitian yang digunakan dalam studi ini adalah pendekatan eksperimental kuantitatif yang dirancang secara sistematis "
                "ke dalam enam tahapan berurutan: (1) Audit Dataset dan Ekstraksi Identitas Pasien, (2) Pembagian Data Berdasarkan Tiga Protokol "
                "(P1: split acak 70:10:20 per citra, P2: split resmi 5.232 latih / 624 uji dengan validasi berkelompok per pasien, dan P3: split ketat per pasien), "
                "(3) Pra-pengolahan Citra Digital (modus basic, clahe, dan Ortiz; penyeragaman resolusi 150×150 piksel) serta Ekstraksi Fitur Tekstur GLCM "
                "(konfigurasi GLCM-6, GLCM-24, GLCM-5 dengan Ng ∈ {32, 64, 256}), (4) Pemodelan Pipeline Klasifikasi menggunakan pustaka imblearn "
                "(penanganan ketidakseimbangan bobot kelas class_weight vs SMOTE, penyetelan hiperparameter default vs tuned pada SVM, Random Forest, dan KNN), "
                "(5) Penggabungan Model melalui Soft Voting Ensemble Classifier, serta (6) Evaluasi Diagnostik Komparatif, Uji Hipotesis McNemar, "
                "Analisis Keragaman Ensemble, Uji TOST, Studi Ablasi, dan Implementasi Purwarupa Web Streamlit. "
                "Rancangan alur penelitian disajikan pada Gambar 1, sedangkan diagram arsitektur metode yang diusulkan disajikan pada Gambar 2."
            ), font_size=10.0)

    # Insert drawing before Gambar 1 caption
    p_g1_caption = None
    for p in doc.paragraphs:
        if p.text.strip().startswith("Gambar 1") and "Alur Penelitian" in p.text:
            p_g1_caption = p
            break
            
    if p_g1_caption is not None:
        p_g1_img = p_g1_caption.insert_paragraph_before()
        p_g1_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_g1_img.paragraph_format.space_before = Pt(6)
        p_g1_img.paragraph_format.space_after = Pt(2)
        r_g1 = p_g1_img.add_run()
        r_g1.add_picture("figures_jurnal/gambar1_alur_penelitian.png", width=Inches(6.5))
        set_p(p_g1_caption, "Gambar 1 Alur Penelitian Sistematis Klasifikasi Pneumonia Berbasis GLCM dan Ensemble Learning", font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        print("   -> Gambar 1 flowchart image successfully inserted at 300 DPI!")

    # Remove the old makeshift 12-row flowchart table (Table index 1)
    tbl_flowchart = doc.tables[1]
    if len(tbl_flowchart.rows) == 12 and 'Tahap 1:' in tbl_flowchart.rows[0].cells[0].text:
        tbl_flowchart._element.getparent().remove(tbl_flowchart._element)
        print("   -> Old makeshift flowchart table successfully removed!")

    # 5. Subbab 3.1: Tambahkan Tabel 2 (Audit Pasien) & Tabel 3 (Distribusi Data)
    print("[5/10] Inserting Tabel 2 Patient ID Audit and Updating Tabel 3...")
    for p in doc.paragraphs:
        if "Dataset utama yang digunakan dalam penelitian ini adalah dataset radiografi ront" in p.text or "Data yang digunakan dalam penelitian ini bersumber dari" in p.text:
            set_p(p, (
                "Data yang digunakan dalam penelitian ini bersumber dari repositori publik Kermany et al. [4] sebanyak 5.856 citra radiografi dada anak "
                "(4.273 Pneumonia [72,97%] dan 1.583 Normal [27,03%]). Pemeriksaan integritas dataset dilakukan melalui kalkulasi hash kriptografis SHA-256 "
                "untuk mendeteksi duplikasi citra eksak serta verifikasi keterbacaan berkas citra JPEG. Hasil audit menunjukkan seluruh berkas valid "
                "secara struktural tanpa citra korup. Pemetaan nomor seri pasien dilakukan melalui ekstraksi ekspresi reguler (regex) nama berkas: "
                "pola penomoran berkas normal untuk kelas Normal dan pola personID untuk kelas Pneumonia. Melalui audit ini, "
                "teridentifikasi total 3.118 pasien unik (1.444 Normal dan 1.674 Pneumonia). Ringkasan temuan audit integritas disajikan pada Tabel 2. "
                "Berdasarkan audit tersebut, partisi eksperimen dirancang ke dalam tiga protokol: "
                "(1) Protokol P1 (split acak 70:10:20 per citra; 4.100 latih, 584 validasi, 1.172 uji [855 Pneumonia, 317 Normal]); "
                "(2) Protokol P2 (partisi resmi Kaggle 5.232 latih/val dan 624 uji [390 Pneumonia, 234 Normal]), di mana data latih resmi dibagi via "
                "StratifiedGroupKFold 5-fold berbasis ID pasien menjadi 4.185 citra sub-latih dan 1.047 citra validasi tetap dengan 5 pengulangan multi-seed "
                "(seed 42, 123, 456, 789, 2026); serta (3) Protokol P3 (split terkontrol per pasien 70:15:15; 4.113 latih, 863 validasi, 880 uji [618 Pneumonia, 262 Normal]) "
                "yang mengontrol pemisahan berbasis pasien pada seluruh partisi. Rincian sebaran citra disajikan pada Tabel 3."
            ), font_size=10.0)

            # Insert Tabel 2 Caption
            p_t2_cap = doc.add_paragraph()
            set_p(p_t2_cap, "Tabel 2 Ringkasan Hasil Audit Integritas Dataset dan Pemetaan Identitas Pasien Lintas Partisi Data", bold=True, font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)
            p._element.addnext(p_t2_cap._element)

            # Create Table 2 (Audit Table)
            t_audit = doc.add_table(rows=9, cols=5)
            t_audit.alignment = WD_TABLE_ALIGNMENT.CENTER
            set_table_borders(t_audit)
            
            headers_audit = ['Parameter / Karakteristik Audit', 'Dataset Keseluruhan', 'Partisi Resmi (P2)', 'Protokol Acak (P1)', 'Protokol Pasien (P3)']
            for c_idx, h in enumerate(headers_audit):
                set_cell(t_audit.rows[0].cells[c_idx], h, bold=True, font_size=7.2)
            
            rows_audit = [
                ['Jumlah Citra (Normal / Pneu / Total)', '1.583 / 4.273 / 5.856', 'Latih 5.232 / Uji 624', 'Latih 4.684 / Uji 1.172', 'Latih 4.976 / Uji 880'],
                ['Jumlah Pasien Unik Teridentifikasi', '1.444 Norm / 1.674 Pneu (3.118)', 'Latih 2.861 / Uji 427', 'Latih 2.784 / Uji 959', 'Latih 2.650 / Uji 468'],
                ['Jumlah Pasien Pneumonia di Uji', '1.674 pasien', '202 pasien', '593 pasien', '288 pasien'],
                ['Pasien Pneumonia Overlap dengan Latih', 'Tidak berlaku', '170 pasien (84,16% Pneu uji)', '510 pasien (86,00% Pneu uji)', '0 pasien terdeteksi overlap'],
                ['Pasien Normal Overlap dengan Latih', 'Tidak berlaku', '0 pasien terdeteksi overlap', '115 pasien (31,42% Norm uji)', '0 pasien terdeteksi overlap'],
                ['Total Pasien Uji Overlap dengan Latih', 'Tidak berlaku', '170 pasien (39,81% uji)', '625 pasien (65,17% uji)', '0 pasien terdeteksi overlap'],
                ['Citra Uji Berbagi ID Pasien (Patient Leakage)', 'Tidak berlaku', '341 citra (54,65% uji)', '1.080 citra (92,15% uji)', '0 citra terdeteksi leakage'],
                ['Duplikasi Hash SHA-256 (Bitwise Identical)', '30 pasang berkas identik', 'Dipertahankan sesuai repositori', 'Dipertahankan acak', 'Diisolasi per pasien']
            ]
            for r_idx, r_data in enumerate(rows_audit, start=1):
                for c_idx, val in enumerate(r_data):
                    set_cell(t_audit.rows[r_idx].cells[c_idx], val, bold=(r_idx in [4, 6]), font_size=7.0, align=WD_ALIGN_PARAGRAPH.CENTER)

            p_t2_cap._element.addnext(t_audit._element)

            # Insert Footnote for Table 2
            p_t2_note = doc.add_paragraph()
            set_p(p_t2_note, (
                "Catatan: Identifikasi pasien diekstraksi berdasarkan pola regex penamaan berkas. Pada partisi resmi Kaggle P2, 84,16% pasien Pneumonia pada data uji "
                "tumpang tindih dengan data latih, sedangkan kelas Normal teridentifikasi bebas tumpang tindih berdasarkan pola nomor berkas. Pada protokol acak P1, "
                "86,00% pasien Pneumonia uji (510 dari 593) dan 31,42% pasien Normal uji (115 dari 366) tumpang tindih dengan data latih (total 625 dari 959 pasien uji atau 65,17%), "
                "mencakup 92,15% citra uji. Protokol P3 mengontrol pemisahan berbasis pasien sehingga bebas kebocoran identitas pasien terdeteksi."
            ), italic=True, font_size=8.0, space_after=6.0)
            t_audit._element.addnext(p_t2_note._element)
            break

    # Caption Table 3 (Partition table)
    for p in doc.paragraphs:
        if p.text.strip().startswith("Tabel 2 Distribusi Dataset Citra Rontgen"):
            set_p(p, "Tabel 3 Distribusi Dataset Citra Rontgen Dada pada Tiga Protokol Pembagian Data", bold=True, font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)
            break

    # Synchronize Table 3 cells
    t3 = None
    for tbl in doc.tables:
        if len(tbl.rows) > 0 and 'Protokol Split' in tbl.rows[0].cells[0].text:
            t3 = tbl
            break
    if t3 is not None:
        set_cell(t3.rows[1].cells[3], "1.172 (855 Pneu / 317 Norm)", font_size=7.2)
        set_cell(t3.rows[2].cells[3], "624 (390 Pneu / 234 Norm)", font_size=7.2)
        set_cell(t3.rows[3].cells[3], "880 (618 Pneu / 262 Norm)", font_size=7.2)
        set_cell(t3.rows[3].cells[5], "Pemisahan terkontrol per pasien (bebas overlap terdeteksi)", font_size=7.2)

    # 6. Subbab 3.2, 3.3, 3.4, 3.5 (Equations & MCC Formula)
    print("[6/10] Updating Subbab 3.2, 3.3, 3.4, and 3.5 (Clean Subsections & MCC Formula)...")
    for p in doc.paragraphs:
        if "Tahap pra-pengolahan mengevaluasi tiga mode:" in p.text:
            set_p(p, (
                "Tahap pra-pengolahan mengevaluasi tiga mode: (1) basic (konversi grayscale 8-bit dan penyeragaman ukuran spasial ke 150×150 piksel menggunakan "
                "interpolasi area piksel untuk meminimalkan distorsi artefak aliasing); (2) clahe (peningkatan kontras adaptif CLAHE "
                "dengan clip limit 2,0 dan ukuran grid 8×8 piksel); serta (3) Ortiz (pipeline pra-pengolahan mengacu pada Ortiz-Toro et al. [15] "
                "yang mengintegrasikan penyesuaian histogram matching terhadap citra referensi median data latih, Background Normalization, "
                "dan CLAHE). Standarisasi skala fitur dilakukan secara ketat di dalam pipeline terisolasi StandardScaler pada Persamaan (1):"
            ), font_size=10.0)
        elif "di mana x adalah nilai fitur asli" in p.text:
            set_p(p, (
                "di mana x adalah nilai fitur asli, sedangkan rata-rata (μtrain) dan simpangan baku (σtrain) dihitung secara eksklusif hanya dari "
                "himpunan data latih dan kemudian diterapkan untuk mentransformasi data validasi dan data uji guna mencegah kebocoran skala "
                "fitur (scaler leakage). Perlu ditegaskan bahwa seluruh fitur tekstur GLCM diekstraksi secara independen per citra sehingga tidak terdapat "
                "kebocoran tingkat dataset selain penskalaan fitur tersebut. Kuantisasi tingkat keabuan Ng ∈ {32, 64, 256} dievaluasi secara parametrik "
                "dengan memetakan intensitas 8-bit [0, 255] ke dalam bin diskret [0, Ng − 1]. Ekstraksi GLCM dilakukan dengan konfigurasi symmetric = True "
                "dan normed = True pada jarak d = 1 piksel dan empat orientasi sudut θ ∈ {0°, 45°, 90°, 135°} sehingga menghasilkan matriks probabilitas p(i, j). "
                "Enam fitur tekstur orde kedua Haralick [10] yang diekstraksi meliputi Contrast (Persamaan 2), Dissimilarity (Persamaan 3), Homogeneity (Persamaan 4), "
                "Angular Second Moment / ASM (Persamaan 5), Energy (Persamaan 6), dan Correlation (Persamaan 7)."
            ), font_size=10.0)

    # Subbab 3.3 Pemodelan Klasifikasi & Penyeimbangan Kelas
    for p in doc.paragraphs:
        if "Seluruh eksperimen pemodelan dibungkus menggunakan" in p.text:
            set_p(p, (
                "Seluruh eksperimen pemodelan dibungkus menggunakan pipeline terisolasi berbasis pustaka imblearn guna menjamin bahwa seluruh transformasi data "
                "berjalan tanpa kebocoran informasi. Ketidakseimbangan kelas pada data latih (74,3% Pneumonia vs 25,7% Normal) berisiko membiaskan model ke kelas mayoritas. "
                "Untuk mengatasinya, dua pendekatan dievaluasi: (1) penyeimbangan algoritmik berbiaya rendah melalui pembobotan kelas terbalik class_weight='balanced' "
                "yang memodifikasi penalti fungsi kerugian secara proporsional wc = N / (2 · Nc) tanpa mengubah ruang fitur asli; serta (2) oversampling sintetis SMOTE "
                "(k_neighbors = 5) yang diisolasi ketat di dalam lipatan data latih dengan menginterpolasikan sampel sintetis di sepanjang garis hubung fitur. "
                "Namun, interpolasi linier SMOTE pada ruang fitur GLCM non-linier membawa risiko pembentukan artefak fitur yang tidak fisiologis. "
                "Pemodelan membandingkan konfigurasi default ala Ortiz-Toro et al. [15] (KNN k=65 jarak Euclidean, SVM RBF C=1,0, dan RF n=100) "
                "versus tuned yang dilatih pada sub-latih 4.185 citra dan diseleksi berbasis F1-Score Pneumonia pada validasi tetap 1.047 citra via PredefinedSplit. "
                "Ruang pencarian hiperparameter mencakup: KNN (k ∈ {3, 5, 7, 11, 25, 45, 65}), SVM (kernel ∈ {linear, RBF}, C ∈ {0,1, 1, 10, 100}, γ ∈ {'scale', 'auto'}), "
                "dan Random Forest (n_estimators ∈ {100, 200, 300}, max_depth ∈ {None, 10, 20}, min_samples_leaf ∈ {1, 2, 4})."
            ), font_size=10.0)
            break

    # Formulas 8-11 Intro
    for p in doc.paragraphs:
        txt = p.text.strip()
        if "Fungsi keputusan Support Vector Machine" in txt:
            set_p(p, "Fungsi keputusan Support Vector Machine (SVM) dengan kernel Radial Basis Function (RBF) dirumuskan pada Persamaan (8):", font_size=10.0)
        elif "di mana γ menyatakan parameter lebar/kelengkungan kernel" in txt:
            set_p(p, "di mana γ menyatakan parameter lebar/kelengkungan kernel dan ||x − x′||² adalah kuadrat jarak Euclidean antar-vektor fitur. Kriteria pemisahan Random Forest berbasis Gini Impurity dirumuskan pada Persamaan (9):", font_size=10.0)
        elif "di mana C menyatakan jumlah kelas diagnostik" in txt:
            set_p(p, "di mana C menyatakan jumlah kelas diagnostik (C = 2: Normal dan Pneumonia) serta pc menyatakan estimasi proporsi sampel kelas c. Fungsi jarak tetangga terdekat K-Nearest Neighbors (KNN) dirumuskan pada Persamaan (10):", font_size=10.0)
        elif "di mana d = 6 menyatakan dimensi ruang fitur GLCM-6" in txt:
            set_p(p, "di mana d = 6 menyatakan dimensi ruang fitur GLCM-6. Prediksi gabungan Soft Voting Ensemble menggabungkan rata-rata estimasi probabilitas posterior terkalibrasi dari ketiga model dasar pada Persamaan (11):", font_size=10.0)

    # Subbab 3.4
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt.startswith("3.4") and ("Protokol Seleksi Model" in txt or "Seleksi Model" in txt):
            set_p(p, "3.4   Protokol Seleksi Model dan Penyetelan Ambang Batas Keputusan", bold=True, font_size=11.0, align=WD_ALIGN_PARAGRAPH.LEFT)
        elif "Protokol seleksi model dilakukan secara bertingkat:" in txt:
            set_p(p, (
                "Protokol seleksi model dilakukan secara bertingkat: model tuned terbaik (SVM RBF C=100, RF n=300, KNN k=11) diseleksi pada himpunan validasi tetap "
                "1.047 citra via PredefinedSplit, kemudian diverifikasi stabilitasnya melalui StratifiedGroupKFold 5-fold per pasien dengan 5 pengulangan multi-seed "
                "(seed 42, 123, 456, 789, 2026). Penyetelan ambang batas keputusan probabilitas (threshold tuning) dievaluasi pada kurva probabilitas validasi "
                "untuk mencari ambang batas optimal T* yang memaksimumkan F1-Score atau mempertahankan recall tetap ≥ 90% (diperoleh T* = 0,49 pada SVM Tuned). "
                "Setelah konfigurasi final ditetapkan, pipeline dilatih ulang pada seluruh 5.232 citra latih resmi sebelum evaluasi komparatif pada 624 citra uji resmi P2."
            ), font_size=10.0)

    # Subbab 3.5
    for p in doc.paragraphs:
        txt = p.text.strip()
        if (txt.startswith("3.5") or txt.startswith("3.4")) and ("Metrik Evaluasi Diagnostik" in txt or "Reproduksibilitas" in txt):
            set_p(p, "3.5   Metrik Evaluasi Diagnostik, Inferensi Statistik, dan Reproduksibilitas", bold=True, font_size=11.0, align=WD_ALIGN_PARAGRAPH.LEFT)

    # Equations 12-19
    for p in doc.paragraphs:
        txt = p.text.strip()
        if "Accuracy = (TP + TN)" in txt:
            set_p(p, "Accuracy = (TP + TN) / (TP + TN + FP + FN)   (12)", font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        elif "Precision = TP / (TP + FP)" in txt:
            set_p(p, "Precision = TP / (TP + FP)   (13)", font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        elif "Recall = TP / (TP + FN)" in txt:
            set_p(p, "Recall = TP / (TP + FN)   (14)", font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        elif "F1-Score = (2 · Precision · Recall)" in txt:
            set_p(p, "F1-Score = (2 · Precision · Recall) / (Precision + Recall)   (15)", font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        elif "Specificity = TN / (TN + FP)" in txt:
            set_p(p, "Specificity = TN / (TN + FP)   (16)", font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        elif "NPV = TN / (TN + FN)" in txt:
            set_p(p, "NPV = TN / (TN + FN)   (17)", font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        elif "Balanced Accuracy = (Sensitivity + Specificity)" in txt:
            set_p(p, "Balanced Accuracy = (Sensitivity + Specificity) / 2   (18)", font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        elif "ROC-AUC = ∫ TPR d(FPR)" in txt:
            set_p(p, "ROC-AUC = ∫ TPR d(FPR)   (19)", font_size=10.0, align=WD_ALIGN_PARAGRAPH.CENTER)

    # Insert MCC Formula at Persamaan 20
    for p in doc.paragraphs:
        if "ROC-AUC = ∫ TPR d(FPR)" in p.text:
            p_mcc_intro = doc.add_paragraph()
            set_p(p_mcc_intro, "Untuk mengevaluasi keseimbangan performa model pada kondisi kelas tidak seimbang tanpa terbias oleh kelas mayoritas, Matthews Correlation Coefficient (MCC) dirumuskan pada Persamaan (20):", font_size=10.0)
            p._element.addnext(p_mcc_intro._element)

            # Table for MCC Equation (1 row, 2 cols)
            t_mcc = doc.add_table(rows=1, cols=2)
            t_mcc.alignment = WD_TABLE_ALIGNMENT.CENTER
            set_cell(t_mcc.rows[0].cells[0], "MCC = (TP · TN − FP · FN) / √((TP + FP)(TP + FN)(TN + FP)(TN + FN))", bold=False, font_size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell(t_mcc.rows[0].cells[1], "(20)", bold=False, font_size=9.5, align=WD_ALIGN_PARAGRAPH.RIGHT)
            p_mcc_intro._element.addnext(t_mcc._element)

            # Narrative explaining MCC, PFI, McNemar, TOST, CPU latency
            p_mcc_desc = doc.add_paragraph()
            set_p(p_mcc_desc, (
                "di mana TP adalah True Positive, TN adalah True Negative, FP adalah False Positive, dan FN adalah False Negative. "
                "Matthews Correlation Coefficient (MCC) mengevaluasi seluruh sel matriks konfusi dengan rentang dari −1 (ketidaksepakatan total), 0 (tebakan acak), hingga +1 (prediksi sempurna). "
                "Karakteristik simetris ini membuat MCC kebal terhadap bias kelas mayoritas sehingga memberikan gambaran keseimbangan diagnostik yang objektif. "
                "Analisis kontribusi fitur dilakukan melalui Permutation Feature Importance (PFI) dengan permutasi acak K=10 iterasi. Untuk mengatasi bias penyamaran "
                "multikolinieritas antar-fitur tekstur berkorelasi tinggi (|r| > 0,85 antara Contrast, Dissimilarity, dan Homogeneity; korelasi kuadratis antara ASM dan Energy), "
                "diterapkan analisis Grouped PFI dengan mengacak serentak kelompok variasi kontras lokal {Contrast, Dissimilarity, Homogeneity}, kelompok keteraturan {ASM, Energy}, "
                "dan kelompok korelasi linier {Correlation}, serta didukung analisis Drop-Column Importance dengan melatih ulang model tanpa kelompok fitur terkait. "
                "Inferensi statistik diuji menggunakan uji McNemar berpasangan dengan koreksi Bonferroni (α = 0,0083) dan uji ekivalensi Two One-Sided Tests (TOST) "
                "pada batas margin klinis Δ = ±2,0%. Ketidakpastian estimasi diukur melalui paired bootstrap 95% Confidence Interval (1.000 iterasi). "
                "Keragaman ensemble dianalisis menggunakan rasio ketidaksepakatan (disagreement measure) dan korelasi koefisien phi (r). "
                "Seluruh eksperimen dijalankan secara deterministik (seed 42) pada lingkungan Python 3.10+ (scikit-learn 1.6, OpenCV 4.11, scikit-image 0.25) "
                "pada prosesor CPU Apple M-series (RAM 16 GB) tanpa akselerasi GPU, dengan pencatatan latensi inferensi batch=1 (mean ± SD)."
            ), font_size=10.0)
            t_mcc._element.addnext(p_mcc_desc._element)
            break

    # 7. Relocate Gambar 4 to Section 4.3
    print("[7/10] Relocating Gambar 4 to Section 4.3...")
    p_g4_img = None
    p_g4_cap = None
    p_pfi_target = None
    for i, p in enumerate(doc.paragraphs):
        if 'rId8' in p._element.xml:
            p_g4_img = p
            p_g4_cap = doc.paragraphs[i+1]
        if 'Analisis Permutation Feature Importance' in p.text and i > 85:
            p_pfi_target = p

    if p_g4_img and p_g4_cap and p_pfi_target:
        p_pfi_target._element.addnext(p_g4_cap._element)
        p_pfi_target._element.addnext(p_g4_img._element)
        print("   -> Gambar 4 successfully relocated to Section 4.3!")

    # 8. Update Bab 4 & Bab 5 Kesimpulan
    print("[8/10] Updating Bab 4 (Subbab 4.1 - 4.4) and Bab 5 Kesimpulan...")
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        
        # Subbab 4.1
        if txt.startswith("4.1") and ("Audit Integritas Data" in txt or "Sensitivitas Protokol" in txt):
            set_p(p, "4.1   Audit Integritas Data dan Kinerja Lintas Protokol Pembagian Data (RQ1)", bold=True, font_size=11.0, align=WD_ALIGN_PARAGRAPH.LEFT)
        elif "Berdasarkan pengamatan visual pada Gambar 3" in txt:
            set_p(p, (
                "Berdasarkan pengamatan visual pada Gambar 3, citra paru Normal memperlihatkan area parenkim yang cenderung lebih gelap "
                "(radiolucent) dan seragam, sedangkan citra Pneumonia menunjukkan area opasitas keputihan (radiopaque) akibat akumulasi eksudat "
                "konsolidasi cairan di ruang alveoli [5], [19]. Tahap pra-pengolahan pada resolusi 150×150 piksel berhasil mempertahankan karakteristik tekstur makro "
                "sembari mereduksi beban komputasi ekstraksi GLCM secara efisien."
            ), font_size=10.0)
        elif "Hasil audit dataset dan evaluasi lintas tiga skema pembagian data" in txt or "Audit integritas data pada 5.856 citra rontgen dada" in txt:
            set_p(p, (
                "Hasil audit dataset dan evaluasi lintas tiga skema pembagian data (Tabel 2 dan Tabel 3) mengungkap temuan krusial: pada skema acak per citra P1 (n=1.172 uji), "
                "model meraih akurasi 87,03% (F1-Score 91,00%). Namun, audit ID pasien mengidentifikasi 510 dari 593 pasien Pneumonia uji (86,00%) dan 115 dari 366 pasien Normal uji "
                "(31,42%) tumpang tindih dengan data latih (total 625 dari 959 pasien uji atau 65,17%), mencakup 92,15% citra uji yang berbagi identitas pasien, "
                "mengonfirmasi adanya kebocoran data. Pada protokol partisi resmi Kaggle P2 (n=624 uji), akurasi terkoreksi menjadi 79,33% (F1-Score 84,88% pada Ensemble dan 84,62% pada SVM Tuned). "
                "Menariknya, pada protokol terkontrol murni per pasien P3 (n=880 uji) yang bebas dari kebocoran pasien, akurasi tetap tinggi mencapai 87,61% (F1-Score 91,13%). "
                "Pola empiris P1 ≈ P3 > P2 mengindikasikan bahwa penurunan performa pada P2 bukan semata-mata dipicu oleh kebocoran identitas pasien, melainkan "
                "oleh perbedaan rasio prevalensi kelas dan pergeseran domain visual akuisisi alat rontgen. Himpunan uji P1 dan P3 membawa prevalensi pneumonia ~72,9%, sedangkan uji P2 "
                "memiliki prevalensi 62,50% (390 Pneumonia vs 234 Normal). "
                "Guna mengisolasi efek prevalensi, dilakukan eksperimen subsampling 100 pengulangan pada himpunan uji P1 dan P3 yang diselaraskan ke prevalensi resmi 62,50%. "
                "Hasil subsampling 100× menunjukkan bahwa akurasi P1 terkoreksi turun menjadi 81,73 ± 0,99% (95% CI: 79,78%–83,65%) dan P3 menjadi 82,40 ± 0,62% (95% CI: 81,18%–83,61%). "
                "Penyelarasan prevalensi ini mempersempit kesenjangan performa terhadap P2 (79,33%), menunjukkan bahwa gap optimisme sebesar ~7,7% hingga ~8,3% pada uji resmi P2 "
                "sekitar 5,3 poin persentase dipicu oleh perbedaan prevalensi kelas positif, dan sisanya (~2,4% hingga ~3,1%) merupakan pergeseran domain visual akuisisi (domain shift) "
                "pada himpunan uji Kaggle. Eksperimen pelatihan silang (cross-protocol training) kian menegaskan temuan ini: model yang dilatih pada P2 dan diuji ke P3 mencapai akurasi "
                "83,07%, sedangkan model yang dilatih pada P3 dan diuji ke P2 mencapai 83,49%."
            ), font_size=10.0)

        # Subbab 4.2
        elif txt.startswith("4.2") and ("Kinerja Model Klasifikasi" in txt or "Komparasi Model" in txt):
            set_p(p, "4.2   Komparasi Model Klasifikasi dan Analisis Keragaman Ensemble (RQ2)", bold=True, font_size=11.0, align=WD_ALIGN_PARAGRAPH.LEFT)
        elif ("Tabel 3" in txt or "Tabel 4" in txt) and "Komparasi Kinerja Model pada Protokol Utama P2" in txt:
            set_p(p, "Tabel 4 Komparasi Kinerja Model pada Protokol Utama P2 (Validasi, CV 5-Fold Per Pasien, dan Uji Resmi 624 Citra)", bold=True, font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        elif "Berdasarkan hasil evaluasi himpunan uji resmi P2" in txt or "Berdasarkan hasil uji resmi P2 (Tabel 3)" in txt:
            set_p(p, (
                "Berdasarkan hasil evaluasi himpunan uji resmi P2 (Tabel 4), di antara konfigurasi model yang dituning tanpa penyeimbangan, Soft Voting Ensemble "
                "dan SVM Tuned (RBF, C=100) meraih akurasi tertinggi sebesar 79,33%. "
                "Soft Voting mencatat recall 92,82%, spesifisitas 56,84%, F1-Score 84,88% (95% CI: 82,23%–87,56%), dan MCC 0,549; "
                "sedangkan SVM Tuned mencatat recall 91,03%, spesifisitas 59,83%, F1-Score 84,62%, dan MCC 0,548. "
                "Sebaliknya, model KNN Default (k=65) meraih spesifisitas 64,53%, recall 89,23%, dan MCC 0,563 di antara model tanpa penyeimbangan. "
                "Penyeimbangan kelas SMOTE pada SVM Tuned menghasilkan recall 87,95%, spesifisitas 66,24%, akurasi 79,81%, F1 84,48%, dan MCC 0,561. "
                "Sementara itu, pembobotan kelas biaya terbalik class_weight='balanced' pada SVM Tuned menghasilkan performa paling solid: "
                "akurasi 80,77%, spesifisitas 68,80%, recall 87,95%, PPV 82,45%, F1-Score 85,11%, Macro-F1 78,98%, dan MCC 0,583. "
                "Hal ini mengonfirmasi bahwa class_weight lebih efektif dalam menggeser batas keputusan tanpa merusak distribusi fitur dengan sampel sintetis. "
                "Penyetelan ambang batas probabilitas (threshold tuning pada T* = 0,49 yang diseleksi untuk mempertahankan recall ≥ 90% pada validasi internal) "
                "pada data uji menghasilkan akurasi 79,33%, recall 91,28%, spesifisitas 59,40%, PPV 78,94%, NPV 80,35%, F1 84,66%, Macro-F1 76,48%, dan MCC 0,548."
            ), font_size=10.0)

        # Captions Gambar 5 & 6
        elif txt.startswith("Gambar 5") and "Matriks Konfusi" in txt:
            set_p(p, "Gambar 5 Matriks Konfusi (Confusion Matrix) pada Data Uji Resmi P2 (624 Citra) untuk Model: (a) SVM Tuned (RBF, C=100), (b) Random Forest Tuned (n=300), (c) KNN Tuned (k=11), dan (d) Soft Voting Ensemble", font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        elif txt.startswith("Gambar 6") and "Perbandingan Metrik Evaluasi" in txt:
            set_p(p, "Gambar 6 Perbandingan Metrik Evaluasi Diagnostik Antar-Model: (a) Kinerja Diagnostik Model pada Protokol P2 (624 Citra) dan (b) Variasi Kinerja Model Lintas Tiga Protokol Pembagian Data (P1, P2, P3)", font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        elif ("Tabel 4" in txt or "Tabel 5" in txt) and "Hasil Uji Signifikansi McNemar" in txt:
            set_p(p, "Tabel 5 Hasil Uji Signifikansi McNemar Berpasangan dan Analisis Keragaman Ensemble (Uji Resmi P2, 624 Citra)", bold=True, font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        elif "Uji McNemar berpasangan dengan koreksi Bonferroni" in txt and i < 115: # Ensure only in Section 4.2!
            set_p(p, (
                "Uji McNemar berpasangan dengan koreksi Bonferroni (α = 0,0083 untuk 6 pasangan model utama) pada Tabel 5 menunjukkan bahwa tidak terdapat perbedaan "
                "signifikan antara Soft Voting dan model individual: pasangan Soft Voting vs SVM mencatat b=10, c=10, statistik 0,050, p = 0,823 (disagreement 3,21%, r=0,883); "
                "Soft Voting vs KNN mencatat b=16, c=9, statistik 1,440, p = 0,230 (disagreement 4,01%, r=0,856); serta Soft Voting vs RF mencatat statistik 6,323, p = 0,012 (disagreement 4,97%, r=0,794). "
                "Pasangan SVM vs RF mencatat statistik 4,356, p = 0,037; dan pasangan SVM Tuned vs SVM + class_weight mencatat b=12, c=21, statistik 1,939, p = 0,164. "
                "Meskipun pasangan Soft Voting vs RF (p = 0,012) dan SVM vs RF (p = 0,037) mencapai p < 0,05 sebelum koreksi, keduanya tidak signifikan pada ambang batas Bonferroni α = 0,0083. "
                "Uji ekivalensi Two One-Sided Tests (TOST) dengan margin batas ekivalensi klinis Δ = ±2,0% menghasilkan selisih akurasi 0,00% (95% CI: [−1,40%, +1,40%]) "
                "dengan p-value TOST sebesar 0,0026 (< 0,0083), mengonfirmasi bahwa performa Soft Voting dan SVM Tuned berada dalam rentang kesetaraan klinis. "
                "Tingginya korelasi galat antar-model (r = 0,713–0,883) mengonfirmasi bahwa batas efektivitas ensemble tercapai akibat homogenitas ruang fitur GLCM."
            ), font_size=10.0)

        # Subbab 4.3 Ablasi
        elif txt.startswith("4.3") and "Studi Ablasi" in txt:
            set_p(p, "4.3   Studi Ablasi Pra-Pengolahan, Kuantisasi Ng, Varian Fitur, dan Urgensi Batas Toraks (RQ3)", bold=True, font_size=11.0, align=WD_ALIGN_PARAGRAPH.LEFT)
        elif ("Tabel 5" in txt or "Tabel 6" in txt) and "Ringkasan Hasil Studi Ablasi" in txt:
            set_p(p, "Tabel 6 Ringkasan Hasil Studi Ablasi (11 Konfigurasi Representatif dari 27 Kombinasi pada Uji Resmi P2)", bold=True, font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        elif "Dekomposisi faktor ablasi (RQ2)" in txt or "Dekomposisi faktor ablasi (RQ3)" in txt:
            set_p(p, (
                "Dekomposisi faktor ablasi (RQ3) pada Tabel 6 mengungkap tiga temuan fundamental: (a) varian fitur GLCM-6 yang merata-ratakan empat sudut spasial "
                "merupakan komponen paling krusial untuk mencegah penurunan drastis spesifisitas. Pada varian GLCM-24 tanpa perataan, model mengalami overfitting parah "
                "di mana F1 validasi mencapai 94,84%–95,55%, namun akurasi uji anjlok ke 71,96%–74,20% dan spesifisitas kolaps ke 34,62%–42,31% akibat kutukan dimensionalitas; "
                "(b) varian GLCM-5 (tanpa Energy: Energy = √ASM) menghasilkan performa uji yang setara dengan GLCM-6 (akurasi 79,01%, F1 84,39%, spesifisitas 59,40%), "
                "mengonfirmasi bahwa Energy merupakan fitur kuadratis redundan terhadap ASM; serta (c) mode pra-pengolahan Ortiz pada Ng = 32 menghasilkan akurasi uji 82,37% "
                "(F1 86,75%, spesifisitas 65,81%), menunjukkan ketangguhan adaptasi kontras histogram."
            ), font_size=10.0)
        elif txt.startswith("Gambar 4") and "Analisis Kontribusi Fitur GLCM" in txt:
            set_p(p, "Gambar 4 Analisis Kontribusi Fitur GLCM Berdasarkan Permutation Feature Importance pada Model SVM Tuned (Uji Resmi P2, 624 Citra)", font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        elif "Analisis Permutation Feature Importance" in txt and not txt.startswith("Gambar"):
            set_p(p, (
                "Analisis Permutation Feature Importance (PFI) individual pada model SVM Tuned (Gambar 4) mengidentifikasi kontribusi keenam fitur secara berurutan: "
                "Dissimilarity (penurunan F1 7,51% ± 0,58%), Homogeneity (penurunan F1 6,46% ± 0,46%), Contrast (penurunan F1 4,93% ± 0,78%), "
                "Energy (penurunan F1 4,11% ± 0,71%), Correlation (penurunan F1 3,46% ± 0,68%), dan ASM (penurunan F1 2,29% ± 0,61%). "
                "Namun, karena fitur GLCM memiliki multikolinieritas tinggi, analisis Grouped PFI dijalankan untuk mengungkap signifikansi sebenarnya tanpa bias penyamaran. "
                "Hasil Grouped PFI menunjukkan bahwa kelompok variasi kontras lokal {Contrast, Dissimilarity, Homogeneity} merupakan pendorong utama klasifikasi "
                "dengan penurunan F1-Score drastis sebesar 18,16% ± 1,19% (dan penurunan akurasi 24,42% ± 1,48%). Kelompok keteraturan {ASM, Energy} menyumbang "
                "penurunan F1 sebesar 10,94% ± 0,96% (penurunan akurasi 15,87% ± 1,39%), sedangkan kelompok Correlation menyumbang 6,34% ± 0,77% (penurunan akurasi 7,44% ± 1,02%). "
                "Analisis Drop-Column Importance mengonfirmasi temuan ini: penghapusan kelompok kontras lokal menyebabkan F1-Score anjlok ke 76,12% (penurunan 8,76%), "
                "sedangkan penghapusan kelompok keteraturan hanya menurunkan F1-Score sebesar 1,12%."
            ), font_size=10.0)
        elif ("Tabel 6" in txt or "Tabel 7" in txt) and "Uji Ketahanan Model terhadap Pemotongan Tepi" in txt:
            set_p(p, "Tabel 7 Uji Ketahanan Model terhadap Pemotongan Tepi Citra (Peripheral Center Crop Test pada Uji Resmi P2)", bold=True, font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        elif "Dari sisi efisiensi komputasi CPU" in txt:
            set_p(p, (
                "Dari sisi efisiensi komputasi CPU, seluruh pipeline inferensi yang disimpan via format joblib beroperasi sangat ringan. "
                "Pengukuran latensi batch=1 pada prosesor CPU Apple M-series (RAM 16 GB) menunjukkan waktu pra-pengolahan dan ekstraksi GLCM rata-rata "
                "24,42 ± 1,15 ms per citra, sedangkan inferensi prediksi model hanya membutuhkan 0,0671 ± 0,0042 ms (SVM Tuned), 0,1150 ± 0,0081 ms (RF Tuned), "
                "dan 0,1821 ± 0,0112 ms (Soft Voting). Hal ini menegaskan bahwa pipeline GLCM terisolasi sangat aplikatif untuk diterapkan pada perangkat berdaya rendah."
            ), font_size=10.0)
        elif ("Tabel 7" in txt or "Tabel 8" in txt) and "Perbandingan Kinerja pada Himpunan Uji Resmi yang Sama" in txt:
            set_p(p, "Tabel 8 Perbandingan Kinerja pada Himpunan Uji Resmi yang Sama (624 Citra) dengan Pembanding Utama [15] dan Studi Pendahulu [19]", bold=True, font_size=9.0, align=WD_ALIGN_PARAGRAPH.CENTER)

        # Subbab 4.4 Ancaman terhadap Validitas (Expanded for McNemar, Bootstrap, Multi-seed)
        elif (txt.startswith("5") or txt.startswith("4.4")) and "Ancaman Validitas" in txt:
            set_p(p, "4.4   Ancaman terhadap Validitas (Threats to Validity)", bold=True, font_size=11.0, align=WD_ALIGN_PARAGRAPH.LEFT)
        elif "Terdapat beberapa ancaman terhadap validitas" in txt:
            set_p(p, (
                "Terdapat beberapa ancaman terhadap validitas (threats to validity) dan keterbatasan metodologis yang perlu dicatat secara transparan: "
                "(1) Validitas konstruk identitas pasien: penomoran berkas kelas Normal berhasil dipetakan ke 1.444 pasien unik bebas overlap terdeteksi, "
                "namun verifikasi identitas biologis definitif tetap memerlukan metadata klinis primer rumah sakit; "
                "(2) Risiko shortcut learning: ketiadaan segmentasi ROI paru otomatis memungkinkan fitur GLCM global menangkap kontras batas toraks luar dan diafragma, "
                "sebagaimana ditunjukkan oleh penurunan akurasi pada uji peripheral center crop 10%–15% (turun ke 63,14%) dan uji masking tepi tanpa resize (turun ke 62,50% dengan spesifisitas 0,00%); "
                "(3) Risiko distorsi tekstur akibat penyesuaian ukuran spasial (resizing 150×150 piksel) yang dimitigasi dengan interpolasi area piksel; "
                "(4) Efek penyamaran multikolinieritas pada analisis PFI individual yang mengaburkan kontribusi fitur berkorelasi tinggi (|r| > 0,85), yang diatasi secara empiris melalui analisis Grouped PFI dan Drop-Column Importance; "
                "(5) Keterbatasan Uji Signifikansi McNemar: Uji McNemar berpasangan mengevaluasi proporsi diskordan pada satu himpunan uji berpasangan statis tunggal (N=624 citra). "
                "Asumsi independensi observasi berpotensi menghasilkan estimasi varians yang lebih sempit ketika terdapat efek kluster dari 170 pasien Pneumonia berulang pada data uji. "
                "Selain itu, uji McNemar pada satu partisi uji tidak mengukur variabilitas pelatihan model yang timbul dari variasi sampling fold latihan; "
                "(6) Keterbatasan Resampling Bootstrap Berpasangan: Estimasi interval kepercayaan berbasis Bootstrap berpasangan (B=1.000) menangkap variabilitas sampling di dalam distribusi "
                "himpunan uji yang diamati, namun tidak mampu memitigasi bias sampling bawaan dari dataset asal maupun memodelkan pergeseran domain di luar distribusi (out-of-distribution shifts) "
                "pada populasi rumah sakit baru dengan protokol pencitraan yang berbeda; "
                "(7) Keterbatasan Partisi Statis Tunggal dan Multi-Seed: Meskipun stabilitas multi-seed (5 pengulangan: seed 42, 123, 456, 789, 2026) telah diverifikasi pada validasi silang 5-fold "
                "dan subsampling Monte Carlo 100 kali telah dijalankan untuk P1 dan P3, himpunan uji resmi P2 (N=624) merupakan partisi statis tunggal tanpa validasi silang tingkat pasien antar-institusi, "
                "sehingga metrik kinerja pada P2 mencerminkan estimasi titik (point estimate) pada kohort uji tersebut dan bukan distribusi performa lintas senter medis independen; serta "
                "(8) Generalisasi klinis terbatas pada populasi pediatrik usia 1–5 tahun dari satu fasilitas kesehatan (Guangzhou Women and Children's Medical Center [4]) tanpa validasi prospektif eksternal multi-pusat, "
                "serta ketiadaan arsitektur CNN pembanding yang dilatih secara mandiri pada protokol yang sama (di mana model GLCM diposisikan sebagai alternatif komputasi ringan berdaya rendah)."
            ), font_size=10.0)

        # Bab 5 Kesimpulan (RQ1, RQ2, RQ3 Explicitly Answered)
        elif (txt.startswith("6") or txt.startswith("5")) and "Kesimpulan" in txt and len(txt) < 30:
            set_p(p, "5   Kesimpulan", bold=True, font_size=12.0, align=WD_ALIGN_PARAGRAPH.LEFT)
            if i + 1 < len(doc.paragraphs):
                p_next = doc.paragraphs[i+1]
                set_p(p_next, (
                    "Penelitian ini menyimpulkan tiga temuan utama sesuai pertanyaan penelitian yang dirumuskan: "
                    "(1) Jawaban RQ1: Kinerja klasifikasi berbasis fitur tekstur GLCM sangat sensitif terhadap skema partisi dan prevalensi kelas uji. "
                    "Akurasi mencapai 87,03% (F1-Score 91,00%) pada protokol acak per citra P1 (n=1.172 uji) dan 87,61% (F1-Score 91,13%) pada protokol terkontrol pasien P3 (n=880 uji), "
                    "namun terkoreksi turun menjadi 79,33% (F1-Score 84,88% pada Ensemble dan 84,62% pada SVM) pada protokol resmi Kaggle P2 (n=624 uji). "
                    "Melalui 100× subsampling Monte Carlo yang menyamakan prevalensi P1 dan P3 ke rasio uji resmi 62,50% (di mana akurasi terkoreksi menjadi 81,73% ± 0,99% pada P1 dan 82,40% ± 0,62% pada P3) "
                    "serta pelatihan silang dua arah (P2 latih → P3 uji: 83,07%; P3 latih → P2 uji: 83,49%), penelitian ini mengonfirmasi bahwa kesenjangan performa sebesar ~7,7% hingga ~8,3% "
                    "pada himpunan uji resmi P2 sebagian besar (~5,3 poin persentase) dipicu oleh perbedaan prevalensi kelas positif dan pergeseran domain visual akuisisi alat rontgen (Kaggle domain shift), "
                    "bukan semata-mata kebocoran identitas pasien; "
                    "(2) Jawaban RQ2: Pada protokol utama P2, Soft Voting Ensemble dan SVM Tuned meraih akurasi tertinggi sebesar 79,33% di antara model yang dituning tanpa penyeimbangan kelas, "
                    "dengan F1-Score masing-masing 84,88% (95% CI: 82,23%–87,56%) dan 84,62% (95% CI: 81,96%–87,32%), serta MCC 0,549 dan 0,548. KNN Default (k=65) meraih spesifisitas 64,53% "
                    "dan MCC 0,563. Penyeimbangan kelas class_weight='balanced' pada SVM Tuned menaikkan spesifisitas secara substansial ke 68,80% (akurasi 80,77%, recall 87,95%, MCC 0,583), "
                    "lebih efisien dan bebas artefak sintetis dibanding SMOTE (akurasi 79,81%, spesifisitas 66,24%, recall 87,95%, MCC 0,561), meskipun secara statistik peningkatan ini tidak berbeda signifikan "
                    "terhadap model tanpa penyeimbangan (p = 0,164). Uji signifikansi McNemar berpasangan (α = 0,0083) dan uji ekivalensi TOST (p = 0,0026, 95% CI: [−1,40%, +1,40%] pada margin ±2%) "
                    "menunjukkan tidak ada bukti perbedaan performa antara Soft Voting dan SVM Tuned, karena tingginya korelasi galat (r = 0,713–0,883) membatasi diversitas ensemble pada ruang fitur GLCM yang homogen; serta "
                    "(3) Jawaban RQ3: Dekomposisi ablasi 27 kombinasi pada validasi tetap menunjukkan bahwa perataan empat arah spasial (μ) pada varian GLCM-6 krusial untuk mencegah penurunan "
                    "drastis spesifisitas hingga 34%–36% yang dialami GLCM-24 tanpa perataan akibat overfitting dimensi fitur. Varian GLCM-5 (tanpa Energy) menunjukkan performa yang praktis setara "
                    "dengan GLCM-6 (akurasi 79,01% vs 79,01%), mengonfirmasi redundansi kuadratis antara Energy dan ASM. Kombinasi pra-pengolahan kontras adaptif Ortiz pada kuantisasi Ng = 32 "
                    "memberikan akurasi uji terbaik 82,37% (F1 86,75%, spesifisitas 65,81%). Analisis permutasi berkelompok (Grouped PFI) dan drop-column menempatkan kelompok variasi kontras lokal "
                    "{Contrast, Dissimilarity, Homogeneity} sebagai pendorong utama klasifikasi (penurunan F1 18,16% ± 1,19% dan penurunan akurasi 24,42%), sedangkan uji pemotongan tepi tanpa resize "
                    "(penurunan akurasi hingga 62,50% dan spesifisitas 0,00%) mengonfirmasi ketergantungan fitur tekstur global pada batas anatomi perifer toraks luar. "
                    "Secara praktis, model GLCM-6 dengan pipeline terisolasi menawarkan alternatif diagnosis komputasi ringan berdaya rendah (latensi model CPU 0,067 ms) untuk skrining awal "
                    "di fasilitas kesehatan perifer. Untuk pengembangan selanjutnya, integrasi segmentasi otomatis batas paru (ROI) sangat direkomendasikan guna mengeliminasi ketergantungan "
                    "pada batas luar dinding dada, disertai validasi prospektif multi-senter dengan citra radiografi beranotasi klinis lengkap."
                ), font_size=10.0)

    # 9. Update Table 4 (Model Evaluation P2)
    print("[9/10] Updating Table 4 with Exact Metrics, MCC, class_weight, and Threshold Tuning...")
    t_mod = None
    for tbl in doc.tables:
        if len(tbl.rows) > 0 and 'Model Konfigurasi' in tbl.rows[0].cells[0].text:
            t_mod = tbl
            break
            
    if t_mod is not None:
        if len(t_mod.columns) == 12:
            new_col = t_mod.add_column(Inches(0.45))
            mcc_vals = ['MCC', '0,000', '0,511', '0,548', '0,471', '0,494', '0,563', '0,522', '0,549', '0,561', '0,560']
            for r_idx in range(len(t_mod.rows)):
                roc_val = t_mod.rows[r_idx].cells[11].text
                set_cell(new_col.cells[r_idx], roc_val, bold=(r_idx==0), font_size=6.8)
                set_cell(t_mod.rows[r_idx].cells[11], mcc_vals[r_idx], bold=(r_idx==0), font_size=6.8)

        # Synchronize exact values for existing rows
        # Row 4: KNN_default -> Recall 89.23%
        if len(t_mod.rows) > 5 and 'KNN' in t_mod.rows[5].cells[0].text:
            set_cell(t_mod.rows[5].cells[5], "89,23", font_size=6.8) # Recall
        # Row 7: SVM SMOTE -> Recall 87.95, Specificity 66.24
        if len(t_mod.rows) > 8 and 'SMOTE' in t_mod.rows[8].cells[0].text:
            set_cell(t_mod.rows[8].cells[5], "87,95", font_size=6.8)
            set_cell(t_mod.rows[8].cells[6], "66,24", font_size=6.8)

        # Add row SVM Tuned + class_weight
        has_cw = any('class_weight' in r.cells[0].text for r in t_mod.rows)
        if not has_cw:
            r_cw = t_mod.add_row()
            cw_data = ['SVM Tuned + class_weight', '-', '-', '80,77', '82,45', '87,95', '68,80', '77,40', '85,11', '78,38', '78,98', '0,583', '0,8190']
            for c_idx, val in enumerate(cw_data):
                set_cell(r_cw.cells[c_idx], val, bold=False, font_size=6.8)

        # Add row SVM Tuned + Threshold Tuning (Exact PPV=78.94, NPV=80.35, Macro-F1=76.48)
        has_tt = any('Threshold Tuning' in r.cells[0].text for r in t_mod.rows)
        if not has_tt:
            r_tt = t_mod.add_row()
            tt_data = ['SVM Tuned + Threshold Tuning (T=0,49)', '92,15', '-', '79,33', '78,94', '91,28', '59,40', '80,35', '84,66', '75,34', '76,48', '0,548', '0,8190']
            for c_idx, val in enumerate(tt_data):
                set_cell(r_tt.cells[c_idx], val, bold=False, font_size=6.8)

    # 10. Update Table 5 (McNemar)
    print("[10/10] Updating Table 5, Table 6, Table 7, Table 8 and Embedding Images...")
    t_mcn = None
    for tbl in doc.tables:
        if len(tbl.rows) > 0 and 'Pasangan Model' in tbl.rows[0].cells[0].text:
            t_mcn = tbl
            break
            
    if t_mcn is not None:
        mcn_data = [
            ['Soft Voting vs SVM (Tuned)', '10', '10', '0,050', '0,823', '3,21%', '0,883', '0,987', 'Tidak ada bukti perbedaan (p > 0,0083)'],
            ['Soft Voting vs KNN (Tuned)', '16', '9', '1,440', '0,230', '4,01%', '0,856', '0,985', 'Tidak ada bukti perbedaan (p > 0,0083)'],
            ['Soft Voting vs RF (Tuned)', '23', '8', '6,323', '0,012', '4,97%', '0,794', '0,976', 'Tidak ada bukti perbedaan (p > 0,0083; p < 0,05 tanpa koreksi)'],
            ['SVM (Tuned) vs KNN (Tuned)', '18', '11', '1,241', '0,265', '4,65%', '0,832', '0,980', 'Tidak ada bukti perbedaan (p > 0,0083)'],
            ['SVM (Tuned) vs RF (Tuned)', '30', '15', '4,356', '0,037', '7,21%', '0,713', '0,961', 'Tidak ada bukti perbedaan (p > 0,0083; p < 0,05 tanpa koreksi)'],
            ['RF (Tuned) vs KNN (Tuned)', '24', '18', '0,595', '0,440', '6,73%', '0,736', '0,966', 'Tidak ada bukti perbedaan (p > 0,0083)'],
            ['SVM (Tuned) vs SVM + class_weight', '12', '21', '1,939', '0,164', '5,29%', '0,841', '0,981', 'Tidak ada bukti perbedaan (p > 0,0083)']
        ]
        while len(t_mcn.rows) > 1:
            t_mcn._element.remove(t_mcn.rows[1]._element)
        for r_data in mcn_data:
            row = t_mcn.add_row()
            for c_idx, val in enumerate(r_data):
                set_cell(row.cells[c_idx], val, bold=False, font_size=6.8)

    # Update Table 6 (Ablation Matrix EXACTLY from e4_tabel_ablasi_validasi_seleksi.csv)
    t_abl = None
    for tbl in doc.tables:
        if len(tbl.rows) > 0 and 'Pra-Pengolahan' in tbl.rows[0].cells[0].text:
            t_abl = tbl
            break
            
    if t_abl is not None:
        abl_headers = ['Pra-Pengolahan', 'Ng', 'Varian Fitur', 'Dimensi', 'F1 Val (%)', 'F1 Uji (%)', 'Akurasi Uji (%)', 'Spesifisitas Uji (%)', 'Rank Val']
        for c_idx, h in enumerate(abl_headers):
            set_cell(t_abl.rows[0].cells[c_idx], h, bold=True, font_size=6.8)
        
        abl_data = [
            ['basic', '256', 'GLCM-24 (4 Arah Terpisah)', '24', '95,55', '81,05', '72,28', '34,62', '1'],
            ['clahe', '256', 'GLCM-24 (4 Arah Terpisah)', '24', '95,50', '81,54', '73,08', '36,32', '2'],
            ['basic', '64', 'GLCM-24 (4 Arah Terpisah)', '24', '95,49', '81,23', '72,60', '35,47', '3'],
            ['basic', '32', 'GLCM-24 (4 Arah Terpisah)', '24', '94,84', '80,71', '71,96', '35,47', '6'],
            ['Ortiz', '32', 'GLCM-24 (4 Arah Terpisah)', '24', '94,46', '81,32', '72,76', '35,90', '7'],
            ['basic', '256', 'GLCM-6 (Rata-rata 4 Arah)', '6', '91,69', '84,00', '78,69', '60,68', '11'],
            ['basic', '64', 'GLCM-6 (Rata-rata 4 Arah)', '6', '91,59', '84,13', '78,85', '60,68', '12'],
            ['basic', '32', 'GLCM-5 (Tanpa Energy: E=√ASM)', '5', '91,18', '84,39', '79,01', '59,40', '16'],
            ['basic', '32', 'GLCM-6 (Rata-rata 4 Arah)', '6', '91,13', '84,42', '79,01', '58,97', '17'],
            ['clahe', '32', 'GLCM-6 (Rata-rata 4 Arah)', '6', '90,92', '84,13', '78,85', '60,68', '20'],
            ['Ortiz', '32', 'GLCM-6 (Rata-rata 4 Arah)', '6', '89,95', '86,75', '82,37', '65,81', '25']
        ]
        while len(t_abl.rows) > 1:
            t_abl._element.remove(t_abl.rows[1]._element)
        for r_idx, r_data in enumerate(abl_data):
            row = t_abl.add_row()
            for c_idx, val in enumerate(r_data):
                set_cell(row.cells[c_idx], val, bold=(r_idx in [0, 8, 10]), font_size=6.8)

    # Update Table 7 (Edge Masking Test)
    t_crop = None
    for tbl in doc.tables:
        if len(tbl.rows) > 0 and ('Kondisi Manipulasi Tepi' in tbl.rows[0].cells[0].text or 'Kondisi Citra Masukan' in tbl.rows[0].cells[0].text):
            t_crop = tbl
            break
            
    if t_crop is not None:
        crop_headers = ['Kondisi Manipulasi Tepi', 'Akurasi (%)', 'F1-Score (%)', 'Spesifisitas (%)', 'Delta Akurasi (%)']
        for c_idx, h in enumerate(crop_headers):
            set_cell(t_crop.rows[0].cells[c_idx], h, bold=True, font_size=7.0)
        
        crop_data = [
            ['Citra Utuh (Baseline 150×150)', '79,33', '84,62', '59,83', '0,00 (Ref)'],
            ['Center Crop 5% Perifer (Resize)', '76,92', '83,18', '53,42', '-2,41'],
            ['Center Crop 10% Perifer (Resize)', '71,15', '79,30', '41,45', '-8,18'],
            ['Center Crop 15% Perifer (Resize)', '63,14', '74,48', '20,94', '-16,19'],
            ['Masking Tepi 10% Tanpa Resize (Non-Resized)', '66,99', '77,93', '14,96', '-12,34'],
            ['Masking Tepi 15% Tanpa Resize (Non-Resized)', '62,50', '76,92', '0,00', '-16,83']
        ]
        while len(t_crop.rows) > 1:
            t_crop._element.remove(t_crop.rows[1]._element)
        for r_idx, r_data in enumerate(crop_data):
            row = t_crop.add_row()
            for c_idx, val in enumerate(r_data):
                set_cell(row.cells[c_idx], val, bold=(r_idx in [0, 4, 5]), font_size=7.0)

    # Update Table 8 (Benchmark with Previous Studies)
    t_lit = None
    for tbl in doc.tables:
        if len(tbl.rows) > 0 and 'Metode / Publikasi' in tbl.rows[0].cells[0].text:
            t_lit = tbl
            break
            
    if t_lit is not None:
        lit_headers = ['Metode / Publikasi', 'Ruang Fitur & Pra-pengolahan', 'Hasil Pengujian Resmi (624 Citra)']
        for c_idx, h in enumerate(lit_headers):
            set_cell(t_lit.rows[0].cells[c_idx], h, bold=True, font_size=7.2)
        
        lit_data = [
            ['Ortiz-Toro et al. (2022) [15]', '18 fitur (Radiomics + Fractal + Histon), pra-pengolahan kontras adaptif tanpa penyeimbangan', 'Random Forest: Akurasi 82,21%, Sensitivitas 89,74%, Spesifisitas 69,66%, F1 86,10%'],
            ['Hismaya & Sulistianingsih (2025) [19]', '24 fitur GLCM (4 arah sudut tanpa perataan, d=1), tanpa penyeimbangan kelas', 'SVM RBF: Akurasi 76,12% (138+337)/624, Sensitivitas 86,41%, Spesifisitas 58,97%, F1 81,80%'],
            ['Penelitian Ini (Baseline Tuned)', 'GLCM-6 (Rata-rata 4 arah, d=1, Ng=32), StandardScaler fit eksklusif di data latih', 'Akurasi 79,33%, Sensitivitas 91,03%, Spesifisitas 59,83%, F1 84,62%, MCC 0,548'],
            ['Penelitian Ini (Tuned + class_weight)', 'GLCM-6 (Ortiz/Basic, Ng=32) + penyeimbangan bobot kelas class_weight=\'balanced\'', 'Akurasi 80,77%, Sensitivitas 87,95%, Spesifisitas 68,80%, F1 85,11%, MCC 0,583']
        ]
        while len(t_lit.rows) > 1:
            t_lit._element.remove(t_lit.rows[1]._element)
        for r_idx, r_data in enumerate(lit_data):
            row = t_lit.add_row()
            for c_idx, val in enumerate(r_data):
                set_cell(row.cells[c_idx], val, bold=(r_idx >= 2), font_size=7.0, align=WD_ALIGN_PARAGRAPH.LEFT if c_idx < 2 else WD_ALIGN_PARAGRAPH.CENTER)

        # Append CNN context row
        r_cnn = t_lit.add_row()
        set_cell(r_cnn.cells[0], "Konteks CNN Mandiri (Arifin & Surya 2022 [20])", bold=False, font_size=6.8, align=WD_ALIGN_PARAGRAPH.LEFT)
        set_cell(r_cnn.cells[1], "Representasi fitur hierarkis konvolusional otomatis (224×224 piksel), bobot ImageNet", bold=False, font_size=6.8, align=WD_ALIGN_PARAGRAPH.LEFT)
        set_cell(r_cnn.cells[2], "Akurasi 88,00% (Model DL kompleks berdaya komputasi tinggi)", bold=False, font_size=6.8, align=WD_ALIGN_PARAGRAPH.CENTER)

    # Embed 300 DPI Figures
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

    # Synchronize References: Replace [22] with Zech et al.
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt.startswith("[5]") and "Cillóniz" in txt:
            set_p(p, "[5]  T. Cherian et al., “Standardized interpretation of paediatric chest radiographs for the diagnosis of pneumonia in epidemiological studies,” Bull. World Health Organ., vol. 83, no. 5, pp. 353–359, 2005.", font_size=9.0)
        elif txt.startswith("[22]") and "Akbar" in txt:
            set_p(p, "[22] J. R. Zech, M. A. Badgeley, M. Liu, A. B. Costa, J. J. Titano, and E. K. Oermann, “Variable generalization performance of a deep learning model to detect pneumonia in chest radiographs: A cross-sectional study,” PLoS Med., vol. 15, no. 11, p. e1002683, 2018.", font_size=9.0)

    # Append Supplementary Materials (Lampiran A, B, C)
    has_supp = any("LAMPIRAN: MATERI SUPLEMEN" in p.text for p in doc.paragraphs)
    if not has_supp:
        p_sep = doc.add_paragraph()
        p_sep.paragraph_format.page_break_before = True
        
        p_sup_title = doc.add_paragraph()
        set_p(p_sup_title, "LAMPIRAN: MATERI SUPLEMEN (SUPPLEMENTARY MATERIALS)", bold=True, font_size=12.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        
        p_sup_desc = doc.add_paragraph()
        set_p(p_sup_desc, (
            "Materi suplemen ini menyediakan rincian empiris lengkap yang mendukung temuan utama pada naskah: "
            "(A) Tabel lengkap dekomposisi 27 kombinasi studi ablasi pada data validasi tetap dan data uji resmi P2, "
            "(B) Rincian eksperimen subsampling Monte Carlo 100× penyelarasan prevalensi dan pelatihan silang lintas protokol (P2 ↔ P3), serta "
            "(C) Analisis mendalam Permutasi Berkelompok (Grouped PFI) dan Drop-Column Feature Importance untuk mengatasi multikolinieritas fitur GLCM."
        ), font_size=9.5)

        # Lampiran A
        p_la = doc.add_paragraph()
        set_p(p_la, "Lampiran A. Tabel Lengkap Studi Ablasi 27 Kombinasi Parametrik (P2 Validasi Tetap & Uji Resmi 624 Citra)", bold=True, font_size=10.0, align=WD_ALIGN_PARAGRAPH.LEFT)
        
        csv_la = "results/tables/e4_tabel_ablasi_validasi_seleksi.csv"
        if os.path.exists(csv_la):
            df_la = pd.read_csv(csv_la)
            t_la = doc.add_table(rows=len(df_la)+1, cols=len(df_la.columns))
            t_la.alignment = WD_TABLE_ALIGNMENT.CENTER
            set_table_borders(t_la)
            for c_idx, col in enumerate(df_la.columns):
                set_cell(t_la.rows[0].cells[c_idx], str(col), bold=True, font_size=6.0)
            for r_idx, row in df_la.iterrows():
                for c_idx, val in enumerate(row):
                    val_str = f"{val:.2f}" if isinstance(val, float) else str(val)
                    set_cell(t_la.rows[r_idx+1].cells[c_idx], val_str, bold=False, font_size=5.8)

        # Lampiran B
        p_lb = doc.add_paragraph()
        set_p(p_lb, "Lampiran B. Hasil Eksperimen Subsampling Prevalensi 100× dan Pelatihan Silang Lintas Protokol", bold=True, font_size=10.0, align=WD_ALIGN_PARAGRAPH.LEFT)
        
        csv_lb = "results/tables/e3_prevalence_subsampling_100x.csv"
        if os.path.exists(csv_lb):
            df_lb = pd.read_csv(csv_lb)
            t_lb = doc.add_table(rows=len(df_lb)+1, cols=len(df_lb.columns))
            t_lb.alignment = WD_TABLE_ALIGNMENT.CENTER
            set_table_borders(t_lb)
            for c_idx, col in enumerate(df_lb.columns):
                set_cell(t_lb.rows[0].cells[c_idx], str(col), bold=True, font_size=6.5)
            for r_idx, row in df_lb.iterrows():
                for c_idx, val in enumerate(row):
                    val_str = f"{val:.2f}" if isinstance(val, float) else str(val)
                    set_cell(t_lb.rows[r_idx+1].cells[c_idx], val_str, bold=False, font_size=6.2)

        # Lampiran C
        p_lc = doc.add_paragraph()
        set_p(p_lc, "Lampiran C. Analisis Permutasi Berkelompok (Grouped PFI) dan Drop-Column Feature Importance", bold=True, font_size=10.0, align=WD_ALIGN_PARAGRAPH.LEFT)
        
        csv_lc = "results/tables/e7_grouped_permutation_importance.csv"
        if os.path.exists(csv_lc):
            df_lc = pd.read_csv(csv_lc)
            t_lc = doc.add_table(rows=len(df_lc)+1, cols=len(df_lc.columns))
            t_lc.alignment = WD_TABLE_ALIGNMENT.CENTER
            set_table_borders(t_lc)
            for c_idx, col in enumerate(df_lc.columns):
                set_cell(t_lc.rows[0].cells[c_idx], str(col), bold=True, font_size=6.5)
            for r_idx, row in df_lc.iterrows():
                for c_idx, val in enumerate(row):
                    val_str = f"{val:.2f}" if isinstance(val, float) else str(val)
                    set_cell(t_lc.rows[r_idx+1].cells[c_idx], val_str, bold=False, font_size=6.2)

    # Save to target locations
    os.makedirs(os.path.dirname(output_docx), exist_ok=True)
    os.makedirs(os.path.dirname(desktop_docx), exist_ok=True)
    
    print(f"Saving finalized manuscript to: {output_docx}...")
    doc.save(output_docx)
    print(f"Saving finalized manuscript to Desktop: {desktop_docx}...")
    doc.save(desktop_docx)

    # Package Supplementary Materials Zip
    zip_path = "manuscript/Supplementary_Materials.zip"
    desktop_zip = "/Users/dicky/Desktop/pneumonia/Supplementary_Materials.zip"
    print(f"Creating supplementary zip archive: {zip_path}...")
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk("results/tables"):
            for file in files:
                if file.endswith('.csv'):
                    zipf.write(os.path.join(root, file), arcname=os.path.join("tables", file))
        for root, dirs, files in os.walk("figures_jurnal"):
            for file in files:
                if file.endswith('.png'):
                    zipf.write(os.path.join(root, file), arcname=os.path.join("figures", file))
    
    shutil.copy2(zip_path, desktop_zip)
    print("Flawless build v2 completed successfully!")

if __name__ == "__main__":
    build_manuscript()
