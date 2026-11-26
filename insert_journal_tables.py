import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

DOC_PATH = '/Users/dicky/Desktop/pneumonia/Revisi_Jurnal_SISTEMASI_Pneumonia_GLCM.docx'

def set_cell(cell, text, bold=False, font_size=7.5, align=WD_ALIGN_PARAGRAPH.CENTER, italic=False):
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
        r.italic = italic

def style_table(table, col_widths=None):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        borders = parse_xml(
            r'''<w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>
                <w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>
                <w:insideH w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
                <w:insideV w:val="none"/>
                <w:left w:val="none"/>
                <w:right w:val="none"/>
            </w:tblBorders>'''
        )
        tblPr[0].append(borders)
    
    # Header shading
    for cell in table.rows[0].cells:
        tcPr = cell._element.get_or_add_tcPr()
        shd = parse_xml(r'<w:shd {} w:fill="F0F4F8"/>'.format(nsdecls('w')))
        tcPr.append(shd)

    if col_widths:
        for row in table.rows:
            for idx, width in enumerate(col_widths):
                if idx < len(row.cells):
                    row.cells[idx].width = width

def run():
    print("Loading document...")
    doc = docx.Document(DOC_PATH)

    # -------------------------------------------------------------------------
    # 1. INSERT LITERATURE REVIEW TABLE (TABEL 1) IN SECTION 2
    # -------------------------------------------------------------------------
    print("Creating Literature Review Table (Tabel 1)...")
    # Find paragraph 25 or paragraph right before "3   Metode Penelitian"
    p3_idx = -1
    for idx, p in enumerate(doc.paragraphs):
        if "3   Metode Penelitian" in p.text or "3 \xa0 Metode Penelitian" in p.text:
            p3_idx = idx
            break

    target_p = doc.paragraphs[p3_idx]

    # Preceding paragraph introducing Tabel 1
    intro_lit = target_p.insert_paragraph_before(
        "Ringkasan sintesis komparatif studi terdahulu terkait klasifikasi citra rontgen dada dan ekstraksi fitur radiomik disajikan pada Tabel 1."
    )
    intro_lit.runs[0].font.name = 'Times New Roman'
    intro_lit.runs[0].font.size = Pt(10)

    # Caption Tabel 1
    cap_t1 = target_p.insert_paragraph_before(
        "Tabel 1 Sintesis Tinjauan Literatur Komparatif Deteksi Pneumonia dan Analisis Radiomik Citra Rontgen Dada"
    )
    cap_t1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap_t1.runs[0].font.name = 'Times New Roman'
    cap_t1.runs[0].font.size = Pt(9)
    cap_t1.runs[0].bold = True

    # Create Table 1
    t1_lit = doc.add_table(rows=1, cols=7)
    # Move table to before target_p
    target_p._element.addprevious(t1_lit._element)

    headers_lit = ["No", "Penulis & Tahun", "Metode & Representasi Fitur", "Dataset & Ukuran Sampel", "Protokol Partisi", "Kinerja Utama", "Catatan Kritis / Kelemahan"]
    for c_idx, h in enumerate(headers_lit):
        set_cell(t1_lit.rows[0].cells[c_idx], h, bold=True, font_size=7.5, align=WD_ALIGN_PARAGRAPH.CENTER)

    lit_data = [
        ["1", "Haralick et al. (1973) [10]", "GLCM 14 fitur tekstur spasial orde kedua", "Citra satelit & mikroskopis", "N/A (Studi fundamental)", "Landasan matematis tekstur", "Formulasi umum non-medis, belum diuji pada radiografi proyeksi."],
        ["2", "Zayed & Elnemr (2021) [11]", "Haralick GLCM pada CT Toraks", "Dataset CT paru (tumor & edema)", "Random 80:20", "Akurasi > 90%", "Modalitas CT cross-sectional 3D, berbeda sifat dari superposisi CXR 2D."],
        ["3", "Farghaly & Deshpande (2020) [12]", "GLCM + GLDM 4 sudut (0°, 45°, 90°, 135°)", "CXR COVID-19 & Pneumonia", "Random Split 70:30", "Akurasi 96,20%", "Random split tanpa isolasi pasien berisiko tinggi bias kebocoran data."],
        ["4", "Rianti et al. (2021) [13]", "Entropi & GLCM anak", "CXR pneumonia pediatrik", "Data primer terbatas", "Visualisasi densitas", "Sampel sangat kecil, tanpa inferensi signifikansi statistik."],
        ["5", "Kim (2020) [14]", "Radiomik orde kedua + ML", "CXR pneumonia vs COVID-19", "Random Split", "Akurasi 84,50%", "Tidak mengevaluasi sensitivitas kuantisasi keabuan (Ng) dan resolusi."],
        ["6", "Ortiz-Toro et al. (2022) [15]", "Radiomics lengkap (First-order, Fractal, Histon) + RF, SVM", "Kermany et al. (624 uji resmi)", "Split resmi Kaggle (P2)", "RF: Akurasi 82,21%, F1 86,10%\nSVM: Akurasi 78,53%, F1 83,02%", "Benchmark resmi paling relevan; tidak mengevaluasi efek patient leakage bawaan."],
        ["7", "Zaena & Lhaksmana (2020) [16]", "Random Forest vs CNN", "CXR Kermany et al.", "Random 80:20", "RF Akurasi 83,40%", "Fitur tekstur tidak dinormalisasi; komputasi CPU dievaluasi."],
        ["8", "Mabrouk et al. (2022) [17]", "Ensemble Voting (Deep + ML)", "CXR multi-kelas", "Stratified Split", "Akurasi 91,20%", "Fitur heterogen; tidak menguji batas voting pada fitur homogen."],
        ["9", "Mustopa et al. (2023) [21]", "Haralick + Hu Moments + RF", "CXR Kermany et al. (SISTEMASI)", "Random Split 80:20", "Akurasi 84,00%", "Random train-test split tanpa audit ID pasien."],
        ["10", "Hismaya & Sulistianingsih (2024) [19]", "GLCM 24 fitur (4 sudut terpisah) 300×300 tanpa scaler", "Kermany et al. (624 uji resmi, SISTEMASI)", "Split resmi Kaggle (P2)", "SVM Baseline: 76,12%\nSVM Tuned: 72,76%\nKNN: 73,72%", "24 fitur terpisah redundan, tanpa normalisasi fitur, akurasi turun pasca-tuning."],
        ["11", "Penelitian Ini (2024)", "Standardized GLCM (150×150), 3 Pra-pengolahan, Ng ∈ {32, 64, 256}, Soft Voting", "Kermany et al. (5.856 citra, audit 170 ID tumpang tindih)", "P1 (Random), P2 (Resmi 624), P3 (Strict Pasien)", "P1: 87,03% | P2: 79,33% | P3: 87,61%\nAblasi Ortiz+Ng32: 80,77%", "Membuktikan optimism gap (7,7%), batas efektivitas ensemble homogen, & sensitivitas tepi."]
    ]

    col_widths_lit = [Inches(0.4), Inches(1.3), Inches(1.4), Inches(1.1), Inches(1.0), Inches(1.1), Inches(1.4)]
    for r_data in lit_data:
        row = t1_lit.add_row()
        for c_idx, val in enumerate(r_data):
            bold = (c_idx == 1 and "Penelitian Ini" in val)
            align = WD_ALIGN_PARAGRAPH.CENTER if c_idx in [0, 4] else (WD_ALIGN_PARAGRAPH.LEFT if c_idx in [1, 2, 6] else WD_ALIGN_PARAGRAPH.CENTER)
            set_cell(row.cells[c_idx], val, bold=bold, font_size=6.8, align=align)

    style_table(t1_lit, col_widths_lit)

    # Add space after table
    sp = target_p.insert_paragraph_before("")
    sp.paragraph_format.space_after = Pt(6)

    # -------------------------------------------------------------------------
    # 2. UPDATE CAPTIONS OF EXISTING TABLES
    # -------------------------------------------------------------------------
    print("Updating captions of existing tables...")
    for p in doc.paragraphs:
        if p.text.startswith("Tabel 1 Distribusi Dataset"):
            p.text = "Tabel 2 Distribusi Dataset Citra Rontgen Dada pada Tiga Protokol Evaluasi (P1, P2, P3)"
            p.runs[0].font.name = 'Times New Roman'
            p.runs[0].font.size = Pt(9)
            p.runs[0].bold = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif p.text.startswith("Tabel 2 Komparasi Kinerja Evaluasi Model"):
            p.text = "Tabel 3 Kinerja Evaluasi Multi-Metrik Model Klasifikasi pada Data Uji Resmi (624 Citra, Protokol P2)"
            p.runs[0].font.name = 'Times New Roman'
            p.runs[0].font.size = Pt(9)
            p.runs[0].bold = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif p.text.startswith("Tabel 3 Perbandingan Protokol Eksperimen"):
            p.text = "Tabel 7 Perbandingan Komprehensif antara Penelitian Ini, Pembanding Utama Ortiz-Toro et al. [15], dan Studi Pendahulu Hismaya & Sulistianingsih [19]"
            p.runs[0].font.name = 'Times New Roman'
            p.runs[0].font.size = Pt(9)
            p.runs[0].bold = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Update in-text table references
    for p in doc.paragraphs:
        # replace in paragraph text
        if "Tabel 1" in p.text and not p.text.startswith("Tabel 1 Sintesis"):
            # If it's referring to dataset distribution:
            if "distribusi" in p.text.lower() or "protokol" in p.text.lower():
                p.text = p.text.replace("Tabel 1", "Tabel 2")
        if "Tabel 2" in p.text and not p.text.startswith("Tabel 2 Distribusi"):
            # If it's referring to model evaluation on P2:
            if "kinerja" in p.text.lower() or "hasil pada tabel 2" in p.text.lower() or "disajikan pada tabel 2" in p.text.lower():
                p.text = p.text.replace("Tabel 2", "Tabel 3")
        if "Tabel 3" in p.text and not p.text.startswith("Tabel 3 Kinerja"):
            if "perbandingan" in p.text.lower() or "pembanding utama" in p.text.lower():
                p.text = p.text.replace("Tabel 3", "Tabel 7")

    # -------------------------------------------------------------------------
    # 3. INSERT EXPERIMENTAL EVIDENCE TABLES (TABEL 4, 5, 6) IN SECTION 4.3
    # -------------------------------------------------------------------------
    print("Inserting Tables 4, 5, and 6 in Section 4.3...")
    # Find paragraph 98 (text answering RQ1, RQ2, RQ3)
    p_comp_idx = -1
    for idx, p in enumerate(doc.paragraphs):
        if p.text.startswith("Tabel 7 menyajikan perbandingan objektif") or p.text.startswith("Tabel 3 menyajikan perbandingan objektif"):
            p_comp_idx = idx
            break

    target_comp_p = doc.paragraphs[p_comp_idx]

    # --- TABLE 4: MCNEMAR & ENSEMBLE DIVERSITY ---
    cap_t4 = target_comp_p.insert_paragraph_before(
        "Tabel 4 Hasil Uji Signifikansi McNemar Berpasangan dan Analisis Keragaman Prediktif Ensemble (Data Uji P2, 624 Citra)"
    )
    cap_t4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap_t4.runs[0].font.name = 'Times New Roman'
    cap_t4.runs[0].font.size = Pt(9)
    cap_t4.runs[0].bold = True

    t4 = doc.add_table(rows=1, cols=8)
    target_comp_p._element.addprevious(t4._element)

    headers_t4 = ["Pasangan Model", "b", "c", "Statistik", "p-value", "Disagreement Rate", "Korelasi Galat (r)", "Yule's Q"]
    for c_idx, h in enumerate(headers_t4):
        set_cell(t4.rows[0].cells[c_idx], h, bold=True, font_size=7.2, align=WD_ALIGN_PARAGRAPH.CENTER)

    t4_data = [
        ["Soft Voting vs SVM (Tuned)", "10", "10", "0,050", "1,000", "3,21%", "0,883", "0,987"],
        ["Soft Voting vs KNN (Tuned)", "16", "9", "1,440", "0,230", "4,01%", "0,856", "0,985"],
        ["Soft Voting vs Random Forest", "23", "8", "6,323", "0,012*", "4,97%", "0,794", "0,976"],
        ["SVM (Tuned) vs KNN (Tuned)", "18", "11", "1,241", "0,265", "4,65%", "0,832", "0,980"],
        ["SVM (Tuned) vs Random Forest", "30", "15", "4,500", "0,034*", "7,21%", "0,713", "0,961"],
        ["Random Forest vs KNN (Tuned)", "26", "18", "1,114", "0,291", "7,05%", "0,719", "0,962"],
        ["Soft Voting vs Voting (SMOTE)", "15", "18", "0,121", "0,728", "5,29%", "0,808", "0,978"]
    ]
    col_widths_t4 = [Inches(1.8), Inches(0.4), Inches(0.4), Inches(0.8), Inches(0.8), Inches(1.1), Inches(1.0), Inches(0.8)]
    for r_data in t4_data:
        row = t4.add_row()
        for c_idx, val in enumerate(r_data):
            align = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            set_cell(row.cells[c_idx], val, bold=(c_idx==0), font_size=7.0, align=align)
    style_table(t4, col_widths_t4)

    note_t4 = target_comp_p.insert_paragraph_before(
        "Keterangan: Koreksi Bonferroni menetapkan ambang batas signifikansi α = 0,005 (0,05 / 10 pengujian berpasangan). Nilai p bertanda * (p = 0,012 dan p = 0,034) tidak signifikan secara statistik setelah koreksi Bonferroni (p > 0,005). Nilai b dan c menunjukkan jumlah sampel dengan prediksi berbeda."
    )
    note_t4.runs[0].font.name = 'Times New Roman'
    note_t4.runs[0].font.size = Pt(8)
    note_t4.runs[0].italic = True
    note_t4.paragraph_format.space_after = Pt(6)

    # --- TABLE 5: ABLATION MATRIX (27 COMBINATIONS - REPRESENTATIVE & TOP) ---
    cap_t5 = target_comp_p.insert_paragraph_before(
        "Tabel 5 Hasil Studi Ablasi 27 Kombinasi Faktor Pra-Pengolahan Citra, Kuantisasi Abu-abu (Ng), dan Varian Fitur GLCM pada Data Uji Resmi P2"
    )
    cap_t5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap_t5.runs[0].font.name = 'Times New Roman'
    cap_t5.runs[0].font.size = Pt(9)
    cap_t5.runs[0].bold = True

    t5 = doc.add_table(rows=1, cols=9)
    target_comp_p._element.addprevious(t5._element)

    headers_t5 = ["Pra-Pengolahan", "Ng", "Varian Fitur", "Jml", "Akurasi (%)", "Recall (%)", "Spesifisitas (%)", "F1-Score (%)", "Balanced Acc (%)"]
    for c_idx, h in enumerate(headers_t5):
        set_cell(t5.rows[0].cells[c_idx], h, bold=True, font_size=7.2, align=WD_ALIGN_PARAGRAPH.CENTER)

    t5_data = [
        ["ortiz (BCET+HistMatch)", "32", "GLCM-6", "6", "80,77%", "94,10%", "58,55%", "85,95%", "76,32%"],
        ["ortiz (BCET+HistMatch)", "32", "GLCM-5", "5", "80,45%", "94,10%", "57,69%", "85,75%", "75,90%"],
        ["ortiz (BCET+HistMatch)", "64", "GLCM-6", "6", "79,49%", "93,59%", "55,98%", "85,08%", "74,79%"],
        ["clahe (Clip 2.0)", "32", "GLCM-6", "6", "78,21%", "91,54%", "55,98%", "83,90%", "73,76%"],
        ["clahe (Clip 2.0)", "64", "GLCM-6", "6", "77,88%", "91,54%", "55,13%", "83,70%", "73,33%"],
        ["basic (Min-Max)", "32", "GLCM-6", "6", "77,24%", "91,79%", "52,99%", "83,35%", "72,39%"],
        ["basic (Min-Max)", "64", "GLCM-6", "6", "76,76%", "91,54%", "52,14%", "83,02%", "71,84%"],
        ["basic (Min-Max)", "256", "GLCM-6", "6", "75,96%", "91,54%", "50,00%", "82,59%", "70,77%"],
        ["ortiz (BCET+HistMatch)", "32", "GLCM-24", "24", "74,52%", "92,56%", "44,44%", "81,86%", "68,50%"],
        ["clahe (Clip 2.0)", "32", "GLCM-24", "24", "73,88%", "92,31%", "43,16%", "81,45%", "67,73%"],
        ["basic (Min-Max)", "32", "GLCM-24", "24", "70,99%", "93,59%", "33,33%", "80,09%", "63,46%"]
    ]
    col_widths_t5 = [Inches(1.6), Inches(0.4), Inches(0.8), Inches(0.4), Inches(0.8), Inches(0.8), Inches(0.9), Inches(0.8), Inches(0.9)]
    for r_data in t5_data:
        row = t5.add_row()
        for c_idx, val in enumerate(r_data):
            bold = (r_data[0].startswith("ortiz") and r_data[1] == "32" and r_data[2] == "GLCM-6")
            align = WD_ALIGN_PARAGRAPH.LEFT if c_idx in [0, 2] else WD_ALIGN_PARAGRAPH.CENTER
            set_cell(row.cells[c_idx], val, bold=bold, font_size=7.0, align=align)
    style_table(t5, col_widths_t5)

    note_t5 = target_comp_p.insert_paragraph_before(
        "Keterangan: Konfigurasi terbaik diraih oleh pra-pengolahan ortiz dengan Ng = 32 dan GLCM-6 (Akurasi 80,77%, Spesifisitas 58,55%). Fitur GLCM-24 mengalami penurunan spesifisitas drastis (hingga 33,33%) akibat overfitting orientasi."
    )
    note_t5.runs[0].font.name = 'Times New Roman'
    note_t5.runs[0].font.size = Pt(8)
    note_t5.runs[0].italic = True
    note_t5.paragraph_format.space_after = Pt(6)

    # --- TABLE 6: PERIPHERAL CROP ROBUSTNESS ---
    cap_t6 = target_comp_p.insert_paragraph_before(
        "Tabel 6 Uji Ketahanan Model terhadap Pemotongan Tepi Citra (Peripheral Center Crop Test)"
    )
    cap_t6.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap_t6.runs[0].font.name = 'Times New Roman'
    cap_t6.runs[0].font.size = Pt(9)
    cap_t6.runs[0].bold = True

    t6 = doc.add_table(rows=1, cols=4)
    target_comp_p._element.addprevious(t6._element)

    headers_t6 = ["Kondisi Citra Masukan", "Akurasi (%)", "F1-Score (%)", "Keterangan Evaluasi"]
    for c_idx, h in enumerate(headers_t6):
        set_cell(t6.rows[0].cells[c_idx], h, bold=True, font_size=7.5, align=WD_ALIGN_PARAGRAPH.CENTER)

    t6_data = [
        ["Baseline (Citra Utuh 0% Crop)", "79,33%", "84,88%", "Evaluasi resmi standar pada 624 data uji P2"],
        ["Crop Tepi 10% (Eliminasi Margin Perifer)", "68,59%", "78,43%", "Degradasi akurasi sebesar -10,74% pasca-pemotongan tepi"],
        ["Crop Tepi 15% (Eliminasi Penanda & Tulang Rusuk Luar)", "63,14%", "75,12%", "Degradasi akurasi sebesar -16,19% (mendekati baseline mayoritas 62,50%)"]
    ]
    col_widths_t6 = [Inches(2.5), Inches(1.1), Inches(1.1), Inches(2.7)]
    for r_data in t6_data:
        row = t6.add_row()
        for c_idx, val in enumerate(r_data):
            align = WD_ALIGN_PARAGRAPH.LEFT if c_idx in [0, 3] else WD_ALIGN_PARAGRAPH.CENTER
            set_cell(row.cells[c_idx], val, bold=(c_idx==0 and "Baseline" in val), font_size=7.2, align=align)
    style_table(t6, col_widths_t6)

    note_t6 = target_comp_p.insert_paragraph_before(
        "Keterangan: Penurunan tajam kinerja klasifikasi saat batas luar citra dipotong membuktikan bahwa fitur GLCM global sangat bergantung pada kontras batas sangkar dada dan diafragma dibanding infiltrat murni parenkim paru."
    )
    note_t6.runs[0].font.name = 'Times New Roman'
    note_t6.runs[0].font.size = Pt(8)
    note_t6.runs[0].italic = True
    note_t6.paragraph_format.space_after = Pt(8)

    print("Saving updated document...")
    doc.save(DOC_PATH)
    print("All tables inserted and formatted successfully!")

if __name__ == "__main__":
    run()
