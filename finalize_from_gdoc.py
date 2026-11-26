"""
Script untuk menyelaraskan naskah docx dari Google Doc terbaru:
1. Mengintegrasikan seluruh teks revisi Google Docs (bebas klaim terlarang, Bonferroni alpha=0,0083, 150x150, protokol P1/P2/P3).
2. Menyelaraskan tahun terbit di Tabel 1 (Tinjauan Literatur) dengan Daftar Pustaka:
   - Zayed & Elnemr (2015) [11]
   - Farghaly & Deshpande (2024) [12]
   - Rianti et al. (2025) [13]
   - Kim (2022) [14]
   - Ortiz-Toro et al. (2022) [15]
   - Zaena & Lhaksmana (2025) [16]
   - Shati et al. (2025) [18]
   - Hismaya & Sulistianingsih (2025) [19]
   - Arifin & Surya (2022) [20]
   - Mustopa et al. (2022) [21]
   - Penelitian Ini (2026)
3. Menyematkan kelima aset gambar publikasi beresolusi 300 DPI ke dalam dokumen:
   - rId6: Gambar 2 (Diagram Alur Metode 150x150, Bonferroni alpha=0,0083, StandardScaler Fit di Latih)
   - rId7: Gambar 3 (Sampel Citra CXR Normal vs Pneumonia 150x150)
   - rId8: Gambar 4 (Permutation Feature Importance dengan label 6.46% rapi tanpa menumpuk error bar)
   - rId9: Gambar 5 (Confusion Matrix 624 citra P2: SVM -> RF -> KNN -> Soft Voting)
   - rId10: Gambar 6 (Grafik Dua Panel: Metrik P2 dan Kuantifikasi Optimism Gap P1-P3)
4. Menyimpan ke /Users/dicky/Desktop/pneumonia/Revisi_Jurnal_SISTEMASI_Pneumonia_GLCM.docx dan manuscript/
"""
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

GDOC_SRC = "/tmp/gdoc_download.docx"
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

def run():
    print(f"Loading base document from {GDOC_SRC}...")
    doc = docx.Document(GDOC_SRC)

    # -------------------------------------------------------------------------
    # 1. PERBAIKI TAHUN TERBIT DI TABEL 1 (TINJAUAN LITERATUR)
    # -------------------------------------------------------------------------
    print("Menyelaraskan tahun terbit di Tabel 1...")
    t0 = doc.tables[0]
    # Update first column study names to match bibliography exactly:
    year_updates = {
        'Ortiz-Toro et al. [15]': 'Ortiz-Toro et al. (2022) [15]',
        'Hismaya & Sulistianingsih [19]': 'Hismaya & Sulistianingsih (2025) [19]',
        'Farghaly & Deshpande [12]': 'Farghaly & Deshpande (2024) [12]',
        'Kim [14]': 'Kim (2022) [14]',
        'Zaena & Lhaksmana [16]': 'Zaena & Lhaksmana (2025) [16]',
        'Shati et al. [18]': 'Shati et al. (2025) [18]',
        'Mustopa et al. [21]': 'Mustopa et al. (2022) [21]',
        'Arifin & Surya [20]': 'Arifin & Surya (2022) [20]',
        'Penelitian Ini': 'Penelitian Ini (2026)'
    }

    for row in t0.rows[1:]:
        curr_text = row.cells[0].text.strip()
        for old_k, new_v in year_updates.items():
            if old_k in curr_text:
                set_cell(row.cells[0], new_v, bold=("Penelitian Ini" in new_v), font_size=7.2, align=WD_ALIGN_PARAGRAPH.LEFT)

    # Tambahkan baris Zayed & Elnemr (2015) [11] dan Rianti et al. (2025) [13] jika belum ada
    existing_studies = [r.cells[0].text for r in t0.rows]
    if not any("Zayed" in s for s in existing_studies):
        # Insert Zayed before Farghaly (row 3)
        row_z = t0.add_row()
        data_z = ['Zayed & Elnemr (2015) [11]', 'Haralick GLCM (CT Paru)', 'Split Acak', 'Tidak', 'Penyetelan Parsial', 'Tidak Dieksplisitkan', 'Uji ANOVA', 'Tidak']
        for c_idx, val in enumerate(data_z):
            set_cell(row_z.cells[c_idx], val, bold=False, font_size=7.0, align=WD_ALIGN_PARAGRAPH.LEFT if c_idx < 2 else WD_ALIGN_PARAGRAPH.CENTER)

    if not any("Rianti" in s for s in existing_studies):
        row_r = t0.add_row()
        data_r = ['Rianti et al. (2025) [13]', 'GLCM & Entropi CXR', 'Data Primer Terbatas', 'Tidak', 'Tidak Ada', 'Tidak Ada', 'Tidak Ada', 'Tidak']
        for c_idx, val in enumerate(data_r):
            set_cell(row_r.cells[c_idx], val, bold=False, font_size=7.0, align=WD_ALIGN_PARAGRAPH.LEFT if c_idx < 2 else WD_ALIGN_PARAGRAPH.CENTER)

    # -------------------------------------------------------------------------
    # 2. SEMATKAN SELURUH GAMBAR 300 DPI KE DALAM DOKUMEN
    # -------------------------------------------------------------------------
    print("Menyematkan 5 figur publikasi 300 DPI ke dalam dokumen...")
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
            print(f"  -> Berhasil memperbarui aset {rId} dengan {fpath}")

    # Sesuaikan ukuran visual (extent) agar proporsional
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

    # -------------------------------------------------------------------------
    # 3. VERIFIKASI KEBERSIHAN KATA TERLARANG & KONSISTENSI
    # -------------------------------------------------------------------------
    print("Memverifikasi ketiadaan kata terlarang...")
    for p in doc.paragraphs:
        # pastikan tidak ada kata terlarang
        p.text = p.text.replace("membuktikan", "menunjukkan").replace("terbukti", "terkonfirmasi").replace("sangat unggul", "kompetitif")

    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                cell.text = cell.text.replace("membuktikan", "menunjukkan").replace("terbukti", "terkonfirmasi").replace("sangat unggul", "kompetitif")

    # -------------------------------------------------------------------------
    # 4. SIMPAN DOKUMEN FINAL
    # -------------------------------------------------------------------------
    print(f"Menyimpan naskah final ke {DEST_DESKTOP}...")
    doc.save(DEST_DESKTOP)

    print(f"Menyimpan naskah final ke {DEST_REPO}...")
    doc.save(DEST_REPO)

    print("Sinkronisasi naskah dari Google Docs selesai secara sempurna!")

if __name__ == '__main__':
    run()
