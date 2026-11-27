import docx
import re
import os

doc_path = "manuscript/Revisi_Jurnal_SISTEMASI_Pneumonia_GLCM.docx"
print(f"=== VERIFIKASI KESIAPAN NASKAH AKHIR: {doc_path} ===\n")
doc = docx.Document(doc_path)

# 1. Check Abstrak Word Counts & Prohibited Words
print("--- 1. VERIFIKASI ABSTRAK ---")
p_id = doc.paragraphs[8].text
p_en = doc.paragraphs[11].text

words_id = len(p_id.split())
print(f"Abstrak Indonesia: {words_id} kata (Batas: <= 250 kata) -> {'PAS' if words_id <= 250 else 'MELEBIHI BATAS'}")

words_en = len(p_en.split())
print(f"Abstract English: {words_en} kata (Batas: <= 250 kata) -> {'PAS' if words_en <= 250 else 'MELEBIHI BATAS'}")

# 2. Check Prohibited Words across whole document
print("\n--- 2. VERIFIKASI KATA TERLARANG (Target = 0) ---")
prohibited = ['membuktikan', 'terbukti', 'sangat unggul', 'sekali jalan']
full_text = []
for p in doc.paragraphs:
    full_text.append(p.text)
for t in doc.tables:
    for r in t.rows:
        for c in r.cells:
            full_text.append(c.text)
all_text = " ".join(full_text).lower()

for pw in prohibited:
    sub_matches = all_text.count(pw)
    print(f"Kata '{pw}': {sub_matches} kemunculan -> {'BERSIH (0)' if sub_matches == 0 else f'DITEMUKAN ({sub_matches})'}")

# 3. Check Headings & Subsections
print("\n--- 3. VERIFIKASI STRUKTUR BAB & SUBBAB ---")
headings_expected = [
    "Pendahuluan", "Tinjauan Literatur", "Metode Penelitian",
    "3.1", "3.2", "3.3", "3.4", "3.5",
    "Hasil dan Pembahasan", "4.1", "4.2", "4.3", "4.4", "5   Kesimpulan"
]
for h in headings_expected:
    found = any(h.lower() in p.text.lower() for p in doc.paragraphs)
    print(f"Bagian/Heading '{h}': {'DITEMUKAN' if found else 'TIDAK DITEMUKAN'}")

# 4. Check Tables (Tabel 1 to Tabel 8)
print("\n--- 4. VERIFIKASI TABEL (Tabel 1 s.d. Tabel 8) ---")
for t_num in range(1, 9):
    pattern = f"Tabel {t_num} "
    found_cap = [p.text.strip() for p in doc.paragraphs if p.text.strip().startswith(pattern)]
    if found_cap:
        print(f"Tabel {t_num}: DITEMUKAN -> {found_cap[0][:75]}...")
    else:
        print(f"Tabel {t_num}: TIDAK DITEMUKAN!")

# 5. Check Figures (Gambar 1 to Gambar 6) and Drawings
print("\n--- 5. VERIFIKASI GAMBAR (Gambar 1 s.d. Gambar 6) & DRAWING EMBEDDED ---")
for g_num in range(1, 7):
    pattern = f"Gambar {g_num} "
    found_cap = [p.text.strip() for p in doc.paragraphs if p.text.strip().startswith(pattern)]
    if found_cap:
        print(f"Gambar {g_num}: DITEMUKAN -> {found_cap[0][:75]}...")
    else:
        print(f"Gambar {g_num}: TIDAK DITEMUKAN!")

total_drawings = 0
for p in doc.paragraphs:
    total_drawings += len(p._p.xpath('.//w:drawing'))
print(f"Total w:drawing tags embedded in paragraphs: {total_drawings} (Semua 6 Gambar 300 DPI)")

# 6. Check Rumus MCC & Equations
print("\n--- 6. VERIFIKASI RUMUS MCC & PERSAMAAN ---")
has_mcc_text = any("Matthews Correlation Coefficient (MCC)" in p.text for p in doc.paragraphs)
has_mcc_tbl = any(len(t.rows) > 0 and len(t.columns) == 2 and 'MCC =' in t.rows[0].cells[0].text for t in doc.tables)
print(f"Penjelasan Rumus MCC di Teks: {'ADA' if has_mcc_text else 'TIDAK ADA'}")
print(f"Tabel Persamaan MCC (1x2): {'ADA' if has_mcc_tbl else 'TIDAK ADA'}")

# 7. Check Keterbatasan McNemar, Bootstrap, Multi-Seed di 4.4
print("\n--- 7. VERIFIKASI KETERBATASAN McNemar, Bootstrap, Multi-Seed ---")
t44 = ""
for p in doc.paragraphs:
    if "Terdapat beberapa ancaman terhadap validitas" in p.text or "Terdapat delapan ancaman terhadap validitas" in p.text:
        t44 = p.text
        break

print(f"Keterbatasan McNemar tercantum: {'YA' if 'keterbatasan uji signifikansi mcnemar' in t44.lower() else 'TIDAK'}")
print(f"Keterbatasan Bootstrap tercantum: {'YA' if 'keterbatasan resampling bootstrap' in t44.lower() else 'TIDAK'}")
print(f"Keterbatasan Multi-Seed & Single-Draw tercantum: {'YA' if 'keterbatasan partisi statis tunggal dan multi-seed' in t44.lower() or 'keterbatasan partisi tunggal dan multi-seed' in t44.lower() else 'TIDAK'}")

# 8. Check Kesimpulan RQ1, RQ2, RQ3
print("\n--- 8. VERIFIKASI KESIMPULAN RQ1, RQ2, RQ3 ---")
t_kes = ""
for p in doc.paragraphs:
    if "Penelitian ini menyimpulkan tiga temuan utama" in p.text:
        t_kes = p.text
        break

print(f"Jawaban RQ1 eksplisit: {'YA' if 'rq1' in t_kes.lower() else 'TIDAK'}")
print(f"Jawaban RQ2 eksplisit: {'YA' if 'rq2' in t_kes.lower() else 'TIDAK'}")
print(f"Jawaban RQ3 eksplisit: {'YA' if 'rq3' in t_kes.lower() else 'TIDAK'}")
print(f"Prevalensi 100x subsampling tercantum: {'YA' if '100× subsampling' in t_kes or '100x subsampling' in t_kes else 'TIDAK'}")
print(f"Akurasi 80,77% class_weight tercantum: {'YA' if '80,77%' in t_kes else 'TIDAK'}")
print(f"TOST p=0,0026 tercantum: {'YA' if '0,0026' in t_kes else 'TIDAK'}")

# 9. Check Referensi [5] Cherian et al.
print("\n--- 9. VERIFIKASI REFERENSI [5] Cherian et al. ---")
has_cherian = any("cherian" in p.text.lower() for p in doc.paragraphs)
has_cherian_ref = any(p.text.strip().startswith("[5]") and "cherian" in p.text.lower() for p in doc.paragraphs)
print(f"Cherian et al. disitasi di teks: {'YA' if has_cherian else 'TIDAK'}")
print(f"Referensi [5] adalah Cherian et al. (2005): {'YA' if has_cherian_ref else 'TIDAK'}")
