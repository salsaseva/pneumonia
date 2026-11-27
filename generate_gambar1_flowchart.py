import matplotlib.pyplot as plt
import matplotlib.patches as patches
import os

OUTPUT_DIR = "figures_jurnal"
os.makedirs(OUTPUT_DIR, exist_ok=True)

plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial', 'Helvetica']

fig, ax = plt.subplots(figsize=(13, 8.5), dpi=300)
ax.set_xlim(0, 13)
ax.set_ylim(0, 9)
ax.axis('off')

# Color palette (Clean Medical Science)
c_header = '#1B365D'      # Deep Navy
c_box_bg = '#F8FAFC'      # Off white / light slate
c_box_border = '#334155'  # Slate gray
c_t1 = '#0284C7'          # Sky Blue (Audit)
c_t2 = '#0D9488'          # Teal (Protocols)
c_t3 = '#D97706'          # Amber (Preprocessing & GLCM)
c_t4 = '#7C3AED'          # Purple (Modeling & Pipelines)
c_t5 = '#2563EB'          # Royal Blue (Ensemble)
c_t6 = '#DC2626'          # Crimson (Evaluation & Ablation)

stages = [
    {
        "title": "Tahap 1: Audit Dataset & Ekstraksi ID Pasien",
        "color": c_t1,
        "x": 0.5, "y": 6.2, "w": 3.6, "h": 2.3,
        "items": [
            "• Total: 5.856 citra radiografi dada pediatrik",
            "• Verifikasi bitwise SHA-256 (30 duplikat, 0 korup)",
            "• Ekstraksi ID Pasien via Regex (personID / NORMAL2)",
            "• Audit Overlap: 170 pasien Pneumonia (84,16%)",
            "  tumpang tindih di uji resmi P2 (Normal: 0%)"
        ]
    },
    {
        "title": "Tahap 2: Pembagian Data Lintas Tiga Protokol",
        "color": c_t2,
        "x": 4.7, "y": 6.2, "w": 3.6, "h": 2.3,
        "items": [
            "• P1: Split Acak Citra 70:10:20 (4.100 / 584 / 1.172)",
            "• P2: Split Resmi Kaggle (4.185 / 1.047 Val / 624 Uji)",
            "• P3: Split Terkontrol Pasien (4.113 / 863 / 880)",
            "• Validasi Tetap 1.047 Citra & 5-Fold Stratified Group CV",
            "• Pengulangan Multi-Seed (5 Run: 42, 123, 456, 789, 2026)"
        ]
    },
    {
        "title": "Tahap 3: Pra-Pengolahan & Ekstraksi Fitur GLCM",
        "color": c_t3,
        "x": 8.9, "y": 6.2, "w": 3.6, "h": 2.3,
        "items": [
            "• Penyeragaman Spasial: 150×150 piksel (cv2.INTER_AREA)",
            "• Mode Kontras: Basic, CLAHE (clip=2.0), Ortiz (adaptif)",
            "• Kuantisasi Gray-Level: Ng ∈ {32, 64, 256}",
            "• GLCM 12 Matriks: Jarak d ∈ {1, 2, 3}, Arah θ ∈ {0°,45°,90°,135°}",
            "• Varian Fitur: GLCM-6 (Rata-rata 4 Arah), GLCM-24, GLCM-5"
        ]
    },
    {
        "title": "Tahap 4: Pipeline Terisolasi & Pemodelan",
        "color": c_t4,
        "x": 0.5, "y": 2.2, "w": 3.6, "h": 2.5,
        "items": [
            "• StandardScaler (Anti-Leak: Fit di Latih, Transform Valid/Uji)",
            "• Penanganan Ketidakseimbangan (74,3% Pneumonia di Latih):",
            "  - class_weight='balanced' (Cost-sensitive)",
            "  - SMOTE (Synthetic Minority Over-sampling)",
            "• Algoritma: SVM (RBF Kernel), Random Forest, KNN",
            "• Penyetelan Hiperparameter (GridSearchCV pada Data Latih)"
        ]
    },
    {
        "title": "Tahap 5: Soft Voting Ensemble Classifier",
        "color": c_t5,
        "x": 4.7, "y": 2.2, "w": 3.6, "h": 2.5,
        "items": [
            "• Integrasi Probabilitas: p(c|x) = (1/M) Σ pm(c|x)",
            "• Kombinasi Model: SVM Tuned + RF Tuned + KNN Tuned",
            "• Evaluasi Threshold Tuning pada Validasi Tetap",
            "• Analisis Korelasi Galat & Matriks Ketidaksepakatan (Q-stat)",
            "• Uji Ekivalensi TOST (Two One-Sided Tests, Margin ±2%)"
        ]
    },
    {
        "title": "Tahap 6: Evaluasi Diagnostik & Analisis Lanjut",
        "color": c_t6,
        "x": 8.9, "y": 2.2, "w": 3.6, "h": 2.5,
        "items": [
            "• Metrik: Akurasi, F1-Score, Sensitivitas, Spesifisitas, MCC, AUC",
            "• Inferensi Statistik: Uji McNemar Berpasangan (α = 0,0083)",
            "• Resampling Paired Bootstrap (B = 1.000, 95% CI)",
            "• Studi Ablasi 27 Kombinasi & Analisis Sensitivitas Prevalensi",
            "• Grouped PFI, Drop-Column, & Uji Masking Tepi Tanpa Resize"
        ]
    }
]

# Title banner
title_rect = patches.FancyBboxPatch((0.5, 8.55), 12.0, 0.45, boxstyle="round,pad=0.02", ec="none", fc=c_header)
ax.add_patch(title_rect)
ax.text(6.5, 8.77, "ALUR KERJA METODOLOGI PENELITIAN SISTEMATIS (6 TAHAP)", color="white", fontsize=11, fontweight="bold", ha="center", va="center")

for st in stages:
    # Outer box
    rect = patches.FancyBboxPatch(
        (st["x"], st["y"]), st["w"], st["h"],
        boxstyle="round,pad=0.04", ec=st["color"], fc=c_box_bg, lw=1.8
    )
    ax.add_patch(rect)
    
    # Header bar
    hbar = patches.FancyBboxPatch(
        (st["x"] + 0.05, st["y"] + st["h"] - 0.42), st["w"] - 0.1, 0.38,
        boxstyle="round,pad=0.02", ec="none", fc=st["color"]
    )
    ax.add_patch(hbar)
    ax.text(
        st["x"] + st["w"] / 2.0, st["y"] + st["h"] - 0.23,
        st["title"], color="white", fontsize=8.5, fontweight="bold", ha="center", va="center"
    )
    
    # Body text
    y_text = st["y"] + st["h"] - 0.65
    for item in st["items"]:
        ax.text(
            st["x"] + 0.15, y_text, item,
            color="#1E293B", fontsize=7.2, va="top", ha="left"
        )
        y_text -= 0.33

# Connectors: Horizontal arrows
# Stage 1 -> Stage 2
ax.annotate("", xy=(4.7, 7.35), xytext=(4.1, 7.35),
            arrowprops=dict(arrowstyle="->", color=c_header, lw=2.0))
# Stage 2 -> Stage 3
ax.annotate("", xy=(8.9, 7.35), xytext=(8.3, 7.35),
            arrowprops=dict(arrowstyle="->", color=c_header, lw=2.0))

# Stage 3 -> Stage 4 (Vertical drop and across)
ax.annotate("", xy=(10.7, 4.7), xytext=(10.7, 6.2),
            arrowprops=dict(arrowstyle="-", color=c_header, lw=2.0))
ax.annotate("", xy=(2.3, 4.7), xytext=(10.7, 4.7),
            arrowprops=dict(arrowstyle="-", color=c_header, lw=2.0))
ax.annotate("", xy=(2.3, 4.7), xytext=(2.3, 4.7),
            arrowprops=dict(arrowstyle="->", color=c_header, lw=2.0))
ax.annotate("", xy=(2.3, 4.7), xytext=(2.3, 5.0),
            arrowprops=dict(arrowstyle="->", color=c_header, lw=2.0))
ax.annotate("", xy=(2.3, 4.7), xytext=(2.3, 4.7))
# Direct arrow from Stage 3 down to Stage 4
ax.annotate("", xy=(2.3, 4.7), xytext=(2.3, 4.7))

# Let's draw clean routing arrows:
# Down from Stage 3, left, down into Stage 4:
path_data = [
    (10.7, 6.2), (10.7, 5.4),
    (2.3, 5.4), (2.3, 4.7)
]
ax.plot([10.7, 10.7], [6.2, 5.4], color=c_header, lw=2.0)
ax.plot([10.7, 2.3], [5.4, 5.4], color=c_header, lw=2.0)
ax.annotate("", xy=(2.3, 4.7), xytext=(2.3, 5.4),
            arrowprops=dict(arrowstyle="->", color=c_header, lw=2.0))

# Stage 4 -> Stage 5
ax.annotate("", xy=(4.7, 3.45), xytext=(4.1, 3.45),
            arrowprops=dict(arrowstyle="->", color=c_header, lw=2.0))
# Stage 5 -> Stage 6
ax.annotate("", xy=(8.9, 3.45), xytext=(8.3, 3.45),
            arrowprops=dict(arrowstyle="->", color=c_header, lw=2.0))

# Bottom banner for Reproducibility & Integrity Guarantee
bot_rect = patches.FancyBboxPatch((0.5, 0.4), 12.0, 1.4, boxstyle="round,pad=0.03", ec="#059669", fc="#ECFDF5", lw=1.5)
ax.add_patch(bot_rect)
ax.text(0.7, 1.55, "JAMINAN INTEGRITAS METODOLOGI & PROTOKOL BENCHMARKING:", color="#065F46", fontsize=8.5, fontweight="bold")
guarantees = [
    "1. Pemisahan Data & Pencegahan Kebocoran: StandardScaler difit secara eksklusif pada data latih masing-masing fold dan protokol (Anti-Leak).",
    "2. Mitigasi Bias Prevalensi & Optimism Gap: Analisis subsampling Monte Carlo 100× menyelaraskan rasio kelas P1 dan P3 ke standar uji resmi P2 (62,5%).",
    "3. Rigoritas Statistik & Bebas Klaim Berlebih: Uji signifikansi McNemar berpasangan (α = 0,0083) dan TOST (margin ±2%) mengonfirmasi tidak ada perbedaan performa signifikan antara Soft Voting dan SVM.",
    "4. Penjelas Fitur Bebas Multikolinieritas: Analisis Grouped PFI dan Drop-Column mengidentifikasi kelompok variasi kontras sebagai prediktor utama tanpa bias korelasi."
]
y_g = 1.3
for g in guarantees:
    ax.text(0.7, y_g, g, color="#047857", fontsize=7.2)
    y_g -= 0.25

plt.tight_layout()
out_fig = os.path.join(OUTPUT_DIR, "gambar1_alur_penelitian.png")
plt.savefig(out_fig, dpi=300, bbox_inches='tight')
print(f"Gambar 1 successfully generated at {out_fig}")
