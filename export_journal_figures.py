import os
import cv2
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

OUTPUT_DIR = "figures_jurnal"
os.makedirs(OUTPUT_DIR, exist_ok=True)

plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.size'] = 11
plt.rcParams['axes.labelsize'] = 12
plt.rcParams['axes.titlesize'] = 12
plt.rcParams['xtick.labelsize'] = 10
plt.rcParams['ytick.labelsize'] = 10

# ==============================================================================
# 1. GENERATE GAMBAR 1: PERBANDINGAN SAMPEL CITRA X-RAY (NORMAL VS PNEUMONIA)
# ==============================================================================
print("[1/5] Mengenerate Gambar 1: Sampel Citra X-Ray (Normal vs Pneumonia)...")

dataset_base = "/Users/dicky/Downloads/chest_xray"
normal_path = os.path.join(dataset_base, "test/NORMAL/IM-0031-0001.jpeg")
pneumonia_path = os.path.join(dataset_base, "test/PNEUMONIA/person100_bacteria_482.jpeg")

if not os.path.exists(normal_path):
    candidates = [f for f in os.listdir(os.path.join(dataset_base, "train/NORMAL")) if f.endswith(('.jpg', '.jpeg', '.png'))]
    normal_path = os.path.join(dataset_base, "train/NORMAL", candidates[0])
if not os.path.exists(pneumonia_path):
    candidates = [f for f in os.listdir(os.path.join(dataset_base, "train/PNEUMONIA")) if f.endswith(('.jpg', '.jpeg', '.png'))]
    pneumonia_path = os.path.join(dataset_base, "train/PNEUMONIA", candidates[0])

img_normal_raw = cv2.imread(normal_path)
img_normal_rgb = cv2.cvtColor(img_normal_raw, cv2.COLOR_BGR2RGB)
img_normal_gray = cv2.cvtColor(img_normal_raw, cv2.COLOR_BGR2GRAY)
img_normal_resized = cv2.resize(img_normal_gray, (150, 150), interpolation=cv2.INTER_AREA)

img_pneu_raw = cv2.imread(pneumonia_path)
img_pneu_rgb = cv2.cvtColor(img_pneu_raw, cv2.COLOR_BGR2RGB)
img_pneu_gray = cv2.cvtColor(img_pneu_raw, cv2.COLOR_BGR2GRAY)
img_pneu_resized = cv2.resize(img_pneu_gray, (150, 150), interpolation=cv2.INTER_AREA)

fig, axes = plt.subplots(2, 2, figsize=(8, 8))

axes[0, 0].imshow(img_normal_rgb)
axes[0, 0].set_title("(a) Normal - Citra Asli", pad=8, weight='bold')
axes[0, 0].axis('off')

axes[0, 1].imshow(img_normal_resized, cmap='gray')
axes[0, 1].set_title("(b) Normal - Grayscale (150x150)", pad=8, weight='bold')
axes[0, 1].axis('off')

axes[1, 0].imshow(img_pneu_rgb)
axes[1, 0].set_title("(c) Pneumonia - Citra Asli", pad=8, weight='bold')
axes[1, 0].axis('off')

axes[1, 1].imshow(img_pneu_resized, cmap='gray')
axes[1, 1].set_title("(d) Pneumonia - Grayscale (150x150)", pad=8, weight='bold')
axes[1, 1].axis('off')

plt.tight_layout()
fig1_path = os.path.join(OUTPUT_DIR, "gambar1_sampel_xray.png")
plt.savefig(fig1_path, dpi=300, bbox_inches='tight')
plt.close()
print(f"   -> Tersimpan di: {fig1_path}")

# ==============================================================================
# 2. GENERATE GAMBAR 2: CONFUSION MATRIX (REVISI HASIL RESMI P2)
# ==============================================================================
print("[2/5] Mengenerate Gambar 2: Confusion Matrix Evaluasi Resmi P2...")

df_metrics = pd.read_csv("results/tables/tabel_evaluasi_model_lengkap.csv")
models_focus = ['KNN (Tuned)', 'SVM (Tuned)', 'Random Forest (Tuned)', 'Ensemble Soft Voting']
df_focus = df_metrics[df_metrics['Model'].isin(models_focus)].reset_index(drop=True)

fig, axes = plt.subplots(1, 4, figsize=(16, 4))
class_labels = ['Normal', 'Pneumonia']

for ax, (_, row) in zip(axes, df_focus.iterrows()):
    cm = np.array([
        [int(row['TN']), int(row['FP'])],
        [int(row['FN']), int(row['TP'])]
    ])
    sns.heatmap(cm, annot=True, fmt='d', cmap='Blues', cbar=False,
                xticklabels=class_labels, yticklabels=class_labels,
                annot_kws={"size": 13, "weight": "bold"}, ax=ax)
    
    clean_name = row['Model'].replace(' (Tuned)', '')
    ax.set_title(f"{clean_name}\n(Acc: {row['Accuracy (%)']}%)", fontsize=11, weight='bold', pad=8)
    ax.set_ylabel('True Class', fontsize=10)
    ax.set_xlabel('Predicted Class', fontsize=10)

plt.tight_layout()
fig2_path = os.path.join(OUTPUT_DIR, "gambar2_confusion_matrix.png")
plt.savefig(fig2_path, dpi=300, bbox_inches='tight')
plt.close()
print(f"   -> Tersimpan di: {fig2_path}")

# ==============================================================================
# 3. GENERATE GAMBAR 3: PERMUTATION FEATURE IMPORTANCE (TANPA BIAS GINI)
# ==============================================================================
print("[3/5] Mengenerate Gambar 3: Permutation Feature Importance...")

df_perm = pd.read_csv("results/tables/permutation_feature_importance.csv")
df_perm_sorted = df_perm.sort_values(by='Permutation Importance Mean', ascending=True)

fig, ax = plt.subplots(figsize=(8, 4.5))
bars = ax.barh(
    df_perm_sorted['Fitur'],
    df_perm_sorted['Permutation Importance Mean'] * 100,
    xerr=df_perm_sorted['Permutation Importance Std'] * 100,
    color='#2b5c8f', edgecolor='black', alpha=0.85, capsize=4
)

for bar in bars:
    w = bar.get_width()
    ax.annotate(f"{w:.2f}%",
                xy=(w, bar.get_y() + bar.get_height()/2),
                xytext=(5, 0), textcoords="offset points",
                ha='left', va='center', fontsize=9.5, weight='bold')

ax.set_xlabel('Penurunan F1-Score (%) saat Fitur Dipermutasi', fontsize=11)
ax.set_title('Kontribusi Relatif Fitur GLCM Berdasarkan Permutation Importance\n(Evaluasi Bebas Bias Multikolinieritas)', fontsize=12, weight='bold', pad=12)
ax.grid(axis='x', linestyle='--', alpha=0.4)
ax.set_xlim(0, max(df_perm_sorted['Permutation Importance Mean'] * 100) * 1.25)

plt.tight_layout()
fig3_path = os.path.join(OUTPUT_DIR, "gambar3_feature_importance.png")
plt.savefig(fig3_path, dpi=300, bbox_inches='tight')
plt.close()
print(f"   -> Tersimpan di: {fig3_path}")

# ==============================================================================
# 4. GENERATE GAMBAR 4: GROUPED BAR CHART 6-METRIK DIAGNOSTIK
# ==============================================================================
print("[4/5] Mengenerate Gambar 4: Komparasi 6-Metrik Evaluasi...")

fig, ax = plt.subplots(figsize=(11, 5.5))
x = np.arange(len(df_focus))
metric_specs = [
    ('Accuracy (%)', 'Accuracy', '#2b5c8f'),
    ('Precision/PPV (%)', 'Precision (PPV)', '#41924b'),
    ('Recall/Sen (%)', 'Recall (Sensitivity)', '#e26d2d'),
    ('Specificity (%)', 'Specificity', '#7b5ea7'),
    ('F1-Score (%)', 'F1-Score', '#8c564b'),
    ('Balanced Accuracy (%)', 'Balanced Accuracy', '#7f7f7f'),
]

width = 0.13
offsets = (np.arange(len(metric_specs)) - (len(metric_specs) - 1) / 2) * width
all_rects = []
for (col, label, color), off in zip(metric_specs, offsets):
    all_rects.append(ax.bar(x + off, df_focus[col], width, label=label, color=color))

ax.set_ylabel('Nilai Persentase (%)', fontsize=11)
ax.set_title('Perbandingan Kinerja Diagnostik Model pada Uji Resmi 624 Citra (Protokol P2)', fontsize=13, weight='bold', pad=12)
ax.set_xticks(x)
ax.set_xticklabels([m.replace(' (Tuned)', '') for m in df_focus['Model']], fontsize=11, weight='bold')
ax.set_ylim(45, 100)
ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.08), ncol=6, frameon=False, fontsize=9)
ax.grid(axis='y', linestyle='--', alpha=0.4)

def autolabel(rects):
    for rect in rects:
        h = rect.get_height()
        ax.annotate(f'{h:.1f}',
                    xy=(rect.get_x() + rect.get_width() / 2, h),
                    xytext=(0, 3), textcoords="offset points",
                    ha='center', va='bottom', fontsize=7, rotation=90)

for rects in all_rects:
    autolabel(rects)

plt.tight_layout()
fig4_path = os.path.join(OUTPUT_DIR, "gambar4_komparasi_metrik.png")
plt.savefig(fig4_path, dpi=300, bbox_inches='tight')
plt.close()
print(f"   -> Tersimpan di: {fig4_path}")

# ==============================================================================
# 5. EKSPOR TABEL EVALUASI CSV
# ==============================================================================
tabel_csv_path = os.path.join(OUTPUT_DIR, "tabel_evaluasi_model.csv")
df_focus.to_csv(tabel_csv_path, index=False)
print(f"[5/5] Tabel evaluasi tersimpan di: {tabel_csv_path}")

print("Seluruh figur publikasi berhasil digenerate di figures_jurnal/!")
