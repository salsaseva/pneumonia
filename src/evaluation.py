"""
Tahap 6 & 7: Evaluasi & Seleksi Model
Menghitung seluruh metrik standar:
Acc, PPV, Sen, Spe, NPV, F1, Balanced Acc, Macro-F1, ROC-AUC.
Plotting Confusion Matrix, ROC Curve, DET Curve.
"""
import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.metrics import (
    confusion_matrix, accuracy_score, precision_score, recall_score,
    f1_score, roc_auc_score, roc_curve, det_curve
)
from src.config import FIGURES_DIR, TABLES_DIR

def compute_all_metrics(y_true, y_pred, y_prob=None, model_name="Model"):
    """
    Menghitung 10 metrik diagnostik lengkap:
    Acc, PPV, Sen, Spe, NPV, F1, Bal Acc, Macro-F1, ROC-AUC, plus TN/FP/FN/TP.
    """
    cm = confusion_matrix(y_true, y_pred)
    tn, fp, fn, tp = cm.ravel()

    acc = accuracy_score(y_true, y_pred)
    ppv = precision_score(y_true, y_pred, zero_division=0)
    sen = recall_score(y_true, y_pred, zero_division=0)
    spe = tn / (tn + fp) if (tn + fp) > 0 else 0.0
    npv = tn / (tn + fn) if (tn + fn) > 0 else 0.0
    f1 = f1_score(y_true, y_pred, zero_division=0)
    bal_acc = (sen + spe) / 2.0
    macro_f1 = f1_score(y_true, y_pred, average='macro', zero_division=0)

    auc_val = 0.0
    if y_prob is not None:
        try:
            if y_prob.ndim == 2:
                prob_pos = y_prob[:, 1]
            else:
                prob_pos = y_prob
            auc_val = roc_auc_score(y_true, prob_pos)
        except Exception:
            auc_val = 0.0

    return {
        'Model': model_name,
        'Accuracy (%)': round(acc * 100, 2),
        'Precision/PPV (%)': round(ppv * 100, 2),
        'Recall/Sen (%)': round(sen * 100, 2),
        'Specificity (%)': round(spe * 100, 2),
        'NPV (%)': round(npv * 100, 2),
        'F1-Score (%)': round(f1 * 100, 2),
        'Balanced Accuracy (%)': round(bal_acc * 100, 2),
        'Macro-F1 (%)': round(macro_f1 * 100, 2),
        'ROC-AUC (%)': round(auc_val * 100, 2),
        'TN': int(tn),
        'FP': int(fp),
        'FN': int(fn),
        'TP': int(tp)
    }

def plot_confusion_matrices(cms, class_names=['Normal', 'Pneumonia'], output_path=None):
    """Plotting perbandingan confusion matrix untuk seluruh model."""
    os.makedirs(FIGURES_DIR, exist_ok=True)
    n_models = len(cms)
    fig, axes = plt.subplots(1, n_models, figsize=(4.5 * n_models, 4))
    if n_models == 1:
        axes = [axes]

    for ax, (name, cm) in zip(axes, cms.items()):
        im = ax.imshow(cm, interpolation='nearest', cmap=plt.cm.Blues)
        ax.set_title(name, fontsize=12, fontweight='bold', pad=10)
        tick_marks = np.arange(len(class_names))
        ax.set_xticks(tick_marks)
        ax.set_xticklabels(class_names, fontsize=10)
        ax.set_yticks(tick_marks)
        ax.set_yticklabels(class_names, fontsize=10)

        thresh = cm.max() / 2.
        for i in range(cm.shape[0]):
            for j in range(cm.shape[1]):
                ax.text(j, i, f"{cm[i, j]}",
                        ha="center", va="center",
                        color="white" if cm[i, j] > thresh else "black",
                        fontsize=12, fontweight='bold')
        ax.set_ylabel('True Label', fontsize=10)
        ax.set_xlabel('Predicted Label', fontsize=10)

    plt.tight_layout()
    if output_path is None:
        output_path = os.path.join(FIGURES_DIR, "confusion_matrices.png")
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"[Evaluation] Confusion matrix disimpan ke {output_path}")

def plot_roc_and_det_curves(model_probs, y_true, output_roc=None, output_det=None):
    """Plotting kurva ROC dan DET (Detection Error Tradeoff)."""
    os.makedirs(FIGURES_DIR, exist_ok=True)

    # 1. ROC Curve
    plt.figure(figsize=(7, 6))
    for name, prob in model_probs.items():
        fpr, tpr, _ = roc_curve(y_true, prob)
        auc = roc_auc_score(y_true, prob)
        plt.plot(fpr, tpr, lw=2, label=f"{name} (AUC = {auc*100:.2f}%)")
    plt.plot([0, 1], [0, 1], color='navy', lw=1.5, linestyle='--', label='Random Guess')
    plt.xlim([0.0, 1.0])
    plt.ylim([0.0, 1.05])
    plt.xlabel('False Positive Rate (1 - Specificity)', fontsize=11)
    plt.ylabel('True Positive Rate (Sensitivity)', fontsize=11)
    plt.title('Receiver Operating Characteristic (ROC) Curves', fontsize=13, fontweight='bold', pad=12)
    plt.legend(loc="lower right", frameon=True, fontsize=10)
    plt.grid(alpha=0.3)
    if output_roc is None:
        output_roc = os.path.join(FIGURES_DIR, "roc_curves.png")
    plt.savefig(output_roc, dpi=300, bbox_inches='tight')
    plt.close()

    # 2. DET Curve
    plt.figure(figsize=(7, 6))
    for name, prob in model_probs.items():
        fpr, fnr, _ = det_curve(y_true, prob)
        plt.plot(fpr, fnr, lw=2, label=f"{name}")
    plt.xlabel('False Positive Rate', fontsize=11)
    plt.ylabel('False Negative Rate', fontsize=11)
    plt.title('Detection Error Tradeoff (DET) Curves', fontsize=13, fontweight='bold', pad=12)
    plt.legend(loc="upper right", frameon=True, fontsize=10)
    plt.grid(alpha=0.3)
    if output_det is None:
        output_det = os.path.join(FIGURES_DIR, "det_curves.png")
    plt.savefig(output_det, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"[Evaluation] Kurva ROC & DET disimpan ke {output_roc} dan {output_det}")
