"""
Tahap 8: Uji Statistik & Analisis Keragaman Ensemble
1. Uji McNemar berpasangan (6 pasangan model) dengan koreksi Bonferroni (alpha = 0.05 / 6 ≈ 0.0083).
2. Bootstrap 95% Confidence Intervals untuk F1-Score dan Akurasi.
3. Analisis Keragaman Ensemble: Disagreement Rate & Korelasi Galat (Yule's Q-statistic & Pearson).
"""
import os
import itertools
import numpy as np
import pandas as pd
from scipy.stats import binom, chi2
from sklearn.metrics import accuracy_score, f1_score
from src.config import TABLES_DIR

def mcnemar_test_paired(y_true, y_pred1, y_pred2):
    """
    Uji McNemar berpasangan dengan koreksi kontinuitas Edwards
    atau exact binomial jika b + c < 25.
    """
    correct1 = (y_pred1 == y_true)
    correct2 = (y_pred2 == y_true)

    # b: Model 1 benar, Model 2 salah
    b = np.sum(correct1 & ~correct2)
    # c: Model 1 salah, Model 2 benar
    c = np.sum(~correct1 & correct2)

    total_discordant = b + c
    if total_discordant == 0:
        return {'b': int(b), 'c': int(c), 'stat': 0.0, 'p_value': 1.0, 'method': 'none'}

    if total_discordant < 25:
        # Exact binomial test (two-sided)
        p_val = 2 * min(binom.cdf(min(b, c), total_discordant, 0.5), 1.0 - binom.cdf(min(b, c) - 1, total_discordant, 0.5))
        p_val = min(1.0, p_val)
        stat = float(min(b, c))
        method = 'exact binomial'
    else:
        # Continuity corrected chi-square
        stat = (abs(b - c) - 1.0)**2 / (b + c)
        p_val = 1.0 - chi2.cdf(stat, df=1)
        method = 'chi2 continuity corrected'

    return {
        'b (A correct, B wrong)': int(b),
        'c (A wrong, B correct)': int(c),
        'statistic': round(stat, 4),
        'p_value': float(p_val),
        'method': method
    }

def run_all_mcnemar_tests(y_true, predictions_dict, alpha=0.05, output_csv=None):
    """
    Menjalankan Uji McNemar untuk semua pasangan model.
    alpha_corrected = alpha / k (Bonferroni)
    """
    model_names = list(predictions_dict.keys())
    pairs = list(itertools.combinations(model_names, 2))
    k = len(pairs)
    alpha_bonf = alpha / k

    results = []
    for m1, m2 in pairs:
        res = mcnemar_test_paired(y_true, predictions_dict[m1], predictions_dict[m2])
        is_sig = res['p_value'] < alpha_bonf
        conclusion = "Signifikan (H0 ditolak)" if is_sig else "Tidak Signifikan (Setara)"

        results.append({
            'Model A': m1,
            'Model B': m2,
            'b': res['b (A correct, B wrong)'],
            'c': res['c (A wrong, B correct)'],
            'Statistik': res['statistic'],
            'p-value': f"{res['p_value']:.5f}",
            'Alpha Bonferroni': round(alpha_bonf, 5),
            'Signifikan?': "Ya" if is_sig else "Tidak",
            'Kesimpulan': conclusion
        })

    df_res = pd.DataFrame(results)
    if output_csv is None:
        os.makedirs(TABLES_DIR, exist_ok=True)
        output_csv = os.path.join(TABLES_DIR, "mcnemar_pairwise_tests.csv")
    df_res.to_csv(output_csv, index=False)
    print(f"[Statistics] Hasil uji McNemar disimpan ke {output_csv}")
    return df_res

def compute_bootstrap_ci(y_true, y_pred, n_bootstraps=1000, ci=0.95, seed=42):
    """Menghitung 95% Confidence Interval untuk Akurasi dan F1 via Bootstrap."""
    rng = np.random.RandomState(seed)
    n = len(y_true)
    accs, f1s = [], []

    for _ in range(n_bootstraps):
        idx = rng.randint(0, n, size=n)
        yt_boot = y_true[idx]
        yp_boot = y_pred[idx]
        accs.append(accuracy_score(yt_boot, yp_boot))
        f1s.append(f1_score(yt_boot, yp_boot, zero_division=0))

    lower_pct = ((1.0 - ci) / 2.0) * 100
    upper_pct = (1.0 - (1.0 - ci) / 2.0) * 100

    return {
        'Acc Mean (%)': round(np.mean(accs) * 100, 2),
        'Acc CI Lower (%)': round(np.percentile(accs, lower_pct) * 100, 2),
        'Acc CI Upper (%)': round(np.percentile(accs, upper_pct) * 100, 2),
        'F1 Mean (%)': round(np.mean(f1s) * 100, 2),
        'F1 CI Lower (%)': round(np.percentile(f1s, lower_pct) * 100, 2),
        'F1 CI Upper (%)': round(np.percentile(f1s, upper_pct) * 100, 2),
    }

def analyze_ensemble_diversity(y_true, preds_dict, output_csv=None):
    """
    Analisis keragaman ensemble antar model dasar:
    1. Disagreement rate (% sampel di mana kedua model berbeda prediksi).
    2. Yule's Q-statistic: Q = (N11*N00 - N10*N01) / (N11*N00 + N10*N01).
    3. Pearson correlation of binary error vectors.
    """
    models = list(preds_dict.keys())
    rows = []
    for m1, m2 in itertools.combinations(models, 2):
        p1 = preds_dict[m1]
        p2 = preds_dict[m2]

        disagreement = np.mean(p1 != p2) * 100.0

        # Error vectors: 1 jika salah, 0 jika benar
        e1 = (p1 != y_true).astype(int)
        e2 = (p2 != y_true).astype(int)

        n11 = np.sum((e1 == 1) & (e2 == 1))
        n00 = np.sum((e1 == 0) & (e2 == 0))
        n10 = np.sum((e1 == 1) & (e2 == 0))
        n01 = np.sum((e1 == 0) & (e2 == 1))

        denom = (n11 * n00 + n10 * n01)
        q_stat = (n11 * n00 - n10 * n01) / denom if denom > 0 else 1.0

        # Pearson error correlation
        corr = np.corrcoef(e1, e2)[0, 1] if np.std(e1) > 0 and np.std(e2) > 0 else 0.0

        rows.append({
            'Model Pair': f"{m1} vs {m2}",
            'Disagreement Rate (%)': round(disagreement, 2),
            'Yule Q-Statistic': round(q_stat, 4),
            'Error Correlation (r)': round(corr, 4),
            'N11 (Both Error)': int(n11),
            'N00 (Both Correct)': int(n00),
            'N10 (M1 Error, M2 Correct)': int(n10),
            'N01 (M1 Correct, M2 Error)': int(n01)
        })

    df_div = pd.DataFrame(rows)
    if output_csv is None:
        os.makedirs(TABLES_DIR, exist_ok=True)
        output_csv = os.path.join(TABLES_DIR, "ensemble_diversity_analysis.csv")
    df_div.to_csv(output_csv, index=False)
    print(f"[Statistics] Analisis keragaman ensemble disimpan ke {output_csv}")
    return df_div
