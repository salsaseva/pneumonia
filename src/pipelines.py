"""
Tahap 5: Pipeline & Konfigurasi Model
Menggunakan imblearn.pipeline.Pipeline untuk isolasi scaler, SMOTE, dan model.
Mendukung:
- Dua kondisi penyeimbangan: Tanpa SMOTE (None / class_weight) vs SMOTE
- Konfigurasi default ala Ortiz-Toro [15] vs Tuned vs Konfigurasi Lama
- Soft Voting dari model dasar yang sudah dituning
"""
import numpy as np
from sklearn.preprocessing import StandardScaler
from sklearn.svm import SVC
from sklearn.ensemble import RandomForestClassifier, VotingClassifier
from sklearn.neighbors import KNeighborsClassifier
from imblearn.pipeline import Pipeline as ImbPipeline
from imblearn.over_sampling import SMOTE
from src.config import RANDOM_SEED

def build_pipeline(model, scaler=True, smote=False, seed=RANDOM_SEED):
    """Membangun pipeline imblearn dengan isolasi ketat terhadap kebocoran data."""
    steps = []
    if scaler:
        steps.append(('scaler', StandardScaler()))
    if smote:
        steps.append(('smote', SMOTE(random_state=seed)))
    steps.append(('model', model))
    return ImbPipeline(steps)

def get_default_models(n_train_samples=4185, seed=RANDOM_SEED):
    """
    Model Default ala Ortiz-Toro et al. [15]:
    - KNN: K = akar N
    - SVM: Linear kernel, C = 1
    - RF: 100 pohon, min_samples_leaf = 2
    """
    k_val = max(3, int(round(np.sqrt(n_train_samples))))
    models = {
        'KNN_default': KNeighborsClassifier(n_neighbors=k_val),
        'SVM_default': SVC(kernel='linear', C=1.0, probability=True, random_state=seed),
        'RF_default': RandomForestClassifier(n_estimators=100, min_samples_leaf=2, random_state=seed)
    }
    return models

def get_legacy_models(seed=RANDOM_SEED):
    """
    Konfigurasi Lama (sebelum revisi):
    - KNN: k=5
    - SVM: RBF kernel, C=1.0
    - RF: 100 pohon, min_samples_leaf=1
    """
    models = {
        'KNN_legacy': KNeighborsClassifier(n_neighbors=5),
        'SVM_legacy': SVC(kernel='rbf', C=1.0, probability=True, random_state=seed),
        'RF_legacy': RandomForestClassifier(n_estimators=100, min_samples_leaf=1, random_state=seed)
    }
    return models

def get_tuning_hyperparameters(n_train_samples=4185):
    """Ruang pencarian hyperparameter untuk tuning pada data validasi."""
    k_sqrt = max(3, int(round(np.sqrt(n_train_samples))))
    k_candidates = sorted(list(set([3, 5, 7, 11, 25, 45, k_sqrt])))

    return {
        'KNN': [{'n_neighbors': k} for k in k_candidates],
        'SVM': [
            {'kernel': 'linear', 'C': c, 'probability': True}
            for c in [0.1, 1.0, 10.0, 100.0]
        ] + [
            {'kernel': 'rbf', 'C': c, 'gamma': g, 'probability': True}
            for c in [0.1, 1.0, 10.0, 100.0]
            for g in ['scale', 'auto', 0.01, 0.1]
        ],
        'RF': [
            {'n_estimators': n, 'max_depth': d, 'min_samples_leaf': l}
            for n in [100, 200, 300]
            for d in [None, 10, 20]
            for l in [1, 2, 4]
        ]
    }

def create_voting_ensemble(best_svm, best_rf, best_knn):
    """Membuat Ensemble Soft Voting dari ketiga model dasar yang telah dituning."""
    return VotingClassifier(
        estimators=[
            ('svm', best_svm),
            ('rf', best_rf),
            ('knn', best_knn)
        ],
        voting='soft'
    )
