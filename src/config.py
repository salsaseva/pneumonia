"""
Configuration and constants for Pneumonia GLCM Detection Revisions.
Based on Revision Guidelines: Ortiz-Toro et al. [15] and Hismaya & Sulistianingsih [19].
"""
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATASET_DIR = "/Users/dicky/Downloads/chest_xray"

RESULTS_DIR = os.path.join(BASE_DIR, "results")
AUDIT_DIR = os.path.join(RESULTS_DIR, "audit")
TABLES_DIR = os.path.join(RESULTS_DIR, "tables")
FIGURES_DIR = os.path.join(RESULTS_DIR, "figures")
MODELS_DIR = os.path.join(RESULTS_DIR, "models")
CACHE_DIR = os.path.join(BASE_DIR, "data", "cache")

RANDOM_SEED = 42
ADDITIONAL_SEEDS = [123, 456, 789, 2026]

# Image parameters
IMG_SIZE = (150, 150)
ANGLES = [0, 3.141592653589793/4, 3.141592653589793/2, 3*3.141592653589793/4]
DISTANCES = [1]
GLCM_PROPERTIES = ['contrast', 'dissimilarity', 'homogeneity', 'energy', 'correlation', 'ASM']

# Evaluation metrics labels
METRIC_COLS = [
    'Model', 'Acc (%)', 'PPV (%)', 'Sen (%)', 'Spe (%)', 'NPV (%)', 
    'F1 (%)', 'Bal Acc (%)', 'Macro-F1 (%)', 'ROC-AUC (%)'
]
