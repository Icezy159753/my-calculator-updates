import importlib.util
import itertools
import os
from pathlib import Path

import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
PROGRAM = ROOT / "All_Programs" / "123_Program_Run_Brandsence2026.py"
SETTING = ROOT / "data" / "BS" / "1_Setting BS Set1 Bangkok Hospital Headquarters.xlsx"
SAV = ROOT / "data" / "BS" / "SPSS_preserved_utf8_Final.sav"
RAW = Path(__file__).resolve().parent / "test_runtime" / "rawdata.csv"
REFERENCE = Path(__file__).resolve().parent / "test_runtime" / "BrandSense_Engine_Input BS Output.xlsx"


def equamax_objective(loadings, kappa=0.5):
    p, k = loadings.shape
    n = np.ones((k, k)) - np.eye(k)
    m = np.ones((p, p)) - np.eye(p)
    squared = loadings ** 2
    f1 = (1 - kappa) * np.trace(squared.T @ squared @ n) / 4
    f2 = kappa * np.trace(squared.T @ m @ squared) / 4
    gradient = (
        (1 - kappa) * loadings * (squared @ n)
        + kappa * loadings * (m @ squared)
    )
    return gradient, f1 + f2


def equamax(loadings, kappa=0.5, max_iter=250, tol=1e-5):
    arr = loadings.copy()
    rotation = np.eye(arr.shape[1])
    alpha = 1.0
    rotated = arr @ rotation
    obj_grad, criterion = equamax_objective(rotated, kappa)
    gradient = arr.T @ obj_grad
    for _ in range(max_iter + 1):
        matrix = rotation.T @ gradient
        symmetric = (matrix + matrix.T) / 2
        projected = gradient - rotation @ symmetric
        s = np.sqrt(np.trace(projected.T @ projected))
        if s < tol:
            break
        alpha *= 2
        for _ in range(11):
            candidate = rotation - alpha * projected
            u, _, vt = np.linalg.svd(candidate)
            new_rotation = u @ vt
            rotated = arr @ new_rotation
            new_grad, new_criterion = equamax_objective(rotated, kappa)
            if new_criterion < criterion - 0.5 * s * s * alpha:
                break
            alpha /= 2
        rotation = new_rotation
        criterion = new_criterion
        gradient = arr.T @ new_grad
    return rotated


def principal_loadings(data):
    values = data.to_numpy(float)
    standardized = (values - values.mean(axis=0)) / values.std(axis=0, ddof=0)
    corr = np.corrcoef(standardized, rowvar=False)
    eigenvalues, eigenvectors = np.linalg.eigh(corr)
    order = np.argsort(eigenvalues)[::-1]
    eigenvalues = eigenvalues[order]
    eigenvectors = eigenvectors[:, order]
    return eigenvectors * np.sqrt(np.maximum(eigenvalues, 0))


def one_to_one_mapping(abs_loadings):
    primary = np.argmax(abs_loadings, axis=1)
    collision = len(set(primary.tolist())) < 4
    if not collision:
        return {factor: variable for variable, factor in enumerate(primary)}
    best = None
    for permutation in itertools.permutations(range(4)):
        score = sum(abs_loadings[variable, factor] for variable, factor in enumerate(permutation))
        if best is None or score > best[0]:
            best = (score, permutation)
    return {factor: variable for variable, factor in enumerate(best[1])}


def pure_ratios(frame):
    variables = ["N_S", "N_P", "N_C", "N_E"]
    factor_data = frame[variables].dropna().copy()
    loadings = equamax(principal_loadings(factor_data))
    ss = np.sum(loadings ** 2, axis=0)
    loadings = loadings[:, np.argsort(ss)[::-1]]
    mapping = one_to_one_mapping(np.abs(loadings))

    values = factor_data.to_numpy(float)
    standardized = (values - values.mean(axis=0)) / values.std(axis=0, ddof=0)
    correlation = factor_data.corr().to_numpy()
    inv_corr = np.linalg.inv(correlation)
    temp = loadings.T @ inv_corr @ loadings
    eigvals, eigvecs = np.linalg.eigh(temp)
    inv_sqrt = eigvecs @ np.diag(np.where(eigvals > 1e-12, 1 / np.sqrt(eigvals), 0)) @ eigvecs.T
    coefficients = inv_corr @ loadings @ inv_sqrt
    scores = standardized @ coefficients

    score_frame = pd.DataFrame(scores, index=factor_data.index)
    joined = frame[["ZA"]].join(score_frame).dropna()
    x = joined.iloc[:, 1:].to_numpy(float)
    y = joined.iloc[:, 0].to_numpy(float)
    x_centered = x - x.mean(axis=0)
    y_centered = y - y.mean()
    coefficients = np.linalg.lstsq(x_centered, y_centered, rcond=None)[0]
    beta = coefficients * (x.std(axis=0, ddof=1) / y.std(ddof=1))
    by_variable = np.zeros(4)
    for factor, variable in mapping.items():
        by_variable[variable] = beta[factor]
    return np.abs(by_variable) / np.abs(by_variable).sum() * 100, loadings


def load_module():
    spec = importlib.util.spec_from_file_location("bs_pure_diag", PROGRAM)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


os.environ["QT_QPA_PLATFORM"] = "offscreen"
module = load_module()
from PyQt6.QtWidgets import QApplication

app = QApplication.instance() or QApplication([])
window = module.SpssProcessorApp()
window._load_settings_file(str(SETTING), require_pathfile=False)
window.load_spss_file(filepath=str(SAV), raise_on_error=True)
window._snapshot_ui_inputs()
window.df = pd.read_csv(RAW, low_memory=False)
window._snapshot_ui_inputs()
window._transform_pipeline(with_compute_c=True)
candidates = window._build_long_qc_candidates()
window.qc_candidates_df = candidates
window._apply_qc_candidate_rows(candidates.index.tolist())
window._qc_use_model_safeguards = True

reference = pd.read_excel(REFERENCE, sheet_name="Summary").set_index("Filter")
max_diff = 0
for name, group in window._build_analysis_groups("Index1", "S1_Gen").items():
    if name not in reference.index:
        continue
    try:
        ratios, loadings = pure_ratios(group)
    except Exception as exc:
        print(name, "ERROR", exc)
        continue
    expected = reference.loc[name, ["B.S", "B.P", "B.C", "B.E"]].to_numpy(float)
    diff = float(np.max(np.abs(ratios - expected)))
    max_diff = max(max_diff, diff)
    print(name, "diff", round(diff, 10), "pure", np.round(ratios, 4).tolist(), "expected", np.round(expected, 4).tolist())
print("MAX_DIFF", max_diff)
window.close()
