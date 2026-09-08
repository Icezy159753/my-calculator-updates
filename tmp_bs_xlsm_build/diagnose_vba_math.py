import importlib.util
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


def load_module():
    spec = importlib.util.spec_from_file_location("bs_diag", PROGRAM)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def standardized_betas(frame):
    cols = ["N_S", "N_P", "N_C", "N_E", "ZA"]
    data = frame[cols].dropna().astype(float)
    if len(data) < 6:
        return None
    x = data.iloc[:, :4]
    y = data.iloc[:, 4]
    xz = (x - x.mean()) / x.std(ddof=1)
    yz = (y - y.mean()) / y.std(ddof=1)
    design = np.column_stack([np.ones(len(xz)), xz.to_numpy()])
    beta = np.linalg.lstsq(design, yz.to_numpy(), rcond=None)[0][1:]
    ratios = np.abs(beta) / np.abs(beta).sum() * 100
    return ratios


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

groups = window._build_analysis_groups("Index1", "S1_Gen")
reference = pd.read_excel(REFERENCE, sheet_name="Summary")
expected = reference.set_index("Filter")[["B.S", "B.P", "B.C", "B.E"]]
rows = []
for name, group in groups.items():
    if name not in expected.index:
        continue
    direct = standardized_betas(group)
    if direct is None:
        continue
    target = expected.loc[name].to_numpy(float)
    rows.append({
        "Filter": name,
        "N": int(group[["N_S", "N_P", "N_C", "N_E", "ZA"]].dropna().shape[0]),
        "MaxAbsDiff": float(np.max(np.abs(direct - target))),
        "Direct": np.round(direct, 4).tolist(),
        "Expected": np.round(target, 4).tolist(),
    })

for row in rows:
    print(row)
print("max_diff", max(row["MaxAbsDiff"] for row in rows))
window.close()
