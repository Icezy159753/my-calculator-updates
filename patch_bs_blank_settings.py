from pathlib import Path


path = Path(__file__).resolve().parent / "tmp_bs_xlsm_build" / "build_xlsm.py"
text = path.read_text(encoding="utf-8")
replacements = {
    "    rawdata = rawdata.where(pd.notna(rawdata), None)":
        "    rawdata = rawdata.astype(object).where(pd.notna(rawdata), None)",
    "    setting_data = setting_data.where(pd.notna(setting_data), None)":
        "    setting_data = setting_data.astype(object).where(pd.notna(setting_data), None)",
    "    label_data = label_data.where(pd.notna(label_data), None)":
        "    label_data = label_data.astype(object).where(pd.notna(label_data), None)",
}
for old, new in replacements.items():
    assert text.count(old) == 1, old
    text = text.replace(old, new, 1)
path.write_text(text, encoding="utf-8")
print("patched blank cells to use real None values")
