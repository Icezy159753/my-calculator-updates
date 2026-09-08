from pathlib import Path


root = Path(__file__).resolve().parent
runner = root / "tmp_bs_xlsm_build" / "runner.py"
vba = root / "tmp_bs_xlsm_build" / "BrandSenseRunner.bas"

text = runner.read_text(encoding="utf-8")
needle = "import openpyxl\nimport pandas as pd\n\n\n"
insert = '''import openpyxl
import pandas as pd


SETTING_COLUMNS = (
    'C', 'A', 'S', 'P', 'E', 'AgreeS', 'AgreeP', 'Filter_Var')


def setting_required_columns(setting_path):
    frame = pd.read_excel(setting_path, sheet_name='Settings')
    required = set()
    for column in SETTING_COLUMNS:
        if column not in frame.columns:
            continue
        for value in frame[column].dropna():
            text = str(value).strip()
            if not text:
                continue
            for item in text.split(','):
                item = item.strip()
                if item:
                    required.add(item)
    return required


def select_setting(settings_dir, raw_columns):
    candidates = []
    for filename in sorted(os.listdir(settings_dir)):
        if not filename.lower().endswith('.xlsx'):
            continue
        path = os.path.join(settings_dir, filename)
        required = setting_required_columns(path)
        missing = sorted(required.difference(raw_columns))
        candidates.append((len(missing), -len(required), filename, path, missing))
    if not candidates:
        raise RuntimeError('No embedded Setting workbook was found')
    candidates.sort()
    missing_count, _, filename, path, missing = candidates[0]
    if missing_count:
        preview = ', '.join(missing[:20])
        raise RuntimeError(
            'Rawdata does not match Setting 1, 4, 5, or 6. '
            'Closest setting is ' + filename + '; missing: ' + preview)
    return path, filename


'''
assert text.count(needle) == 1
text = text.replace(needle, insert, 1)
replacements = {
    "setting_path = os.path.abspath(sys.argv[2])":
        "settings_dir = os.path.abspath(sys.argv[2])",
    "        patch_setting_path(setting_path, sav_path)\n        rawdata = pd.read_csv(raw_csv, low_memory=False)":
        "        rawdata = pd.read_csv(raw_csv, low_memory=False)",
    "        if rawdata.empty:\n            raise RuntimeError('Rawdata sheet contains no data rows')\n\n        module = load_module(program_path)":
        "        if rawdata.empty:\n            raise RuntimeError('Rawdata sheet contains no data rows')\n        setting_path, setting_name = select_setting(\n            settings_dir, set(rawdata.columns))\n        patch_setting_path(setting_path, sav_path)\n\n        module = load_module(program_path)",
    "            'mode': mode,\n            'raw_rows'":
        "            'mode': mode,\n            'setting': setting_name,\n            'raw_rows'",
}
for old, new in replacements.items():
    assert text.count(old) == 1, old
    text = text.replace(old, new, 1)
runner.write_text(text, encoding="utf-8")

text = vba.read_text(encoding="utf-8-sig")
old = '    ExtractAsset "setting.xlsx", workFolder & "\\setting.xlsx"'
new = '''    MkDir workFolder & "\\settings"
    ExtractAsset "setting_1.xlsx", workFolder & "\\settings\\setting_1.xlsx"
    ExtractAsset "setting_4.xlsx", workFolder & "\\settings\\setting_4.xlsx"
    ExtractAsset "setting_5.xlsx", workFolder & "\\settings\\setting_5.xlsx"
    ExtractAsset "setting_6.xlsx", workFolder & "\\settings\\setting_6.xlsx"'''
assert text.count(old) == 1
text = text.replace(old, new, 1)
old = '        QuoteText(workFolder & "\\setting.xlsx") & " " & _'
new = '        QuoteText(workFolder & "\\settings") & " " & _'
assert text.count(old) == 1
text = text.replace(old, new, 1)
vba.write_text(text, encoding="utf-8-sig")

print("patched runner.py and BrandSenseRunner.bas")
