from pathlib import Path

p = Path("tmp_bs_xlsm_build/build_xlsm_pure.py")
s = p.read_text(encoding="utf-8")
old = '''        labels = pd.read_excel(source, sheet_name="Label")
        labels.insert(0, "Set_ID", set_id)
        labels["Filter_Code"] = filter_codes[: len(labels)]
        labels["Filter_Label"] = filter_labels[: len(labels)]'''
new = '''        labels = pd.read_excel(source, sheet_name="Label")
        labels = labels.reindex(range(max(len(labels), len(filter_codes))))
        labels.insert(0, "Set_ID", set_id)
        labels["Filter_Code"] = filter_codes + [None] * (len(labels) - len(filter_codes))
        labels["Filter_Label"] = filter_labels + [None] * (len(labels) - len(filter_labels))'''
if old not in s:
    raise SystemExit("filter label anchor not found")
s = s.replace(old, new)
p.write_text(s, encoding="utf-8")
