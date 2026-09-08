from pathlib import Path


path = Path(__file__).resolve().parent / "tmp_bs_xlsm_build" / "BrandSenseRunner.bas"
text = path.read_text(encoding="utf-8-sig")
old = '    MsgBox Err.Description, vbCritical, "BrandSense Excel Runner"'
new = '''    If Application.Visible Then
        MsgBox Err.Description, vbCritical, "BrandSense Excel Runner"
    End If'''
assert text.count(old) == 1
path.write_text(text.replace(old, new, 1), encoding="utf-8-sig")
print("patched headless macro error handling")
