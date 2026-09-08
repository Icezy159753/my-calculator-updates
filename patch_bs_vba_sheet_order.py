from pathlib import Path


path = Path(__file__).resolve().parent / "tmp_bs_xlsm_build" / "BrandSenseRunner.bas"
text = path.read_text(encoding="ascii")
old = '''    sourceBook.Close SaveChanges:=False
    ThisWorkbook.Worksheets(CONTROL_SHEET).Activate
End Sub'''
new = '''    sourceBook.Close SaveChanges:=False
    ThisWorkbook.Worksheets(CONTROL_SHEET).Move _
        Before:=ThisWorkbook.Worksheets(1)
    ThisWorkbook.Worksheets(RAW_SHEET).Move _
        After:=ThisWorkbook.Worksheets(CONTROL_SHEET)
    ThisWorkbook.Worksheets(CONTROL_SHEET).Activate
End Sub'''
assert text.count(old) == 1
path.write_text(text.replace(old, new, 1), encoding="ascii")
print("patched result sheet order")
