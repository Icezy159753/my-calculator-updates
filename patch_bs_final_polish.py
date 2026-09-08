from pathlib import Path


path = Path(__file__).resolve().parent / "tmp_bs_xlsm_build" / "build_xlsm.py"
text = path.read_text(encoding="utf-8")
replacements = {
    '        control.Columns("A").ColumnWidth = 7':
        '        control.Columns("A").ColumnWidth = 12',
    '''        raw.Columns.ColumnWidth = 12

        engine.Cells(1, 1).Value = "Asset"''':
        '''        raw.Columns.ColumnWidth = 12
        raw.Activate()
        excel.ActiveWindow.SplitRow = 1
        excel.ActiveWindow.FreezePanes = True

        engine.Cells(1, 1).Value = "Asset"''',
}
for old, new in replacements.items():
    assert text.count(old) == 1, old
    text = text.replace(old, new, 1)
path.write_text(text, encoding="utf-8")
print("polished Control labels and Rawdata freeze panes")
