from pathlib import Path


root = Path(__file__).resolve().parent
runner_path = root / "tmp_bs_xlsm_build" / "runner.py"
vba_path = root / "tmp_bs_xlsm_build" / "BrandSenseRunner.bas"
build_path = root / "tmp_bs_xlsm_build" / "build_xlsm.py"


# --- runner.py: accept editable Setting tables from the current xlsm ---
text = runner_path.read_text(encoding="utf-8")
old = '''def setting_required_columns(setting_path):
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
new = '''def required_columns_from_frame(frame):
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


def setting_required_columns(setting_path):
    frame = pd.read_excel(setting_path, sheet_name='Settings')
    return required_columns_from_frame(frame)


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


def append_frame(worksheet, frame):
    worksheet.append(list(frame.columns))
    for row in frame.itertuples(index=False, name=None):
        worksheet.append([
            None if pd.isna(value) else value
            for value in row
        ])


def select_workbook_setting(workbook_path, raw_columns, workdir):
    settings = pd.read_excel(
        workbook_path, sheet_name='Setting Variables')
    labels = pd.read_excel(
        workbook_path, sheet_name='Setting Labels')
    if 'Set_ID' not in settings.columns:
        raise RuntimeError('Setting Variables has no Set_ID column')
    if 'Set_ID' not in labels.columns:
        raise RuntimeError('Setting Labels has no Set_ID column')

    candidates = []
    for set_id, group in settings.groupby('Set_ID', sort=True):
        active = group.drop(columns=['Set_ID']).reset_index(drop=True)
        required = required_columns_from_frame(active)
        missing = sorted(required.difference(raw_columns))
        display_id = str(int(set_id)) if float(set_id).is_integer() else str(set_id)
        candidates.append((
            len(missing), -len(required), display_id, active, missing))
    if not candidates:
        raise RuntimeError('Setting Variables contains no Setting rows')
    candidates.sort(key=lambda item: (item[0], item[1], item[2]))
    missing_count, _, display_id, active, missing = candidates[0]
    if missing_count:
        preview = ', '.join(missing[:20])
        raise RuntimeError(
            'Rawdata does not match Setting 1, 4, 5, or 6. '
            'Closest workbook Set is ' + display_id + '; missing: ' + preview)

    label_group = labels[
        labels['Set_ID'].astype(str).str.replace('.0', '', regex=False)
        == display_id
    ].drop(columns=['Set_ID']).reset_index(drop=True)
    active_path = os.path.join(workdir, 'active_setting.xlsx')
    workbook = openpyxl.Workbook()
    settings_sheet = workbook.active
    settings_sheet.title = 'Settings'
    append_frame(settings_sheet, active)
    labels_sheet = workbook.create_sheet('Label')
    append_frame(labels_sheet, label_group)
    workbook.save(active_path)
    return active_path, 'Set ' + display_id + ' (workbook)'
'''
assert text.count(old) == 1
text = text.replace(old, new, 1)
text = text.replace(
    "        settings_dir = os.path.abspath(sys.argv[2])",
    "        settings_source = os.path.abspath(sys.argv[2])",
    1,
)
old = '''        setting_path, setting_name = select_setting(
            settings_dir, set(rawdata.columns))
        patch_setting_path(setting_path, sav_path)'''
new = '''        if os.path.isdir(settings_source):
            setting_path, setting_name = select_setting(
                settings_source, set(rawdata.columns))
        else:
            setting_path, setting_name = select_workbook_setting(
                settings_source, set(rawdata.columns), workdir)
        patch_setting_path(setting_path, sav_path)'''
assert text.count(old) == 1
text = text.replace(old, new, 1)
runner_path.write_text(text, encoding="utf-8")


# --- VBA: save and pass this workbook as the Setting source ---
text = vba_path.read_text(encoding="ascii")
old = '''    If Not RawdataIsReady() Then
        Err.Raise vbObjectError + 101, , _
            "Rawdata must have a header row and at least one data row."
    End If

    workFolder = Environ$("TEMP")'''
new = '''    If Not RawdataIsReady() Then
        Err.Raise vbObjectError + 101, , _
            "Rawdata must have a header row and at least one data row."
    End If
    ThisWorkbook.Save

    workFolder = Environ$("TEMP")'''
assert text.count(old) == 1
text = text.replace(old, new, 1)
old = '''    ExtractAsset "runner.py", workFolder & "\\runner.py"
    MkDir workFolder & "\\settings"
    ExtractAsset "setting_1.xlsx", workFolder & "\\settings\\setting_1.xlsx"
    ExtractAsset "setting_4.xlsx", workFolder & "\\settings\\setting_4.xlsx"
    ExtractAsset "setting_5.xlsx", workFolder & "\\settings\\setting_5.xlsx"
    ExtractAsset "setting_6.xlsx", workFolder & "\\settings\\setting_6.xlsx"
    ExtractAsset "metadata.sav", workFolder & "\\metadata.sav"'''
new = '''    ExtractAsset "runner.py", workFolder & "\\runner.py"
    ExtractAsset "metadata.sav", workFolder & "\\metadata.sav"'''
assert text.count(old) == 1
text = text.replace(old, new, 1)
old = '        QuoteText(workFolder & "\\settings") & " " & _'
new = '        QuoteText(ThisWorkbook.FullName) & " " & _'
assert text.count(old) == 1
text = text.replace(old, new, 1)
old = '''    ThisWorkbook.Worksheets(RAW_SHEET).Move _
        After:=ThisWorkbook.Worksheets(CONTROL_SHEET)
    ThisWorkbook.Worksheets(CONTROL_SHEET).Activate'''
new = '''    ThisWorkbook.Worksheets(RAW_SHEET).Move _
        After:=ThisWorkbook.Worksheets(CONTROL_SHEET)
    ThisWorkbook.Worksheets("Setting Variables").Move _
        After:=ThisWorkbook.Worksheets(RAW_SHEET)
    ThisWorkbook.Worksheets("Setting Labels").Move _
        After:=ThisWorkbook.Worksheets("Setting Variables")
    ThisWorkbook.Worksheets(CONTROL_SHEET).Activate'''
assert text.count(old) == 1
text = text.replace(old, new, 1)
vba_path.write_text(text, encoding="ascii")


# --- build_xlsm.py: populate visible editable Setting tables ---
text = build_path.read_text(encoding="utf-8")
old = '''ASSETS = {
    "program.py": ROOT / "All_Programs" / "123_Program_Run_Brandsence2026.py",
    "runner.py": BUILD / "runner.py",
    "metadata.sav": ROOT / "data" / "BS" / "SPSS_preserved_utf8_Final.sav",
    "setting_1.xlsx": ROOT / "data" / "BS" / "1_Setting BS Set1 Bangkok Hospital Headquarters.xlsx",
    "setting_4.xlsx": ROOT / "data" / "BS" / "4_Setting BS Set4 Bangkok Hospital Pattaya.xlsx",
    "setting_5.xlsx": ROOT / "data" / "BS" / "5_Setting BS Set5 Bangkok Hospital Chanthaburi.xlsx",
    "setting_6.xlsx": ROOT / "data" / "BS" / "6_Setting BS Set6 Bangkok Hospital Rayong.xlsx",
}'''
new = '''SETTING_SOURCES = {
    1: ROOT / "data" / "BS" / "1_Setting BS Set1 Bangkok Hospital Headquarters.xlsx",
    4: ROOT / "data" / "BS" / "4_Setting BS Set4 Bangkok Hospital Pattaya.xlsx",
    5: ROOT / "data" / "BS" / "5_Setting BS Set5 Bangkok Hospital Chanthaburi.xlsx",
    6: ROOT / "data" / "BS" / "6_Setting BS Set6 Bangkok Hospital Rayong.xlsx",
}
ASSETS = {
    "program.py": ROOT / "All_Programs" / "123_Program_Run_Brandsence2026.py",
    "runner.py": BUILD / "runner.py",
    "metadata.sav": ROOT / "data" / "BS" / "SPSS_preserved_utf8_Final.sav",
}'''
assert text.count(old) == 1
text = text.replace(old, new, 1)
old = '    for path in [*ASSETS.values(), RAW_CSV, MODULE]:'
new = '    for path in [*ASSETS.values(), *SETTING_SOURCES.values(), RAW_CSV, MODULE]:'
assert text.count(old) == 1
text = text.replace(old, new, 1)
old = '''    matrix = [tuple(rawdata.columns)] + [tuple(row) for row in rawdata.itertuples(index=False, name=None)]

    red = rgb(198, 40, 45)'''
new = '''    matrix = [tuple(rawdata.columns)] + [tuple(row) for row in rawdata.itertuples(index=False, name=None)]

    setting_frames = []
    label_frames = []
    for set_id, source in SETTING_SOURCES.items():
        settings = pd.read_excel(source, sheet_name="Settings")
        settings.insert(0, "Set_ID", set_id)
        setting_frames.append(settings)
        labels = pd.read_excel(source, sheet_name="Label")
        labels.insert(0, "Set_ID", set_id)
        label_frames.append(labels)
    setting_data = pd.concat(setting_frames, ignore_index=True)
    label_data = pd.concat(label_frames, ignore_index=True)
    setting_data = setting_data.where(pd.notna(setting_data), None)
    label_data = label_data.where(pd.notna(label_data), None)
    setting_matrix = [tuple(setting_data.columns)] + [
        tuple(row) for row in setting_data.itertuples(index=False, name=None)]
    label_matrix = [tuple(label_data.columns)] + [
        tuple(row) for row in label_data.itertuples(index=False, name=None)]

    red = rgb(198, 40, 45)'''
assert text.count(old) == 1
text = text.replace(old, new, 1)
old = '''        raw = workbook.Worksheets.Add(After=control)
        raw.Name = "Rawdata"
        engine = workbook.Worksheets.Add(After=raw)
        engine.Name = "__Engine"'''
new = '''        raw = workbook.Worksheets.Add(After=control)
        raw.Name = "Rawdata"
        setting_variables = workbook.Worksheets.Add(After=raw)
        setting_variables.Name = "Setting Variables"
        setting_labels = workbook.Worksheets.Add(After=setting_variables)
        setting_labels.Name = "Setting Labels"
        engine = workbook.Worksheets.Add(After=setting_labels)
        engine.Name = "__Engine"'''
assert text.count(old) == 1
text = text.replace(old, new, 1)
old = '''        raw.Activate()
        excel.ActiveWindow.SplitRow = 1
        excel.ActiveWindow.FreezePanes = True

        engine.Cells(1, 1).Value = "Asset"'''
new = '''        raw.Activate()
        excel.ActiveWindow.SplitRow = 1
        excel.ActiveWindow.FreezePanes = True

        for setting_sheet, setting_values in (
                (setting_variables, setting_matrix),
                (setting_labels, label_matrix)):
            setting_rows = len(setting_values)
            setting_cols = len(setting_values[0])
            setting_target = setting_sheet.Range(
                setting_sheet.Cells(1, 1),
                setting_sheet.Cells(setting_rows, setting_cols))
            setting_target.Value = tuple(setting_values)
            setting_sheet.Cells.Font.Name = "Calibri"
            setting_sheet.Cells.Font.Size = 10
            setting_header = setting_sheet.Range(
                setting_sheet.Cells(1, 1),
                setting_sheet.Cells(1, setting_cols))
            setting_header.Interior.Color = blue
            setting_header.Font.Color = white
            setting_header.Font.Bold = True
            setting_header.WrapText = True
            setting_header.HorizontalAlignment = -4108
            setting_header.AutoFilter()
            setting_sheet.Rows(1).RowHeight = 34
            if setting_rows > 1:
                setting_body = setting_sheet.Range(
                    setting_sheet.Cells(2, 1),
                    setting_sheet.Cells(setting_rows, setting_cols))
                setting_body.Interior.Color = light_yellow
            setting_sheet.Columns.AutoFit()
            for column_index in range(1, setting_cols + 1):
                column = setting_sheet.Columns(column_index)
                column.ColumnWidth = min(max(column.ColumnWidth, 9), 42)
            setting_sheet.Activate()
            excel.ActiveWindow.SplitRow = 1
            excel.ActiveWindow.FreezePanes = True
            excel.ActiveWindow.DisplayGridlines = False

        setting_variables.Tab.Color = rgb(255, 192, 0)
        setting_labels.Tab.Color = rgb(255, 217, 102)

        setting_note = merge_value(
            control,
            "B20:G20",
            "Setting ที่ใช้รันอยู่ในชีท Setting Variables / Setting Labels และแก้ไขได้โดยตรง",
        )
        setting_note.Interior.Color = light_blue
        setting_note.Font.Bold = True
        setting_note.Font.Color = dark_red
        setting_note.HorizontalAlignment = -4108

        engine.Cells(1, 1).Value = "Asset"'''
assert text.count(old) == 1
text = text.replace(old, new, 1)
build_path.write_text(text, encoding="utf-8")

print("patched runner, VBA, and workbook builder for internal editable Settings")
