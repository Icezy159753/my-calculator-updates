import base64
import os
from pathlib import Path

import pandas as pd
import pythoncom
import win32com.client


ROOT = Path(__file__).resolve().parents[1]
BUILD = Path(__file__).resolve().parent
OUTPUT_DIR = ROOT / "outputs" / "brandsense_rawdata_xlsm"
OUTPUT = OUTPUT_DIR / "BrandSense_Rawdata_Auto_Set_1_4_5_6.xlsm"

SETTING_SOURCES = {
    1: ROOT / "data" / "BS" / "1_Setting BS Set1 Bangkok Hospital Headquarters.xlsx",
    4: ROOT / "data" / "BS" / "4_Setting BS Set4 Bangkok Hospital Pattaya.xlsx",
    5: ROOT / "data" / "BS" / "5_Setting BS Set5 Bangkok Hospital Chanthaburi.xlsx",
    6: ROOT / "data" / "BS" / "6_Setting BS Set6 Bangkok Hospital Rayong.xlsx",
}
ASSETS = {
    "program.py": ROOT / "All_Programs" / "123_Program_Run_Brandsence2026.py",
    "runner.py": BUILD / "runner.py",
    "metadata.sav": ROOT / "data" / "BS" / "SPSS_preserved_utf8_Final.sav",
}
RAW_CSV = BUILD / "test_runtime" / "rawdata.csv"
MODULE = BUILD / "BrandSenseRunner.bas"


def rgb(red, green, blue):
    return red + green * 256 + blue * 65536


def merge_value(sheet, address, value):
    area = sheet.Range(address)
    area.Merge()
    area.Cells(1, 1).Value = value
    return area


def set_shape_text(shape, text, font_size, font_color, bold=True):
    shape.TextFrame2.TextRange.Text = text
    shape.TextFrame2.TextRange.Font.Name = "Leelawadee UI"
    shape.TextFrame2.TextRange.Font.Size = font_size
    shape.TextFrame2.TextRange.Font.Bold = -1 if bold else 0
    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = font_color


def main():
    for path in [*ASSETS.values(), *SETTING_SOURCES.values(), RAW_CSV, MODULE]:
        if not path.exists():
            raise FileNotFoundError(path)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    if OUTPUT.exists():
        OUTPUT.unlink()

    rawdata = pd.read_csv(RAW_CSV, low_memory=False)
    rawdata = rawdata.astype(object).where(pd.notna(rawdata), None)
    matrix = [tuple(rawdata.columns)] + [tuple(row) for row in rawdata.itertuples(index=False, name=None)]

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
    setting_data = setting_data.astype(object).where(pd.notna(setting_data), None)
    label_data = label_data.astype(object).where(pd.notna(label_data), None)
    setting_matrix = [tuple(setting_data.columns)] + [
        tuple(row) for row in setting_data.itertuples(index=False, name=None)]
    label_matrix = [tuple(label_data.columns)] + [
        tuple(row) for row in label_data.itertuples(index=False, name=None)]

    red = rgb(198, 40, 45)
    dark_red = rgb(139, 24, 28)
    blue = rgb(79, 129, 189)
    light_blue = rgb(221, 235, 247)
    light_yellow = rgb(255, 242, 204)
    light_green = rgb(226, 239, 218)
    dark_text = rgb(48, 48, 48)
    white = rgb(255, 255, 255)

    pythoncom.CoInitialize()
    # Create a fresh dynamic COM instance without relying on pywin32's
    # generated type-library cache, which may be stale on user machines.
    excel_clsid = pythoncom.MakeIID(
        "{00024500-0000-0000-C000-000000000046}")
    excel_dispatch = pythoncom.CoCreateInstance(
        excel_clsid,
        None,
        pythoncom.CLSCTX_LOCAL_SERVER,
        pythoncom.IID_IDispatch,
    )
    def dispatch_without_cache(
            dispatch, userName=None, resultCLSID=None,
            typeinfo=None, clsctx=pythoncom.CLSCTX_SERVER):
        return win32com.client.dynamic.Dispatch(
            dispatch, userName, typeinfo=typeinfo, clsctx=clsctx)

    # Child objects returned by Excel normally route through gencache too.
    # Preserve Dispatch's public signature while forcing dynamic wrappers.
    win32com.client.Dispatch = dispatch_without_cache
    excel = win32com.client.dynamic.DumbDispatch(excel_dispatch)
    workbook = None
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        workbook = excel.Workbooks.Add()
        while workbook.Worksheets.Count > 1:
            workbook.Worksheets(workbook.Worksheets.Count).Delete()

        control = workbook.Worksheets(1)
        control.Name = "Control"
        raw = workbook.Worksheets.Add(After=control)
        raw.Name = "Rawdata"
        setting_variables = workbook.Worksheets.Add(After=raw)
        setting_variables.Name = "Setting Variables"
        setting_labels = workbook.Worksheets.Add(After=setting_variables)
        setting_labels.Name = "Setting Labels"
        engine = workbook.Worksheets.Add(After=setting_labels)
        engine.Name = "__Engine"

        control.Cells.Font.Name = "Leelawadee UI"
        control.Cells.Font.Size = 11
        control.Columns("A").ColumnWidth = 12
        control.Columns("B:G").ColumnWidth = 16
        control.Columns("H").ColumnWidth = 4
        control.Rows("1:30").RowHeight = 24
        control.Rows("1:2").RowHeight = 34

        title = merge_value(control, "A1:H2", "BrandSense — Rawdata Runner")
        title.Interior.Color = red
        title.Font.Color = white
        title.Font.Bold = True
        title.Font.Size = 22
        title.HorizontalAlignment = -4108
        title.VerticalAlignment = -4108

        intro = merge_value(
            control,
            "B4:G4",
            "วาง Rawdata แล้วกดรัน — ระบบเลือก Setting 1 / 4 / 5 / 6 ให้อัตโนมัติ",
        )
        intro.Font.Bold = True
        intro.Font.Size = 13
        intro.Font.Color = dark_red
        intro.HorizontalAlignment = -4108

        instructions = (
            "ไปที่ชีท Rawdata แล้วลบข้อมูลเดิมทั้งหมด",
            "วางหัวคอลัมน์และข้อมูล ตั้งแต่เซลล์ A1 (ห้ามเปลี่ยนชื่อตัวแปร)",
            "กลับมาชีท Control แล้วเลือกปุ่มรันด้านล่าง",
        )
        for offset, instruction in enumerate(instructions):
            row = 6 + offset
            number = control.Range(f"A{row}")
            number.Value = offset + 1
            number.Interior.Color = blue
            number.Font.Color = white
            number.Font.Bold = True
            number.HorizontalAlignment = -4108
            line = merge_value(control, f"B{row}:G{row}", instruction)
            line.Interior.Color = light_blue

        labels = ((10, "สถานะ"), (11, "โหมดล่าสุด"), (12, "เวลาล่าสุด"), (13, "Setting"))
        for row, label in labels:
            control.Range(f"A{row}").Value = label
            control.Range(f"A{row}").Font.Bold = True
            merge_value(control, f"B{row}:G{row}", "-")
        control.Range("B10").Value = "พร้อมรัน"
        control.Range("B10:G10").Interior.Color = light_green
        control.Range("B13").Value = "ตรวจจับจากชื่อคอลัมน์เมื่อกดรัน"
        control.Range("A10:G13").Borders.LineStyle = 1
        control.Range("A10:G13").Borders.Color = rgb(217, 217, 217)

        safe_button = control.Shapes.AddShape(5, 62, 345, 655, 62)
        safe_button.Name = "btnRunSafe"
        set_shape_text(
            safe_button,
            "รัน: ตัดเคส QC ทั้งหมด + Safe Mapping",
            15,
            white,
        )
        safe_button.Fill.ForeColor.RGB = red
        safe_button.Line.ForeColor.RGB = red
        safe_button.OnAction = "RunBrandSenseSafe"

        normal_button = control.Shapes.AddShape(5, 62, 420, 655, 46)
        normal_button.Name = "btnRunNormal"
        set_shape_text(
            normal_button,
            "รันปกติ (ไม่ตัด QC / Legacy Mapping)",
            12,
            dark_red,
        )
        normal_button.Fill.ForeColor.RGB = white
        normal_button.Line.ForeColor.RGB = red
        normal_button.Line.Weight = 1.5
        normal_button.OnAction = "RunBrandSenseNormal"

        note = merge_value(
            control,
            "B22:G25",
            "หมายเหตุ: ปุ่มแรกจะคัดเคสคำตอบคุณภาพต่ำออกจาก Long Format "
            "และใช้การจับ Factor แบบปลอดภัยตามโปรแกรมปัจจุบัน ส่วนปุ่มรันปกติ"
            "จะไม่ตัดเคสและคง Legacy Mapping ไว้\n"
            "เครื่องนี้ต้องมี Python และแพ็กเกจเดียวกับโปรแกรม BrandSense",
        )
        note.WrapText = True
        note.Interior.Color = light_yellow
        note.Font.Color = dark_text
        note.VerticalAlignment = -4108
        control.Rows("22:25").RowHeight = 28

        row_count = len(matrix)
        column_count = len(matrix[0])
        target = raw.Range(raw.Cells(1, 1), raw.Cells(row_count, column_count))
        target.Value = tuple(matrix)
        raw.Cells.Font.Name = "Calibri"
        raw.Cells.Font.Size = 10
        header = raw.Range(raw.Cells(1, 1), raw.Cells(1, column_count))
        header.Interior.Color = blue
        header.Font.Color = white
        header.Font.Bold = True
        header.WrapText = True
        header.HorizontalAlignment = -4108
        header.AutoFilter()
        raw.Rows(1).RowHeight = 34
        target.Borders.Color = rgb(225, 225, 225)
        target.Borders.LineStyle = 1
        raw.Columns.ColumnWidth = 12
        raw.Activate()
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

        engine.Cells(1, 1).Value = "Asset"
        engine.Cells(1, 2).Value = "Sequence"
        engine.Cells(1, 3).Value = "Base64"
        engine_row = 2
        for asset_name, path in ASSETS.items():
            encoded = base64.b64encode(path.read_bytes()).decode("ascii")
            chunks = [encoded[index : index + 4000] for index in range(0, len(encoded), 4000)]
            for sequence, chunk in enumerate(chunks, start=1):
                engine.Cells(engine_row, 1).Value = asset_name
                engine.Cells(engine_row, 2).Value = sequence
                engine.Cells(engine_row, 3).Value = chunk
                engine_row += 1
        engine.Visible = 2

        component = workbook.VBProject.VBComponents.Import(str(MODULE))
        del component

        if workbook.Worksheets(1).Name != "Control":
            control.Move(workbook.Worksheets(1))
        if workbook.Worksheets(2).Name != "Rawdata":
            raw.Move(workbook.Worksheets(2))

        workbook.SaveAs(str(OUTPUT), FileFormat=52)
        control.Activate()
        excel.ActiveWindow.DisplayGridlines = False
        excel.ActiveWindow.Zoom = 90
        workbook.Save()
        print(OUTPUT)
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=True)
            except Exception:
                pass
        try:
            excel.Quit()
        except Exception:
            pass
        del workbook
        del excel
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
