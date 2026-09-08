from pathlib import Path

import pandas as pd
import pythoncom
import win32com.client


ROOT = Path(__file__).resolve().parents[1]
BUILD = Path(__file__).resolve().parent
OUTPUT = ROOT / "outputs" / "brandsense_rawdata_xlsm" / "BrandSense_Rawdata_Auto_Set_1_4_5_6.xlsm"
RAW_CSV = BUILD / "test_runtime" / "rawdata.csv"
MODULES = [BUILD / "BSMath.bas", BUILD / "BrandSenseVBA.bas"]
SETTING_SOURCES = {
    1: ROOT / "data" / "BS" / "1_Setting BS Set1 Bangkok Hospital Headquarters.xlsx",
    4: ROOT / "data" / "BS" / "4_Setting BS Set4 Bangkok Hospital Pattaya.xlsx",
    5: ROOT / "data" / "BS" / "5_Setting BS Set5 Bangkok Hospital Chanthaburi.xlsx",
    6: ROOT / "data" / "BS" / "6_Setting BS Set6 Bangkok Hospital Rayong.xlsx",
}


def rgb(r, g, b):
    return r + g * 256 + b * 65536


def merge_value(sheet, address, value):
    area = sheet.Range(address)
    area.Merge()
    area.Cells(1, 1).Value = value
    return area


def put_matrix(sheet, matrix):
    rows, cols = len(matrix), len(matrix[0])
    sheet.Range(sheet.Cells(1, 1), sheet.Cells(rows, cols)).Value = tuple(matrix)
    return rows, cols


def style_data_sheet(sheet, rows, cols, body_color=None):
    blue, white = rgb(79, 129, 189), rgb(255, 255, 255)
    sheet.Cells.Font.Name = "Calibri"
    sheet.Cells.Font.Size = 10
    header = sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, cols))
    header.Interior.Color = blue
    header.Font.Color = white
    header.Font.Bold = True
    header.WrapText = True
    header.HorizontalAlignment = -4108
    header.AutoFilter()
    sheet.Rows(1).RowHeight = 34
    if body_color is not None and rows > 1:
        sheet.Range(sheet.Cells(2, 1), sheet.Cells(rows, cols)).Interior.Color = body_color
    sheet.Columns.AutoFit()
    for ci in range(1, cols + 1):
        sheet.Columns(ci).ColumnWidth = min(max(sheet.Columns(ci).ColumnWidth, 9), 42)


def set_shape_text(shape, value, size, color):
    shape.TextFrame2.TextRange.Text = value
    shape.TextFrame2.TextRange.Font.Name = "Leelawadee UI"
    shape.TextFrame2.TextRange.Font.Size = size
    shape.TextFrame2.TextRange.Font.Bold = -1
    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = color


def create_excel():
    clsid = pythoncom.MakeIID("{00024500-0000-0000-C000-000000000046}")
    disp = pythoncom.CoCreateInstance(clsid, None, pythoncom.CLSCTX_LOCAL_SERVER, pythoncom.IID_IDispatch)

    def dynamic_dispatch(dispatch, userName=None, resultCLSID=None, typeinfo=None, clsctx=pythoncom.CLSCTX_SERVER):
        return win32com.client.dynamic.Dispatch(dispatch, userName, typeinfo=typeinfo, clsctx=clsctx)

    win32com.client.Dispatch = dynamic_dispatch
    return win32com.client.dynamic.DumbDispatch(disp)


def settings_matrices():
    settings_all, labels_all = [], []
    filter_codes = [2, 3, 4, 5]
    filter_labels = ["Gen Z (18-27)", "Gen Y (28-44)", "Gen X (45-59)", "Baby boomer (60 - )"]
    for set_id, source in SETTING_SOURCES.items():
        settings = pd.read_excel(source, sheet_name="Settings")
        settings = settings.drop(columns=["PathFile"], errors="ignore")
        settings.insert(0, "Set_ID", set_id)
        settings_all.append(settings)
        labels = pd.read_excel(source, sheet_name="Label")
        labels = labels.reindex(range(max(len(labels), len(filter_codes))))
        labels.insert(0, "Set_ID", set_id)
        labels["Filter_Code"] = filter_codes + [None] * (len(labels) - len(filter_codes))
        labels["Filter_Label"] = filter_labels + [None] * (len(labels) - len(filter_labels))
        labels_all.append(labels)
    settings = pd.concat(settings_all, ignore_index=True).astype(object).where(lambda x: pd.notna(x), None)
    labels = pd.concat(labels_all, ignore_index=True).astype(object).where(lambda x: pd.notna(x), None)
    return (
        [tuple(settings.columns)] + [tuple(x) for x in settings.itertuples(index=False, name=None)],
        [tuple(labels.columns)] + [tuple(x) for x in labels.itertuples(index=False, name=None)],
    )


def main():
    for path in [RAW_CSV, *SETTING_SOURCES.values(), *MODULES]:
        if not path.exists():
            raise FileNotFoundError(path)
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    raw = pd.read_csv(RAW_CSV, low_memory=False).astype(object).where(lambda x: pd.notna(x), None)
    raw_matrix = [tuple(raw.columns)] + [tuple(x) for x in raw.itertuples(index=False, name=None)]
    settings_matrix, labels_matrix = settings_matrices()
    red, dark_red = rgb(198, 40, 45), rgb(139, 24, 28)
    blue, light_blue = rgb(79, 129, 189), rgb(221, 235, 247)
    yellow, green, white = rgb(255, 242, 204), rgb(226, 239, 218), rgb(255, 255, 255)
    pythoncom.CoInitialize()
    excel, book = None, None
    try:
        excel = create_excel()
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        book = excel.Workbooks.Add()
        while book.Worksheets.Count > 1:
            book.Worksheets(book.Worksheets.Count).Delete()
        control = book.Worksheets(1)
        control.Name = "Control"
        raw_ws = book.Worksheets.Add(After=control)
        raw_ws.Name = "Rawdata"
        vars_ws = book.Worksheets.Add(After=raw_ws)
        vars_ws.Name = "Setting Variables"
        labels_ws = book.Worksheets.Add(After=vars_ws)
        labels_ws.Name = "Setting Labels"

        control.Cells.Font.Name = "Leelawadee UI"
        control.Cells.Font.Size = 11
        control.Columns("A").ColumnWidth = 12
        control.Columns("B:G").ColumnWidth = 16
        control.Columns("H").ColumnWidth = 4
        control.Rows("1:30").RowHeight = 24
        control.Rows("1:2").RowHeight = 34
        title = merge_value(control, "A1:H2", "BrandSense — Excel 100% Rawdata Runner")
        title.Interior.Color = red
        title.Font.Color = white
        title.Font.Bold = True
        title.Font.Size = 22
        title.HorizontalAlignment = -4108
        title.VerticalAlignment = -4108
        intro = merge_value(control, "B4:G4", "วาง Rawdata แล้วกดรัน — ไม่ใช้ Python และไม่เรียก SPSS")
        intro.Font.Bold = True
        intro.Font.Size = 13
        intro.Font.Color = dark_red
        intro.HorizontalAlignment = -4108
        instructions = (
            "ไปที่ชีท Rawdata แล้วลบข้อมูลเดิมทั้งหมด",
            "วางหัวคอลัมน์และข้อมูลตั้งแต่ A1 โดยคงชื่อตัวแปรเดิม",
            "กลับมาชีท Control แล้วเลือกปุ่มรัน",
        )
        for offset, instruction in enumerate(instructions):
            row = 6 + offset
            control.Range(f"A{row}").Value = offset + 1
            control.Range(f"A{row}").Interior.Color = blue
            control.Range(f"A{row}").Font.Color = white
            control.Range(f"A{row}").Font.Bold = True
            control.Range(f"A{row}").HorizontalAlignment = -4108
            line = merge_value(control, f"B{row}:G{row}", instruction)
            line.Interior.Color = light_blue
        for row, label in ((10, "สถานะ"), (11, "โหมดล่าสุด"), (12, "เวลาล่าสุด"), (13, "Setting")):
            control.Range(f"A{row}").Value = label
            control.Range(f"A{row}").Font.Bold = True
            merge_value(control, f"B{row}:G{row}", "-")
        control.Range("B10").Value = "พร้อมรัน — VBA ภายในไฟล์"
        control.Range("B10:G10").Interior.Color = green
        control.Range("B13").Value = "ตรวจจับ Set 1 / 4 / 5 / 6 จากหัว Rawdata"
        control.Range("A10:G13").Borders.LineStyle = 1
        control.Range("A10:G13").Borders.Color = rgb(217, 217, 217)

        safe = control.Shapes.AddShape(5, 62, 345, 655, 62)
        safe.Name = "btnRunSafe"
        set_shape_text(safe, "รัน: ตัดเคส QC ทั้งหมด + Safe Mapping", 15, white)
        safe.Fill.ForeColor.RGB = red
        safe.Line.ForeColor.RGB = red
        safe.OnAction = "RunBrandSenseSafe"
        normal = control.Shapes.AddShape(5, 62, 420, 655, 46)
        normal.Name = "btnRunNormal"
        set_shape_text(normal, "รันปกติ (ไม่ตัด QC / Legacy Mapping)", 12, dark_red)
        normal.Fill.ForeColor.RGB = white
        normal.Line.ForeColor.RGB = red
        normal.OnAction = "RunBrandSenseNormal"
        note = merge_value(
            control,
            "B20:G24",
            "Setting Variables และ Setting Labels ฝังอยู่ในไฟล์นี้และแก้ไขได้โดยตรง\n"
            "Engine เป็น VBA 100%: Wide-to-Long, QC, PCA/Equamax, Anderson-Rubin, Regression, Summary และ Agree/T2B ทำใน Excel ทั้งหมด",
        )
        note.WrapText = True
        note.Interior.Color = yellow
        note.Font.Color = dark_red
        note.VerticalAlignment = -4108

        rr, rc = put_matrix(raw_ws, raw_matrix)
        style_data_sheet(raw_ws, rr, rc)
        raw_ws.Columns.ColumnWidth = 12
        vr, vc = put_matrix(vars_ws, settings_matrix)
        style_data_sheet(vars_ws, vr, vc, yellow)
        lr, lc = put_matrix(labels_ws, labels_matrix)
        style_data_sheet(labels_ws, lr, lc, yellow)
        vars_ws.Tab.Color = rgb(255, 192, 0)
        labels_ws.Tab.Color = rgb(255, 217, 102)

        for module in MODULES:
            book.VBProject.VBComponents.Import(str(module))
        if OUTPUT.exists():
            OUTPUT.unlink()
        book.SaveAs(str(OUTPUT), FileFormat=52)
        control.Activate()
        excel.ActiveWindow.DisplayGridlines = False
        excel.ActiveWindow.Zoom = 90
        book.Save()
        print(OUTPUT)
    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=True)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
