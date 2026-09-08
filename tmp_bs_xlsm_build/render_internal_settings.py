from pathlib import Path

import fitz
import pythoncom
import win32com.client


ROOT = Path(__file__).resolve().parents[1]
WORKBOOK_PATH = (
    ROOT / "outputs" / "brandsense_rawdata_xlsm" /
    "BrandSense_Rawdata_Auto_Set_1_4_5_6.xlsm")
RENDER_DIR = Path(__file__).resolve().parent / "internal_setting_renders"
RANGES = {
    "Control": "A1:H26",
    "Setting Variables": "A1:L22",
    "Setting Labels": "A1:E15",
}


def dispatch_without_cache(
        dispatch, userName=None, resultCLSID=None,
        typeinfo=None, clsctx=pythoncom.CLSCTX_SERVER):
    return win32com.client.dynamic.Dispatch(
        dispatch, userName, typeinfo=typeinfo, clsctx=clsctx)


RENDER_DIR.mkdir(parents=True, exist_ok=True)
pythoncom.CoInitialize()
win32com.client.Dispatch = dispatch_without_cache
clsid = pythoncom.MakeIID("{00024500-0000-0000-C000-000000000046}")
dispatch = pythoncom.CoCreateInstance(
    clsid, None, pythoncom.CLSCTX_LOCAL_SERVER, pythoncom.IID_IDispatch)
excel = win32com.client.dynamic.DumbDispatch(dispatch)
workbook = None
try:
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = excel.Workbooks.Open(str(WORKBOOK_PATH), 0, True)
    for sheet_name, address in RANGES.items():
        sheet = workbook.Worksheets(sheet_name)
        setup = sheet.PageSetup
        setup.PrintArea = address
        setup.Orientation = 2
        setup.Zoom = False
        setup.FitToPagesWide = 1
        setup.FitToPagesTall = 1
        safe_name = sheet_name.replace(" ", "_")
        pdf_path = RENDER_DIR / f"{safe_name}.pdf"
        png_path = RENDER_DIR / f"{safe_name}.png"
        sheet.ExportAsFixedFormat(0, str(pdf_path), 0, True, False)
        document = fitz.open(pdf_path)
        pixmap = document[0].get_pixmap(dpi=170, alpha=False)
        pixmap.save(png_path)
        document.close()
        print(png_path)
finally:
    if workbook is not None:
        try:
            workbook.Close(False)
        except Exception:
            pass
    try:
        excel.Quit()
    except Exception:
        pass
    del workbook
    del excel
    pythoncom.CoUninitialize()
