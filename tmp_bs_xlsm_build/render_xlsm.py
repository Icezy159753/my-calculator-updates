from pathlib import Path

import pythoncom
import win32com.client


ROOT = Path(__file__).resolve().parents[1]
WORKBOOK_PATH = (
    ROOT / "outputs" / "brandsense_rawdata_xlsm" /
    "BrandSense_Rawdata_Auto_Set_1_4_5_6.xlsm")
RENDER_DIR = Path(__file__).resolve().parent / "renders"
RANGES = {
    "Control": "A1:H26",
    "Rawdata": "A1:P18",
    "Summary": "A1:O26",
    "SandP": "A1:L20",
    "Correspondence(S)": "A1:N20",
    "Correspondence(P)": "A1:N20",
    "QC Excluded": "A1:K20",
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
        sheet.Activate()
        capture = sheet.Range(address)
        capture.CopyPicture(1, 2)
        chart_object = sheet.ChartObjects().Add(
            0, 0, max(300, capture.Width), max(200, capture.Height))
        chart_object.Chart.Paste()
        safe_name = sheet_name.replace("(", "_").replace(")", "_").replace(" ", "_")
        output = RENDER_DIR / f"{safe_name}.png"
        chart_object.Chart.Export(str(output))
        chart_object.Delete()
        print(output)
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
