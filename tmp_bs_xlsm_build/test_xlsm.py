from pathlib import Path

import pythoncom
import win32com.client


ROOT = Path(__file__).resolve().parents[1]
WORKBOOK_PATH = (
    ROOT / "outputs" / "brandsense_rawdata_xlsm" /
    "BrandSense_Rawdata_Auto_Set_1_4_5_6.xlsm")


def dispatch_without_cache(
        dispatch, userName=None, resultCLSID=None,
        typeinfo=None, clsctx=pythoncom.CLSCTX_SERVER):
    return win32com.client.dynamic.Dispatch(
        dispatch, userName, typeinfo=typeinfo, clsctx=clsctx)


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
    excel.AutomationSecurity = 1
    workbook = excel.Workbooks.Open(str(WORKBOOK_PATH))
    excel.Run(f"'{workbook.Name}'!RunBrandSenseSafe")
    status = workbook.Worksheets("Control").Range("B10").Value
    setting = workbook.Worksheets("Control").Range("B13").Value
    sheets = [workbook.Worksheets(index).Name for index in range(1, workbook.Worksheets.Count + 1)]
    workbook.Save()
    print({"status": status, "setting": setting, "sheets": sheets})
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
