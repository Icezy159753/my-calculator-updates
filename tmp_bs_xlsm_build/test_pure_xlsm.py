from pathlib import Path
import json
import pythoncom
import win32com.client

ROOT = Path(__file__).resolve().parents[1]
BOOK = ROOT / "outputs" / "brandsense_rawdata_xlsm" / "BrandSense_Rawdata_Auto_Set_1_4_5_6.xlsm"
RESULT = Path(__file__).resolve().parent / "pure_test_result.json"


def excel_instance():
    clsid = pythoncom.MakeIID("{00024500-0000-0000-C000-000000000046}")
    disp = pythoncom.CoCreateInstance(clsid, None, pythoncom.CLSCTX_LOCAL_SERVER, pythoncom.IID_IDispatch)
    def dynamic_dispatch(dispatch, userName=None, resultCLSID=None, typeinfo=None, clsctx=pythoncom.CLSCTX_SERVER):
        return win32com.client.dynamic.Dispatch(dispatch, userName, typeinfo=typeinfo, clsctx=clsctx)
    win32com.client.Dispatch = dynamic_dispatch
    return win32com.client.dynamic.DumbDispatch(disp)


pythoncom.CoInitialize()
excel = None
book = None
result = {}
try:
    excel = excel_instance()
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.AutomationSecurity = 1
    book = excel.Workbooks.Open(str(BOOK), ReadOnly=False)
    excel.Run(f"'{book.Name}'!RunBrandSenseSafe")
    book.Save()
    result["status"] = book.Worksheets("Control").Range("B10").Value
    result["sheets"] = [book.Worksheets(i).Name for i in range(1, book.Worksheets.Count + 1)]
    summary = book.Worksheets("Summary")
    result["summary_rows"] = summary.UsedRange.Rows.Count
    result["summary_cols"] = summary.UsedRange.Columns.Count
    result["overall"] = [summary.Cells(2, i).Value for i in range(1, 16)]
    qc = book.Worksheets("QC Excluded")
    result["qc_rows"] = qc.UsedRange.Rows.Count
    result["modules"] = [book.VBProject.VBComponents(i).Name for i in range(1, book.VBProject.VBComponents.Count + 1)]
    result["hidden_sheets"] = [book.Worksheets(i).Name for i in range(1, book.Worksheets.Count + 1) if book.Worksheets(i).Visible != -1]
finally:
    if book is not None:
        book.Close(SaveChanges=True)
    if excel is not None:
        excel.Quit()
    pythoncom.CoUninitialize()

RESULT.write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
