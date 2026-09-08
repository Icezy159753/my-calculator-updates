from pathlib import Path
import pythoncom
import win32com.client

book_path = Path(__file__).resolve().parents[1] / "outputs" / "brandsense_rawdata_xlsm" / "BrandSense_Rawdata_Auto_Set_1_4_5_6.xlsm"
pythoncom.CoInitialize()
clsid = pythoncom.MakeIID("{00024500-0000-0000-C000-000000000046}")
dispatch = pythoncom.CoCreateInstance(clsid, None, pythoncom.CLSCTX_LOCAL_SERVER, pythoncom.IID_IDispatch)


def dynamic_dispatch(dispatch, userName=None, resultCLSID=None, typeinfo=None, clsctx=pythoncom.CLSCTX_SERVER):
    return win32com.client.dynamic.Dispatch(dispatch, userName, typeinfo=typeinfo, clsctx=clsctx)


win32com.client.Dispatch = dynamic_dispatch
excel = win32com.client.dynamic.DumbDispatch(dispatch)
book = None
try:
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.AutomationSecurity = 1
    book = excel.Workbooks.Open(str(book_path), ReadOnly=False)
    excel.Run(f"'{book.Name}'!RunBrandSenseNormal")
    print(book.Worksheets("Control").Range("B10").Value)
    print(book.Worksheets("Control").Range("B11").Value)
    print(book.Worksheets("Summary").UsedRange.Rows.Count, book.Worksheets("Summary").UsedRange.Columns.Count)
finally:
    if book is not None:
        book.Close(SaveChanges=False)
    excel.Quit()
    pythoncom.CoUninitialize()
