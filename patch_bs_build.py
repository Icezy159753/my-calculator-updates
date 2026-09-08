from pathlib import Path


path = Path(__file__).resolve().parent / "tmp_bs_xlsm_build" / "build_xlsm.py"
text = path.read_text(encoding="utf-8")
old = '    excel = win32com.client.DispatchEx("Excel.Application")'
new = '''    # Create a fresh dynamic COM instance without relying on pywin32's
    # generated type-library cache, which may be stale on user machines.
    excel_clsid = pythoncom.MakeIID(
        "{00024500-0000-0000-C000-000000000046}")
    excel_dispatch = pythoncom.CoCreateInstance(
        excel_clsid,
        None,
        pythoncom.CLSCTX_LOCAL_SERVER,
        pythoncom.IID_IDispatch,
    )
    excel = win32com.client.dynamic.Dispatch(excel_dispatch)'''
assert text.count(old) == 1
path.write_text(text.replace(old, new, 1), encoding="utf-8")
print("patched build_xlsm.py dynamic Excel COM creation")
