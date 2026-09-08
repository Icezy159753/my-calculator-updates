from pathlib import Path


path = Path(__file__).resolve().parent / "tmp_bs_xlsm_build" / "build_xlsm.py"
text = path.read_text(encoding="utf-8")
old = '''    excel = win32com.client.dynamic.Dispatch(excel_dispatch)
    workbook = None'''
new = '''    def dispatch_without_cache(
            dispatch, userName=None, resultCLSID=None,
            typeinfo=None, clsctx=pythoncom.CLSCTX_SERVER):
        return win32com.client.dynamic.Dispatch(
            dispatch, userName, typeinfo=typeinfo, clsctx=clsctx)

    # Child objects returned by Excel normally route through gencache too.
    # Preserve Dispatch's public signature while forcing dynamic wrappers.
    win32com.client.Dispatch = dispatch_without_cache
    excel = win32com.client.dynamic.DumbDispatch(excel_dispatch)
    workbook = None'''
assert text.count(old) == 1
path.write_text(text.replace(old, new, 1), encoding="utf-8")
print("patched dynamic child COM wrapping")
