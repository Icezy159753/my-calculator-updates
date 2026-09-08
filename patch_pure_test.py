from pathlib import Path

p = Path("tmp_bs_xlsm_build/test_pure_xlsm.py")
s = p.read_text(encoding="utf-8")
s = s.replace(
    '    disp = pythoncom.CoCreateInstance(clsid, None, pythoncom.CLSCTX_LOCAL_SERVER, pythoncom.IID_IDispatch)\n    return win32com.client.dynamic.DumbDispatch(disp)',
    '    disp = pythoncom.CoCreateInstance(clsid, None, pythoncom.CLSCTX_LOCAL_SERVER, pythoncom.IID_IDispatch)\n    def dynamic_dispatch(dispatch, userName=None, resultCLSID=None, typeinfo=None, clsctx=pythoncom.CLSCTX_SERVER):\n        return win32com.client.dynamic.Dispatch(dispatch, userName, typeinfo=typeinfo, clsctx=clsctx)\n    win32com.client.Dispatch = dynamic_dispatch\n    return win32com.client.dynamic.DumbDispatch(disp)',
)
p.write_text(s, encoding="utf-8")
