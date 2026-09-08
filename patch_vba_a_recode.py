from pathlib import Path

p = Path("tmp_bs_xlsm_build/BrandSenseVBA.bas")
s = p.read_text(encoding="utf-8")
old = 'If mapA(i) > 0 Then lng.A(rowL) = NumericOrMiss(raw(r, mapA(i))): If lng.A(rowL) <> MISS Then lng.ZA(rowL) = ZAValue(lng.A(rowL))'
new = '''If mapA(i) > 0 Then
                lng.A(rowL) = NumericOrMiss(raw(r, mapA(i)))
                If lng.A(rowL) = 9 Then lng.A(rowL) = 0
                If lng.A(rowL) <> MISS Then lng.ZA(rowL) = ZAValue(lng.A(rowL))
            End If'''
if old not in s:
    raise SystemExit("A recode anchor not found")
s = s.replace(old, new)
p.write_text(s, encoding="utf-8")
