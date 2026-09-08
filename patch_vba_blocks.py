from pathlib import Path

p = Path("tmp_bs_xlsm_build/BSMath.bas")
s = p.read_text(encoding="utf-8")
s = s.replace(
    '    If total > 0# Then For i = 1 To p: ratios(i) = Abs(ratios(i)) / total * 100#: Next i',
    '    If total > 0# Then\n        For i = 1 To p\n            ratios(i) = Abs(ratios(i)) / total * 100#\n        Next i\n    End If',
)
p.write_text(s, encoding="utf-8")

p = Path("tmp_bs_xlsm_build/BrandSenseVBA.bas")
s = p.read_text(encoding="utf-8")
s = s.replace(
    '    Dim i As Long, j As Long: ReDim a(1 To n, 1 To IIf(p > 0, p, 1))\n    For i = 1 To n: For j = 1 To IIf(p > 0, p, 1): a(i, j) = MISS: Next j, i',
    '    Dim i As Long, j As Long: ReDim a(1 To n, 1 To IIf(p > 0, p, 1))\n    For i = 1 To n\n        For j = 1 To IIf(p > 0, p, 1)\n            a(i, j) = MISS\n        Next j\n    Next i',
)
s = s.replace(
    '    For i = LBound(map, 1) To UBound(map, 1): For j = LBound(map, 2) To UBound(map, 2): If map(i, j) > 0 Then CountMapped = CountMapped + 1\n    Next j, i',
    '    For i = LBound(map, 1) To UBound(map, 1)\n        For j = LBound(map, 2) To UBound(map, 2)\n            If map(i, j) > 0 Then CountMapped = CountMapped + 1\n        Next j\n    Next i',
)
s = s.replace(
    '        If lng.Active(i) And Not lng.Excluded(i) Then key = CStr(lng.CrossValue(i)): If Len(key) > 0 Then If Not d.Exists(key) Then d.Add key, lng.CrossValue(i)',
    '        If lng.Active(i) And Not lng.Excluded(i) Then\n            key = CStr(lng.CrossValue(i))\n            If Len(key) > 0 Then\n                If Not d.Exists(key) Then d.Add key, lng.CrossValue(i)\n            End If\n        End If',
)
s = s.replace(
    '        For i = 0 To UBound(keys) - 1: For j = i + 1 To UBound(keys): If Val(keys(j)) < Val(keys(i)) Then t = keys(i): keys(i) = keys(j): keys(j) = t\n        Next j, i',
    '        For i = 0 To UBound(keys) - 1\n            For j = i + 1 To UBound(keys)\n                If Val(keys(j)) < Val(keys(i)) Then t = keys(i): keys(i) = keys(j): keys(j) = t\n            Next j\n        Next i',
)
s = s.replace(
    '    If regN >= 4 Then ratios = BSFactorRatios(x, y, safeMode, collision) Else ReDim ratios(1 To 4)',
    '    If regN >= 4 Then\n        ratios = BSFactorRatios(x, y, safeMode, collision)\n    Else\n        ReDim ratios(1 To 4)\n    End If',
)
s = s.replace(
    '        If GroupMatch(typ, indexCode, crossValue, lng, i) Then For g = 1 To groups\n            If values(i, g) <> MISS Then sums(g) = sums(g) + values(i, g): counts(g) = counts(g) + 1\n        Next g',
    '        If GroupMatch(typ, indexCode, crossValue, lng, i) Then\n            For g = 1 To groups\n                If values(i, g) <> MISS Then sums(g) = sums(g) + values(i, g): counts(g) = counts(g) + 1\n            Next g\n        End If',
)
p.write_text(s, encoding="utf-8")
