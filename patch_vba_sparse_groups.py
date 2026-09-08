from pathlib import Path

p = Path("tmp_bs_xlsm_build/BrandSenseVBA.bas")
s = p.read_text(encoding="utf-8")

s = s.replace(
    'attrTotal = sN + pN: pTotal = pN',
    'attrTotal = CountUsedGroups(mapS) + CountUsedGroups(mapP): pTotal = CountUsedGroups(mapP)',
)

anchor = 'Private Function MaxGroup(ByRef vars() As String) As Long\n'
insert = '''Private Function UsedGroups(ByRef vars() As String) As Variant
    Dim d As Object, i As Long, g As Long, idx As Long, keys As Variant, out() As Long, j As Long, t As Variant
    Set d = CreateObject("Scripting.Dictionary")
    For i = LBound(vars) To UBound(vars)
        If Len(vars(i)) > 0 Then
            ParseSPE vars(i), g, idx
            If g > 0 Then If Not d.Exists(CStr(g)) Then d.Add CStr(g), g
        End If
    Next i
    If d.Count = 0 Then ReDim out(0 To 0): UsedGroups = out: Exit Function
    keys = d.Keys
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If CLng(keys(j)) < CLng(keys(i)) Then t = keys(i): keys(i) = keys(j): keys(j) = t
        Next j
    Next i
    ReDim out(1 To d.Count)
    For i = 0 To UBound(keys): out(i + 1) = CLng(keys(i)): Next i
    UsedGroups = out
End Function

Private Function GroupArrayCount(ByVal groups As Variant) As Long
    If LBound(groups) = 0 And UBound(groups) = 0 Then GroupArrayCount = 0 Else GroupArrayCount = UBound(groups) - LBound(groups) + 1
End Function

Private Function CountUsedGroups(ByRef map() As Long) As Long
    Dim g As Long, i As Long, found As Boolean
    For g = LBound(map, 1) To UBound(map, 1)
        found = False
        For i = LBound(map, 2) To UBound(map, 2)
            If map(g, i) > 0 Then found = True: Exit For
        Next i
        If found Then CountUsedGroups = CountUsedGroups + 1
    Next g
End Function

'''
s = s.replace(anchor, insert + anchor)

s = s.replace(
    'Dim sN As Long, pN As Long, cN As Long, eN As Long, headersOut As Variant, col As Long',
    'Dim sGroups As Variant, pGroups As Variant, cGroups As Variant, eGroups As Variant, headersOut As Variant, col As Long',
)
s = s.replace(
    '    sN = MaxGroup(st.S): pN = MaxGroup(st.P): cN = MaxGroup(st.C): eN = MaxGroup(st.E)\n    totalCols = 15 + sN + pN + cN + eN + eN + sN + pN + sN + pN',
    '    sGroups = UsedGroups(st.S): pGroups = UsedGroups(st.P): cGroups = UsedGroups(st.C): eGroups = UsedGroups(st.E)\n    totalCols = 15 + GroupArrayCount(sGroups) + GroupArrayCount(pGroups) + GroupArrayCount(cGroups) + GroupArrayCount(eGroups) + GroupArrayCount(eGroups) + GroupArrayCount(sGroups) + GroupArrayCount(pGroups) + GroupArrayCount(sGroups) + GroupArrayCount(pGroups)',
)
s = s.replace(
    '    AddSeriesHeaders "S_", sN, headersOut, col: AddSeriesHeaders "P_", pN, headersOut, col\n    AddSeriesHeaders "C_", cN, headersOut, col: AddSeriesHeaders "E_", eN, headersOut, col\n    AddSeriesHeaders "CorE_", eN, headersOut, col: AddSeriesHeaders "cor_S_", sN, headersOut, col\n    AddSeriesHeaders "cor_P_", pN, headersOut, col: AddSeriesHeaders "agree_S_", sN, headersOut, col\n    AddSeriesHeaders "agree_P_", pN, headersOut, col',
    '    AddGroupHeaders "S_", sGroups, headersOut, col: AddGroupHeaders "P_", pGroups, headersOut, col\n    AddGroupHeaders "C_", cGroups, headersOut, col: AddGroupHeaders "E_", eGroups, headersOut, col\n    AddGroupHeaders "CorE_", eGroups, headersOut, col: AddGroupHeaders "cor_S_", sGroups, headersOut, col\n    AddGroupHeaders "cor_P_", pGroups, headersOut, col: AddGroupHeaders "agree_S_", sGroups, headersOut, col\n    AddGroupHeaders "agree_P_", pGroups, headersOut, col',
)
s = s.replace(
    'ComputeSummaryRow ws, outRow, CStr(groupNames(g)), CLng(groupTypes(g)), CLng(groupIdx(g)), groupCross(g), raw, headers, st, lng, safeMode, sN, pN, cN, eN',
    'ComputeSummaryRow ws, outRow, CStr(groupNames(g)), CLng(groupTypes(g)), CLng(groupIdx(g)), groupCross(g), raw, headers, st, lng, safeMode, sGroups, pGroups, cGroups, eGroups',
)
s = s.replace(
    '''Private Sub AddSeriesHeaders(ByVal prefix As String, ByVal n As Long, ByRef out As Variant, ByRef col As Long)
    Dim i As Long: For i = 1 To n: col = col + 1: out(1, col) = prefix & i: Next i
End Sub''',
    '''Private Sub AddGroupHeaders(ByVal prefix As String, ByVal groups As Variant, ByRef out As Variant, ByRef col As Long)
    Dim i As Long
    If GroupArrayCount(groups) = 0 Then Exit Sub
    For i = LBound(groups) To UBound(groups): col = col + 1: out(1, col) = prefix & groups(i): Next i
End Sub''',
)
s = s.replace(
    'ByVal safeMode As Boolean, ByVal sN As Long, ByVal pN As Long, ByVal cN As Long, ByVal eN As Long)',
    'ByVal safeMode As Boolean, ByVal sGroups As Variant, ByVal pGroups As Variant, ByVal cGroups As Variant, ByVal eGroups As Variant)',
)
s = s.replace(
    'sMeans = CategoryMeans(lng.SVal, sN, typ, indexCode, crossValue, lng)\n    pMeans = CategoryMeans(lng.PVal, pN, typ, indexCode, crossValue, lng)\n    cMeans = CategoryMeans(lng.CVal, cN, typ, indexCode, crossValue, lng)\n    eMeans = CategoryMeans(lng.EVal, eN, typ, indexCode, crossValue, lng)',
    'sMeans = CategoryMeansSelected(lng.SVal, sGroups, typ, indexCode, crossValue, lng)\n    pMeans = CategoryMeansSelected(lng.PVal, pGroups, typ, indexCode, crossValue, lng)\n    cMeans = CategoryMeansSelected(lng.CVal, cGroups, typ, indexCode, crossValue, lng)\n    eMeans = CategoryMeansSelected(lng.EVal, eGroups, typ, indexCode, crossValue, lng)',
)
s = s.replace(
    'WriteCorrelationRow ws, outRow, col, lng.EVal, eN, lng.A, typ, indexCode, crossValue, lng\n    WriteCorrelationRow ws, outRow, col, lng.SVal, sN, lng.A, typ, indexCode, crossValue, lng\n    WriteCorrelationRow ws, outRow, col, lng.PVal, pN, lng.A, typ, indexCode, crossValue, lng\n    mStage = "Agree/T2B - " & filterName\n    WriteAgreeRow ws, outRow, col, st.AgreeS, sN, raw, headers, resp, st.T2BChoice\n    WriteAgreeRow ws, outRow, col, st.AgreeP, pN, raw, headers, resp, st.T2BChoice',
    'WriteCorrelationSelected ws, outRow, col, lng.EVal, eGroups, lng.A, typ, indexCode, crossValue, lng\n    WriteCorrelationSelected ws, outRow, col, lng.SVal, sGroups, lng.A, typ, indexCode, crossValue, lng\n    WriteCorrelationSelected ws, outRow, col, lng.PVal, pGroups, lng.A, typ, indexCode, crossValue, lng\n    mStage = "Agree/T2B - " & filterName\n    WriteAgreeRow ws, outRow, col, st.AgreeS, GroupArrayCount(sGroups), raw, headers, resp, st.T2BChoice\n    WriteAgreeRow ws, outRow, col, st.AgreeP, GroupArrayCount(pGroups), raw, headers, resp, st.T2BChoice',
)

start = s.index('Private Function CategoryMeans(')
end = s.index('Private Function MeanArray', start)
s = s[:start] + '''Private Function CategoryMeansSelected(ByRef values() As Double, ByVal groups As Variant, ByVal typ As Long, ByVal indexCode As Long, ByVal crossValue As Variant, ByRef lng As BSLong) As Variant
    Dim out() As Double, i As Long, pos As Long, actual As Long, sums() As Double, counts() As Long, nGroups As Long
    nGroups = GroupArrayCount(groups)
    If nGroups = 0 Then ReDim out(0 To 0): CategoryMeansSelected = out: Exit Function
    ReDim out(1 To nGroups): ReDim sums(1 To nGroups): ReDim counts(1 To nGroups)
    For i = 1 To lng.Count
        If GroupMatch(typ, indexCode, crossValue, lng, i) Then
            For pos = 1 To nGroups
                actual = groups(pos)
                If values(i, actual) <> MISS Then sums(pos) = sums(pos) + values(i, actual): counts(pos) = counts(pos) + 1
            Next pos
        End If
    Next i
    For pos = 1 To nGroups: If counts(pos) > 0 Then out(pos) = sums(pos) / counts(pos) Else out(pos) = MISS
    Next pos
    CategoryMeansSelected = out
End Function

''' + s[end:]

start = s.index('Private Sub WriteCorrelationRow(')
end = s.index('Private Sub WriteAgreeRow', start)
s = s[:start] + '''Private Sub WriteCorrelationSelected(ByVal ws As Worksheet, ByVal outRow As Long, ByRef col As Long, ByRef values() As Double, ByVal groups As Variant, ByRef target() As Double, ByVal typ As Long, ByVal indexCode As Long, ByVal crossValue As Variant, ByRef lng As BSLong)
    Dim pos As Long, actual As Long, i As Long, n As Long, sx As Double, sy As Double, sxx As Double, syy As Double, sxy As Double, x As Double, y As Double, den As Double
    If GroupArrayCount(groups) = 0 Then Exit Sub
    For pos = LBound(groups) To UBound(groups)
        actual = groups(pos): n = 0: sx = 0#: sy = 0#: sxx = 0#: syy = 0#: sxy = 0#
        For i = 1 To lng.Count
            If GroupMatch(typ, indexCode, crossValue, lng, i) And values(i, actual) <> MISS And target(i) <> MISS Then
                x = values(i, actual): y = target(i): n = n + 1: sx = sx + x: sy = sy + y: sxx = sxx + x * x: syy = syy + y * y: sxy = sxy + x * y
            End If
        Next i
        col = col + 1: den = (n * sxx - sx * sx) * (n * syy - sy * sy)
        If n >= 2 And den > 0# Then ws.Cells(outRow, col).Value = Abs((n * sxy - sx * sy) / Sqr(den))
    Next pos
End Sub

''' + s[end:]

old = '''        If g >= LBound(vars) And g <= UBound(vars) And Len(vars(g)) > 0 And headers.Exists(vars(g)) And total > 0 Then
            For Each key In resp.Keys'''
new = '''        If g >= LBound(vars) And g <= UBound(vars) Then
            If Len(vars(g)) > 0 And headers.Exists(vars(g)) And total > 0 Then
            For Each key In resp.Keys'''
s = s.replace(old, new)
s = s.replace(
    '''            ws.Cells(outRow, col).Value = countGood / total * 100#
        End If
    Next g
End Sub''',
    '''            ws.Cells(outRow, col).Value = countGood / total * 100#
            End If
        End If
    Next g
End Sub''',
    1,
)

p.write_text(s, encoding="utf-8")
