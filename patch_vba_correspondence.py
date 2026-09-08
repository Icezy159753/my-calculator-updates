from pathlib import Path

p = Path("tmp_bs_xlsm_build/BrandSenseVBA.bas")
s = p.read_text(encoding="utf-8")
s = s.replace(
    '''    WriteCorrespondencePlaceholder "Correspondence(S)", "S"
    WriteCorrespondencePlaceholder "Correspondence(P)", "P"''',
    '''    WriteCorrespondence st, lng, "Correspondence(S)", "S", UsedGroups(st.S)
    WriteCorrespondence st, lng, "Correspondence(P)", "P", UsedGroups(st.P)''',
)
start = s.index('Private Sub WriteCorrespondencePlaceholder(')
end = s.index('Private Sub StyleTable', start)
replacement = '''Private Sub WriteCorrespondence(ByRef st As BSSetting, ByRef lng As BSLong, ByVal sheetName As String, ByVal prefix As String, ByVal groups As Variant)
    Dim ws As Worksheet, crosses As Collection, block As Long, typ As Long, crossValue As Variant, title As String, colOffset As Long
    DeleteSheetIfExists sheetName
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)): ws.Name = sheetName
    Set crosses = UniqueCrossValues(lng)
    For block = 0 To crosses.Count
        If block = 0 Then typ = 0: crossValue = "": title = "Total" Else typ = 2: crossValue = crosses(block): title = CrossLabel(st, crossValue)
        WriteCABlock ws, colOffset + 1, title, prefix, groups, typ, crossValue, st, lng
        colOffset = colOffset + 5
    Next block
    ws.Tab.Color = RGB(255, 0, 0)
End Sub

Private Sub WriteCABlock(ByVal ws As Worksheet, ByVal startCol As Long, ByVal title As String, ByVal prefix As String, ByVal groups As Variant, ByVal typ As Long, ByVal crossValue As Variant, ByRef st As BSSetting, ByRef lng As BSLong)
    Dim m As Long, k As Long, cont As Variant, massR() As Double, massC() As Double
    Dim prob As Variant, smat As Variant, gram As Variant, vals As Variant, vecs As Variant
    Dim i As Long, j As Long, q As Long, row As Long, actual As Long, n As Long, count As Long
    Dim total As Double, mean As Double, sv(1 To 2) As Double, evTotal As Double
    Dim u As Variant, rowScore As Variant, colScore As Variant, value As Double
    m = GroupArrayCount(groups): k = UBound(st.IndexCodes)
    If m < 2 Or k < 2 Then Exit Sub
    cont = BSMat(m, k)
    For i = 1 To m
        actual = groups(i)
        For j = 1 To k
            total = 0#: count = 0
            For q = 1 To lng.Count
                If GroupMatch(typ, 0, crossValue, lng, q) And lng.IndexCode(q) = st.IndexCodes(j) Then
                    If prefix = "S" Then value = lng.SVal(q, actual) Else value = lng.PVal(q, actual)
                    If value <> MISS Then total = total + value: count = count + 1
                End If
            Next q
            If count > 0 Then cont(i, j) = total / count
        Next j
    Next i
    total = 0#: For i = 1 To m: For j = 1 To k: total = total + cont(i, j): Next j, i
    If total <= 0# Then Exit Sub
    ReDim massR(1 To m): ReDim massC(1 To k): prob = BSMat(m, k): smat = BSMat(m, k)
    For i = 1 To m
        For j = 1 To k
            prob(i, j) = cont(i, j) / total: massR(i) = massR(i) + prob(i, j): massC(j) = massC(j) + prob(i, j)
        Next j
    Next i
    For i = 1 To m
        For j = 1 To k
            If massR(i) > 0# And massC(j) > 0# Then smat(i, j) = (prob(i, j) - massR(i) * massC(j)) / Sqr(massR(i) * massC(j))
        Next j
    Next i
    gram = BSMultiply(BSTranspose(smat), smat): BSJacobiEigen gram, vals, vecs: BSSortEigenDesc vals, vecs
    For j = 1 To k: If vals(j) > 0# Then evTotal = evTotal + vals(j)
    Next j
    For j = 1 To 2: If vals(j) > 0# Then sv(j) = Sqr(vals(j))
    Next j
    u = BSMat(m, 2): rowScore = BSMat(m, 2): colScore = BSMat(k, 2)
    For i = 1 To m
        For j = 1 To 2
            If sv(j) > 0# Then
                For q = 1 To k: u(i, j) = u(i, j) + smat(i, q) * vecs(q, j) / sv(j): Next q
                If massR(i) > 0# Then rowScore(i, j) = u(i, j) * sv(j) / Sqr(massR(i))
            End If
        Next j
    Next i
    For i = 1 To k
        For j = 1 To 2: If massC(i) > 0# Then colScore(i, j) = vecs(i, j) * sv(j) / Sqr(massC(i))
        Next j
    Next i
    With ws.Cells(1, startCol): .Value = title: .Interior.Color = RGB(255, 215, 0): .Font.Bold = True: End With
    row = 2: CASection ws, row, startCol, "Axis information": row = row + 1
    ws.Cells(row, startCol + 1).Resize(1, 3).Value = Array("Singular value", "Eigen value", "Contribution ratio"): CAHeader ws, row, startCol, 4
    For j = 1 To 2
        row = row + 1: ws.Cells(row, startCol).Value = "Axis" & j: ws.Cells(row, startCol + 1).Value = sv(j): ws.Cells(row, startCol + 2).Value = vals(j)
        If evTotal > 0# Then ws.Cells(row, startCol + 3).Value = vals(j) / evTotal
    Next j
    row = row + 3: CASection ws, row, startCol, "Row category score": row = row + 1
    ws.Cells(row, startCol + 1).Value = "Axis1": ws.Cells(row, startCol + 2).Value = "Axis2": CAHeader ws, row, startCol, 4
    For i = 1 To m
        row = row + 1: ws.Cells(row, startCol).Value = prefix & "_" & groups(i): ws.Cells(row, startCol + 1).Value = rowScore(i, 1): ws.Cells(row, startCol + 2).Value = rowScore(i, 2)
    Next i
    row = row + 3: CASection ws, row, startCol, "Column category score": row = row + 1
    ws.Cells(row, startCol + 1).Value = "Axis1": ws.Cells(row, startCol + 2).Value = "Axis2": CAHeader ws, row, startCol, 4
    For i = 1 To k
        row = row + 1: ws.Cells(row, startCol).Value = "(" & st.IndexLabels(i) & ")": ws.Cells(row, startCol + 1).Value = colScore(i, 1): ws.Cells(row, startCol + 2).Value = colScore(i, 2)
    Next i
    With ws.Range(ws.Cells(1, startCol), ws.Cells(row, startCol + 3)).Borders: .LineStyle = xlContinuous: .Color = RGB(160, 160, 160): .Weight = xlThin: End With
    ws.Range(ws.Cells(2, startCol + 1), ws.Cells(row, startCol + 3)).NumberFormat = "0.0000000"
    ws.Columns(startCol).ColumnWidth = 40: ws.Columns(startCol + 1).Resize(, 3).ColumnWidth = 18
End Sub

Private Sub CASection(ByVal ws As Worksheet, ByVal row As Long, ByVal col As Long, ByVal textValue As String)
    ws.Cells(row, col).Value = textValue
    With ws.Range(ws.Cells(row, col), ws.Cells(row, col + 3)): .Interior.Color = RGB(255, 218, 185): .Font.Bold = True: End With
End Sub

Private Sub CAHeader(ByVal ws As Worksheet, ByVal row As Long, ByVal col As Long, ByVal width As Long)
    With ws.Range(ws.Cells(row, col), ws.Cells(row, col + width - 1)): .Interior.Color = RGB(255, 218, 185): .Font.Bold = True: .HorizontalAlignment = xlCenter: End With
End Sub

'''
s = s[:start] + replacement + s[end:]
p.write_text(s, encoding="utf-8")
