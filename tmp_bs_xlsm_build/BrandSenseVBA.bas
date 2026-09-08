Attribute VB_Name = "BrandSenseVBA"
Option Explicit

Private Const CONTROL_SHEET As String = "Control"
Private Const RAW_SHEET As String = "Rawdata"
Private Const VAR_SHEET As String = "Setting Variables"
Private Const LABEL_SHEET As String = "Setting Labels"
Private Const MISS As Double = 1E+99
Private Const LOW_MODEL_N As Long = 30
Private mStage As String

Private Type BSSetting
    SetID As Long
    T2BChoice As String
    CrossFilter As String
    C() As String
    A() As String
    S() As String
    P() As String
    E() As String
    AgreeS() As String
    AgreeP() As String
    IndexCodes() As Long
    IndexLabels() As String
End Type

Private Type BSLong
    Count As Long
    Respondent() As Long
    IndexCode() As Long
    CrossValue() As Variant
    Active() As Boolean
    Excluded() As Boolean
    A() As Double
    ZA() As Double
    NS() As Double
    NP() As Double
    NC() As Double
    NE() As Double
    SVal() As Double
    PVal() As Double
    CVal() As Double
    EVal() As Double
End Type

Public Sub RunBrandSenseSafe()
    RunBrandSense True
End Sub

Public Sub RunBrandSenseNormal()
    RunBrandSense False
End Sub

Private Sub RunBrandSense(ByVal safeMode As Boolean)
    Dim oldCalc As XlCalculation, raw As Variant, headers As Object
    Dim st As BSSetting, lng As BSLong, setID As Long
    On Error GoTo Failed
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False: Application.EnableEvents = False
    Application.DisplayAlerts = False: Application.Calculation = xlCalculationManual
    mStage = "Read Rawdata / Setting"
    UpdateStatus "กำลังอ่าน Rawdata และ Setting ภายใน Excel...", RGB(255, 242, 204)
    raw = ReadRawdata(headers)
    setID = DetectSetting(headers)
    If setID = 0 Then Err.Raise vbObjectError + 801, , "ไม่พบ Setting ที่ตรงกับหัวตัวแปรใน Rawdata"
    LoadSetting setID, st
    ValidateSetting st, headers
    mStage = "Build Long / QC"
    UpdateStatus "กำลังแปลง Wide เป็น Long Format...", RGB(221, 235, 247)
    BuildLong raw, headers, st, lng, safeMode
    mStage = "Write outputs"
    UpdateStatus "กำลังคำนวณ Factor / Regression / Summary...", RGB(226, 239, 218)
    WriteOutputs raw, headers, st, lng, safeMode
    With ThisWorkbook.Worksheets(CONTROL_SHEET)
        .Range("B11").Value = IIf(safeMode, "QC All + Safe Mapping", "Normal / Legacy")
        .Range("B12").Value = Now: .Range("B12").NumberFormat = "yyyy-mm-dd hh:mm:ss"
        .Range("B13").Value = "Set " & st.SetID & " (Setting ภายในไฟล์)"
    End With
    UpdateStatus "เสร็จสมบูรณ์ — Excel/VBA 100%", RGB(198, 239, 206)
    ThisWorkbook.Save
CleanExit:
    Application.Calculation = oldCalc: Application.DisplayAlerts = True
    Application.EnableEvents = True: Application.ScreenUpdating = True
    Exit Sub
Failed:
    UpdateStatus "ผิดพลาดที่ " & mStage & ": " & Err.Description, RGB(255, 199, 206)
    If Application.Visible Then MsgBox Err.Description, vbCritical, "BrandSense Excel 100%"
    Resume CleanExit
End Sub

Private Function ReadRawdata(ByRef headers As Object) As Variant
    Dim ws As Worksheet, lr As Long, lc As Long, a As Variant, j As Long, key As String
    Set ws = ThisWorkbook.Worksheets(RAW_SHEET)
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lr < 2 Or lc < 2 Then Err.Raise vbObjectError + 802, , "Rawdata ต้องมีหัวคอลัมน์และข้อมูลอย่างน้อย 1 แถว"
    a = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc)).Value2
    Set headers = CreateObject("Scripting.Dictionary"): headers.CompareMode = vbTextCompare
    For j = 1 To lc
        key = Trim$(CStr(a(1, j)))
        If Len(key) > 0 Then headers(key) = j
    Next j
    ReadRawdata = a
End Function

Private Function DetectSetting(ByVal headers As Object) As Long
    Dim ws As Worksheet, data As Variant, h As Object, r As Long, c As Long
    Dim sid As Long, score As Object, k As Variant, v As String, best As Long, bestScore As Long
    Set ws = ThisWorkbook.Worksheets(VAR_SHEET)
    data = ws.UsedRange.Value2: Set h = HeaderMap(data): Set score = CreateObject("Scripting.Dictionary")
    For r = 2 To UBound(data, 1)
        If IsNumeric(data(r, h("Set_ID"))) Then
            sid = CLng(data(r, h("Set_ID")))
            If Not score.Exists(CStr(sid)) Then score(CStr(sid)) = 0
            For Each k In Array("C", "A", "S", "P", "E")
                If h.Exists(CStr(k)) Then
                    v = Trim$(CStr(data(r, h(CStr(k)))))
                    If Len(v) > 0 And headers.Exists(v) Then score(CStr(sid)) = score(CStr(sid)) + 1
                End If
            Next k
        End If
    Next r
    For Each k In score.Keys
        If CLng(score(k)) > bestScore Then bestScore = CLng(score(k)): best = CLng(k)
    Next k
    DetectSetting = best
End Function

Private Function HeaderMap(ByVal data As Variant) As Object
    Dim d As Object, j As Long
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    For j = 1 To UBound(data, 2): d(Trim$(CStr(data(1, j)))) = j: Next j
    Set HeaderMap = d
End Function

Private Sub LoadSetting(ByVal setID As Long, ByRef st As BSSetting)
    Dim data As Variant, labels As Variant, h As Object, hl As Object
    Dim r As Long, sid As Long, k As Variant, v As String
    Dim lists As Object, col As String, idx As Long
    data = ThisWorkbook.Worksheets(VAR_SHEET).UsedRange.Value2: Set h = HeaderMap(data)
    Set lists = CreateObject("Scripting.Dictionary")
    For Each k In Array("C", "A", "S", "P", "E", "AgreeS", "AgreeP")
        lists.Add CStr(k), NewCollection()
    Next k
    st.SetID = setID: st.T2BChoice = "5+4"
    For r = 2 To UBound(data, 1)
        If IsNumeric(data(r, h("Set_ID"))) Then sid = CLng(data(r, h("Set_ID"))) Else sid = -1
        If sid = setID Then
            If h.Exists("T2B_Choice") And Len(Trim$(CStr(data(r, h("T2B_Choice"))))) > 0 Then st.T2BChoice = Trim$(CStr(data(r, h("T2B_Choice"))))
            If h.Exists("Filter_Var") And Len(Trim$(CStr(data(r, h("Filter_Var"))))) > 0 Then st.CrossFilter = Trim$(CStr(data(r, h("Filter_Var"))))
            For Each k In lists.Keys
                col = CStr(k)
                If h.Exists(col) Then
                    v = Trim$(CStr(data(r, h(col))))
                    If Len(v) > 0 Then AddUnique lists(col), v
                End If
            Next k
        End If
    Next r
    CollectionToStringArray lists("C"), st.C
    CollectionToStringArray lists("A"), st.A
    CollectionToStringArray lists("S"), st.S
    CollectionToStringArray lists("P"), st.P
    CollectionToStringArray lists("E"), st.E
    CollectionToStringArray lists("AgreeS"), st.AgreeS
    CollectionToStringArray lists("AgreeP"), st.AgreeP
    labels = ThisWorkbook.Worksheets(LABEL_SHEET).UsedRange.Value2: Set hl = HeaderMap(labels)
    idx = 0
    For r = 2 To UBound(labels, 1)
        If IsNumeric(labels(r, hl("Set_ID"))) And CLng(labels(r, hl("Set_ID"))) = setID Then
            If IsNumeric(labels(r, hl("Index1_Code"))) And Len(Trim$(CStr(labels(r, hl("Index1_Label"))))) > 0 Then idx = idx + 1
        End If
    Next r
    If idx = 0 Then Err.Raise vbObjectError + 803, , "Setting Labels ไม่มี Index1 สำหรับ Set " & setID
    ReDim st.IndexCodes(1 To idx): ReDim st.IndexLabels(1 To idx): idx = 0
    For r = 2 To UBound(labels, 1)
        If IsNumeric(labels(r, hl("Set_ID"))) And CLng(labels(r, hl("Set_ID"))) = setID Then
            If IsNumeric(labels(r, hl("Index1_Code"))) And Len(Trim$(CStr(labels(r, hl("Index1_Label"))))) > 0 Then
                idx = idx + 1: st.IndexCodes(idx) = CLng(labels(r, hl("Index1_Code")))
                st.IndexLabels(idx) = CStr(labels(r, hl("Index1_Label")))
            End If
        End If
    Next r
End Sub

Private Function NewCollection() As Collection
    Set NewCollection = New Collection
End Function

Private Sub AddUnique(ByVal c As Collection, ByVal textValue As String)
    On Error Resume Next: c.Add textValue, LCase$(textValue): On Error GoTo 0
End Sub

Private Sub CollectionToStringArray(ByVal c As Collection, ByRef out() As String)
    Dim i As Long
    If c.Count = 0 Then ReDim out(0 To 0): Exit Sub
    ReDim out(1 To c.Count)
    For i = 1 To c.Count: out(i) = CStr(c(i)): Next i
End Sub

Private Sub ValidateSetting(ByRef st As BSSetting, ByVal headers As Object)
    Dim i As Long, missing As String
    If Len(st.CrossFilter) > 0 And Not headers.Exists(st.CrossFilter) Then missing = missing & st.CrossFilter & vbCrLf
    For i = LBound(st.A) To UBound(st.A): If Len(st.A(i)) > 0 And Not headers.Exists(st.A(i)) Then missing = missing & st.A(i) & vbCrLf
    Next i
    For i = LBound(st.S) To UBound(st.S): If Len(st.S(i)) > 0 And Not headers.Exists(st.S(i)) Then missing = missing & st.S(i) & vbCrLf
    Next i
    For i = LBound(st.P) To UBound(st.P): If Len(st.P(i)) > 0 And Not headers.Exists(st.P(i)) Then missing = missing & st.P(i) & vbCrLf
    Next i
    For i = LBound(st.E) To UBound(st.E): If Len(st.E(i)) > 0 And Not headers.Exists(st.E(i)) Then missing = missing & st.E(i) & vbCrLf
    Next i
    If Len(missing) > 0 Then Err.Raise vbObjectError + 804, , "Rawdata ขาดตัวแปรที่ Setting ต้องใช้:" & vbCrLf & Left$(missing, 1200)
End Sub

Private Sub ParseSPE(ByVal varName As String, ByRef groupNo As Long, ByRef indexNo As Long)
    Dim pHash As Long, pDollar As Long
    pHash = InStrRev(varName, "#"): pDollar = InStrRev(varName, "$")
    groupNo = 0: indexNo = 0
    If pHash > 0 And pDollar > pHash Then
        groupNo = Val(Mid$(varName, pHash + 1, pDollar - pHash - 1))
        indexNo = Val(Mid$(varName, pDollar + 1))
    End If
End Sub

Private Function ParseAIndex(ByVal varName As String) As Long
    Dim p As Long: p = InStrRev(varName, "#")
    If p > 0 Then ParseAIndex = Val(Mid$(varName, p + 1))
End Function

Private Function FindIndexPosition(ByRef st As BSSetting, ByVal code As Long) As Long
    Dim i As Long
    For i = LBound(st.IndexCodes) To UBound(st.IndexCodes)
        If st.IndexCodes(i) = code Then FindIndexPosition = i: Exit Function
    Next i
End Function

Private Function UsedGroups(ByRef vars() As String) As Variant
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

Private Function MaxGroup(ByRef vars() As String) As Long
    Dim i As Long, g As Long, idx As Long
    For i = LBound(vars) To UBound(vars)
        If Len(vars(i)) > 0 Then ParseSPE vars(i), g, idx: If g > MaxGroup Then MaxGroup = g
    Next i
End Function

Private Function NumericOrMiss(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Or Len(Trim$(CStr(v))) = 0 Or Not IsNumeric(v) Then NumericOrMiss = MISS Else NumericOrMiss = CDbl(v)
End Function

Private Function ZAValue(ByVal a As Double) As Double
    Select Case CLng(a)
        Case 0, 9: ZAValue = 0#
        Case 1: ZAValue = 0.05
        Case 2: ZAValue = 0.12
        Case 3: ZAValue = 0.27
        Case 4: ZAValue = 0.5
        Case 5: ZAValue = 0.73
        Case 6: ZAValue = 0.88
        Case 7: ZAValue = 0.95
        Case 8: ZAValue = 1#
        Case Else: ZAValue = MISS
    End Select
End Function

Private Sub InitMissing2D(ByRef a() As Double, ByVal n As Long, ByVal p As Long)
    Dim i As Long, j As Long: ReDim a(1 To n, 1 To IIf(p > 0, p, 1))
    For i = 1 To n
        For j = 1 To IIf(p > 0, p, 1)
            a(i, j) = MISS
        Next j
    Next i
End Sub

Private Sub BuildLong(ByVal raw As Variant, ByVal headers As Object, ByRef st As BSSetting, ByRef lng As BSLong, ByVal safeMode As Boolean)
    Dim nResp As Long, nIdx As Long, nLong As Long, sN As Long, pN As Long, cN As Long, eN As Long
    Dim mapA() As Long, mapS() As Long, mapP() As Long, mapC() As Long, mapE() As Long
    Dim r As Long, i As Long, j As Long, g As Long, idx As Long, pos As Long, q As Long, rowL As Long
    Dim v As Double, aQC As Double, sum As Double, cnt As Long, attrCount As Long, pCount As Long, attrTotal As Long, pTotal As Long
    Dim hard As Boolean, review As Boolean, idCol As Long, keyCol As Long, reason As String
    nResp = UBound(raw, 1) - 1: nIdx = UBound(st.IndexCodes): nLong = nResp * nIdx
    sN = MaxGroup(st.S): pN = MaxGroup(st.P): cN = MaxGroup(st.C): eN = MaxGroup(st.E)
    ReDim mapA(1 To nIdx): ReDim mapS(1 To IIf(sN > 0, sN, 1), 1 To nIdx)
    ReDim mapP(1 To IIf(pN > 0, pN, 1), 1 To nIdx): ReDim mapC(1 To IIf(cN > 0, cN, 1), 1 To nIdx)
    ReDim mapE(1 To IIf(eN > 0, eN, 1), 1 To nIdx)
    For i = LBound(st.A) To UBound(st.A)
        If Len(st.A(i)) > 0 Then idx = ParseAIndex(st.A(i)): pos = FindIndexPosition(st, idx): If pos > 0 Then mapA(pos) = headers(st.A(i))
    Next i
    FillVarMap st.S, st, headers, mapS: FillVarMap st.P, st, headers, mapP
    FillVarMap st.C, st, headers, mapC: FillVarMap st.E, st, headers, mapE
    lng.Count = nLong: ReDim lng.Respondent(1 To nLong): ReDim lng.IndexCode(1 To nLong)
    ReDim lng.CrossValue(1 To nLong): ReDim lng.Active(1 To nLong): ReDim lng.Excluded(1 To nLong)
    ReDim lng.A(1 To nLong): ReDim lng.ZA(1 To nLong): ReDim lng.NS(1 To nLong)
    ReDim lng.NP(1 To nLong): ReDim lng.NC(1 To nLong): ReDim lng.NE(1 To nLong)
    InitMissing1D lng.A, nLong: InitMissing1D lng.ZA, nLong: InitMissing1D lng.NS, nLong
    InitMissing1D lng.NP, nLong: InitMissing1D lng.NC, nLong: InitMissing1D lng.NE, nLong
    InitMissing2D lng.SVal, nLong, sN: InitMissing2D lng.PVal, nLong, pN
    InitMissing2D lng.CVal, nLong, cN: InitMissing2D lng.EVal, nLong, eN
    DeleteSheetIfExists "QC Excluded": CreateQCSheet
    If headers.Exists("SBJNUM") Then idCol = headers("SBJNUM")
    If headers.Exists("KEY") Then keyCol = headers("KEY")
    attrTotal = CountUsedGroups(mapS) + CountUsedGroups(mapP): pTotal = CountUsedGroups(mapP)
    For r = 2 To UBound(raw, 1)
        For i = 1 To nIdx
            rowL = rowL + 1: aQC = MISS: lng.Respondent(rowL) = r - 1: lng.IndexCode(rowL) = st.IndexCodes(i)
            If Len(st.CrossFilter) > 0 Then lng.CrossValue(rowL) = raw(r, headers(st.CrossFilter))
            If mapA(i) > 0 Then
                aQC = NumericOrMiss(raw(r, mapA(i)))
                lng.A(rowL) = aQC
                If lng.A(rowL) = 9 Then lng.A(rowL) = 0
                If lng.A(rowL) <> MISS Then lng.ZA(rowL) = ZAValue(lng.A(rowL))
            End If
            FillLongCategory raw, r, rowL, i, mapS, lng.SVal, sN, lng.NS(rowL)
            FillLongCategory raw, r, rowL, i, mapP, lng.PVal, pN, lng.NP(rowL)
            FillLongCategory raw, r, rowL, i, mapE, lng.EVal, eN, lng.NE(rowL)
            FillLongC raw, r, rowL, i, mapC, lng.CVal, cN, nIdx, lng.NC(rowL)
            lng.Active(rowL) = (lng.A(rowL) <> MISS Or lng.NS(rowL) <> MISS Or lng.NP(rowL) <> MISS Or lng.NC(rowL) <> MISS Or lng.NE(rowL) <> MISS)
            If lng.Active(rowL) Then
                attrCount = PositiveMappedCount(raw, r, i, mapS) + PositiveMappedCount(raw, r, i, mapP)
                pCount = PositiveMappedCount(raw, r, i, mapP): hard = False: review = False
                If aQC <> MISS Then
                    hard = ((aQC = 1 And attrCount > 0) Or (aQC = 2 And attrCount >= Application.WorksheetFunction.RoundUp(attrTotal * 0.85, 0)) Or (aQC >= 6 And aQC <= 8 And attrCount = 0))
                    review = ((aQC <= 3 And pCount >= Application.WorksheetFunction.RoundUp(pTotal * 0.85, 0)) Or (aQC = 3 And attrCount >= Application.WorksheetFunction.RoundUp(attrTotal * 0.9, 0)))
                End If
                If hard Or review Then
                    reason = QCReason(aQC, attrCount, attrTotal, pCount, pTotal)
                    AppendQC st, raw, r, idCol, keyCol, lng.IndexCode(rowL), IIf(hard, "HARD", "REVIEW"), reason, attrCount, attrTotal, pCount, pTotal, safeMode
                    If safeMode Then lng.Excluded(rowL) = True
                End If
            End If
        Next i
    Next r
    FormatQCSheet
End Sub

Private Sub InitMissing1D(ByRef a() As Double, ByVal n As Long)
    Dim i As Long: ReDim a(1 To n): For i = 1 To n: a(i) = MISS: Next i
End Sub

Private Sub FillVarMap(ByRef vars() As String, ByRef st As BSSetting, ByVal headers As Object, ByRef map() As Long)
    Dim i As Long, g As Long, idx As Long, pos As Long
    For i = LBound(vars) To UBound(vars)
        If Len(vars(i)) > 0 Then
            ParseSPE vars(i), g, idx: pos = FindIndexPosition(st, idx)
            If g > 0 And pos > 0 And headers.Exists(vars(i)) Then map(g, pos) = headers(vars(i))
        End If
    Next i
End Sub

Private Sub FillLongCategory(ByVal raw As Variant, ByVal rawRow As Long, ByVal longRow As Long, ByVal idxPos As Long, ByRef map() As Long, ByRef dest() As Double, ByVal groups As Long, ByRef nMean As Double)
    Dim g As Long, v As Double, sum As Double, cnt As Long
    If groups = 0 Then nMean = MISS: Exit Sub
    For g = 1 To groups
        If map(g, idxPos) > 0 Then
            v = NumericOrMiss(raw(rawRow, map(g, idxPos)))
            If v <> MISS Then dest(longRow, g) = v: sum = sum + v: cnt = cnt + 1
        End If
    Next g
    If cnt > 0 Then nMean = sum / cnt Else nMean = MISS
End Sub

Private Sub FillLongC(ByVal raw As Variant, ByVal rawRow As Long, ByVal longRow As Long, ByVal idxPos As Long, ByRef map() As Long, ByRef dest() As Double, ByVal groups As Long, ByVal nIdx As Long, ByRef nMean As Double)
    Dim g As Long, k As Long, main As Double, other As Double, sumOther As Double, cntOther As Long, cval As Double, total As Double, cnt As Long
    If groups = 0 Then nMean = MISS: Exit Sub
    For g = 1 To groups
        If map(g, idxPos) > 0 Then
            main = NumericOrMiss(raw(rawRow, map(g, idxPos)))
            If main <> MISS Then
                sumOther = 0#: cntOther = 0
                For k = 1 To nIdx
                    If k <> idxPos And map(g, k) > 0 Then
                        other = NumericOrMiss(raw(rawRow, map(g, k)))
                        If other <> MISS Then sumOther = sumOther + other: cntOther = cntOther + 1
                    End If
                Next k
                If cntOther > 0 Then cval = ((main - sumOther / cntOther) + 1#) / 2#: dest(longRow, g) = cval: total = total + cval: cnt = cnt + 1
            End If
        End If
    Next g
    If cnt > 0 Then nMean = total / cnt Else nMean = MISS
End Sub

Private Function CountMapped(ByRef map() As Long) As Long
    Dim i As Long, j As Long
    For i = LBound(map, 1) To UBound(map, 1)
        For j = LBound(map, 2) To UBound(map, 2)
            If map(i, j) > 0 Then CountMapped = CountMapped + 1
        Next j
    Next i
End Function

Private Function PositiveMappedCount(ByVal raw As Variant, ByVal r As Long, ByVal idxPos As Long, ByRef map() As Long) As Long
    Dim g As Long, v As Double
    For g = LBound(map, 1) To UBound(map, 1)
        If map(g, idxPos) > 0 Then v = NumericOrMiss(raw(r, map(g, idxPos))): If v <> MISS And v > 0 Then PositiveMappedCount = PositiveMappedCount + 1
    Next g
End Function

Private Function QCReason(ByVal a As Double, ByVal ac As Long, ByVal at As Long, ByVal pc As Long, ByVal pt As Long) As String
    If a = 1 And ac > 0 Then QCReason = "A=1 แต่เลือก Attribute"
    If a = 2 And ac >= Application.WorksheetFunction.RoundUp(at * 0.85, 0) Then QCReason = "A=2 แต่เลือก Attribute เกือบทั้งหมด"
    If a >= 6 And a <= 8 And ac = 0 Then QCReason = "A สูง แต่ไม่เลือก Attribute"
    If a <= 3 And pc >= Application.WorksheetFunction.RoundUp(pt * 0.85, 0) Then QCReason = QCReason & IIf(Len(QCReason) > 0, " | ", "") & "A ต่ำ แต่เลือก P เกือบทั้งหมด"
    If a = 3 And ac >= Application.WorksheetFunction.RoundUp(at * 0.9, 0) Then QCReason = QCReason & IIf(Len(QCReason) > 0, " | ", "") & "A=3 แต่เลือก Attribute เกือบทั้งหมด"
End Function

Private Sub CreateQCSheet()
    Dim ws As Worksheet, headers As Variant, j As Long
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)): ws.Name = "QC Excluded"
    headers = Array("Setting_File", "SBJNUM", "KEY", "Index1", "Index1_Label", "QC_Level", "QC_Reason", "A_Original", "Attribute_Count", "Attribute_Total", "P_Count", "P_Total", "Rows_Removed", "QC_Action")
    For j = 0 To UBound(headers): ws.Cells(1, j + 1).Value = headers(j): Next j
End Sub

Private Sub AppendQC(ByRef st As BSSetting, ByVal raw As Variant, ByVal r As Long, ByVal idCol As Long, ByVal keyCol As Long, ByVal indexCode As Long, ByVal level As String, ByVal reason As String, ByVal ac As Long, ByVal at As Long, ByVal pc As Long, ByVal pt As Long, ByVal safeMode As Boolean)
    Dim ws As Worksheet, nr As Long, aCol As Long, i As Long, idxLabel As String, aVal As Variant
    Set ws = ThisWorkbook.Worksheets("QC Excluded"): nr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    For i = 1 To UBound(st.IndexCodes): If st.IndexCodes(i) = indexCode Then idxLabel = st.IndexLabels(i): Exit For
    Next i
    For i = LBound(st.A) To UBound(st.A): If ParseAIndex(st.A(i)) = indexCode Then aCol = FindRawHeader(st.A(i)): Exit For
    Next i
    If aCol > 0 Then aVal = raw(r, aCol)
    ws.Cells(nr, 1).Resize(1, 14).Value = Array("Set " & st.SetID, IIf(idCol > 0, raw(r, idCol), r - 1), IIf(keyCol > 0, raw(r, keyCol), ""), indexCode, idxLabel, level, reason, aVal, ac, at, pc, pt, IIf(safeMode, 1, 0), IIf(safeMode, "Excluded", "Flag only"))
End Sub

Private Function FindRawHeader(ByVal name As String) As Long
    Dim j As Long, ws As Worksheet: Set ws = ThisWorkbook.Worksheets(RAW_SHEET)
    For j = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If StrComp(Trim$(CStr(ws.Cells(1, j).Value2)), name, vbTextCompare) = 0 Then FindRawHeader = j: Exit Function
    Next j
End Function

Private Sub FormatQCSheet()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("QC Excluded")
    StyleTable ws, ws.Cells(ws.Rows.Count, 1).End(xlUp).Row, 14
    ws.Columns("A:N").AutoFit: ws.Tab.Color = RGB(255, 192, 0)
End Sub

Private Sub WriteOutputs(ByVal raw As Variant, ByVal headers As Object, ByRef st As BSSetting, ByRef lng As BSLong, ByVal safeMode As Boolean)
    DeleteSheetIfExists "Summary": DeleteSheetIfExists "SandP"
    DeleteSheetIfExists "Correspondence(S)": DeleteSheetIfExists "Correspondence(P)"
    WriteSummary raw, headers, st, lng, safeMode
    WriteSandP st
    WriteCorrespondence st, lng, "Correspondence(S)", "S", UsedGroups(st.S)
    WriteCorrespondence st, lng, "Correspondence(P)", "P", UsedGroups(st.P)
    ReorderSheets
End Sub

Private Sub WriteSummary(ByVal raw As Variant, ByVal headers As Object, ByRef st As BSSetting, ByRef lng As BSLong, ByVal safeMode As Boolean)
    Dim ws As Worksheet, crossValues As Collection, groupNames As Collection, groupTypes As Collection, groupIdx As Collection, groupCross As Collection
    Dim i As Long, j As Long, k As Long, g As Long, r As Long, outRow As Long, totalCols As Long
    Dim sGroups As Variant, pGroups As Variant, cGroups As Variant, eGroups As Variant, headersOut As Variant, col As Long
    Set crossValues = UniqueCrossValues(lng): Set groupNames = New Collection: Set groupTypes = New Collection: Set groupIdx = New Collection: Set groupCross = New Collection
    AddGroup groupNames, groupTypes, groupIdx, groupCross, "Overall", 0, 0, ""
    For i = 1 To UBound(st.IndexCodes): AddGroup groupNames, groupTypes, groupIdx, groupCross, "Index1=" & st.IndexCodes(i), 1, st.IndexCodes(i), "": Next i
    For j = 1 To crossValues.Count
        AddGroup groupNames, groupTypes, groupIdx, groupCross, st.CrossFilter & "=" & CrossLabel(st, crossValues(j)), 2, 0, crossValues(j)
        For i = 1 To UBound(st.IndexCodes): AddGroup groupNames, groupTypes, groupIdx, groupCross, "Index1=" & st.IndexCodes(i) & "+" & st.CrossFilter & "=" & CrossLabel(st, crossValues(j)), 3, st.IndexCodes(i), crossValues(j): Next i
    Next j
    sGroups = UsedGroups(st.S): pGroups = UsedGroups(st.P): cGroups = UsedGroups(st.C): eGroups = UsedGroups(st.E)
    totalCols = 15 + GroupArrayCount(sGroups) + GroupArrayCount(pGroups) + GroupArrayCount(cGroups) + GroupArrayCount(eGroups) + GroupArrayCount(eGroups) + GroupArrayCount(sGroups) + GroupArrayCount(pGroups) + GroupArrayCount(sGroups) + GroupArrayCount(pGroups)
    ReDim headersOut(1 To 1, 1 To totalCols): col = 0
    EachHeader Array("Code Index1", "Labe Index1", "SampleSize", "Filter", "S", "P", "A level", "A score", "Index", "C", "E", "B.S", "B.P", "B.C", "B.E"), headersOut, col
    AddGroupHeaders "S_", sGroups, headersOut, col: AddGroupHeaders "P_", pGroups, headersOut, col
    AddGroupHeaders "C_", cGroups, headersOut, col: AddGroupHeaders "E_", eGroups, headersOut, col
    AddGroupHeaders "CorE_", eGroups, headersOut, col: AddGroupHeaders "cor_S_", sGroups, headersOut, col
    AddGroupHeaders "cor_P_", pGroups, headersOut, col: AddGroupHeaders "agree_S_", sGroups, headersOut, col
    AddGroupHeaders "agree_P_", pGroups, headersOut, col
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(RAW_SHEET)): ws.Name = "Summary"
    ws.Cells(1, 1).Resize(1, totalCols).Value = headersOut
    outRow = 1
    For g = 1 To groupNames.Count
        mStage = "Summary group " & g & "/" & groupNames.Count & " - " & CStr(groupNames(g))
        outRow = outRow + 1
        ComputeSummaryRow ws, outRow, CStr(groupNames(g)), CLng(groupTypes(g)), CLng(groupIdx(g)), groupCross(g), raw, headers, st, lng, safeMode, sGroups, pGroups, cGroups, eGroups
    Next g
    StyleSummary ws, outRow, totalCols
End Sub

Private Sub AddGroup(ByVal names As Collection, ByVal types As Collection, ByVal idx As Collection, ByVal cross As Collection, ByVal name As String, ByVal typ As Long, ByVal indexCode As Long, ByVal crossValue As Variant)
    names.Add name: types.Add typ: idx.Add indexCode: cross.Add crossValue
End Sub

Private Function UniqueCrossValues(ByRef lng As BSLong) As Collection
    Dim d As Object, i As Long, c As Collection, key As String, keys As Variant, a As Variant, j As Long, t As Variant
    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To lng.Count
        If lng.Active(i) And Not lng.Excluded(i) Then
            key = CStr(lng.CrossValue(i))
            If Len(key) > 0 Then
                If Not d.Exists(key) Then d.Add key, lng.CrossValue(i)
            End If
        End If
    Next i
    keys = d.Keys
    If d.Count > 1 Then
        For i = 0 To UBound(keys) - 1
            For j = i + 1 To UBound(keys)
                If Val(keys(j)) < Val(keys(i)) Then t = keys(i): keys(i) = keys(j): keys(j) = t
            Next j
        Next i
    End If
    Set c = New Collection: For i = 0 To d.Count - 1: c.Add d(keys(i)): Next i: Set UniqueCrossValues = c
End Function

Private Function CrossLabel(ByRef st As BSSetting, ByVal value As Variant) As String
    Dim ws As Worksheet, a As Variant, h As Object, r As Long
    Set ws = ThisWorkbook.Worksheets(LABEL_SHEET): a = ws.UsedRange.Value2: Set h = HeaderMap(a)
    If h.Exists("Filter_Code") And h.Exists("Filter_Label") Then
        For r = 2 To UBound(a, 1)
            If IsNumeric(a(r, h("Set_ID"))) And CLng(a(r, h("Set_ID"))) = st.SetID Then
                If CStr(a(r, h("Filter_Code"))) = CStr(value) And Len(Trim$(CStr(a(r, h("Filter_Label"))))) > 0 Then CrossLabel = CStr(a(r, h("Filter_Label"))): Exit Function
            End If
        Next r
    End If
    CrossLabel = CStr(value)
End Function

Private Sub EachHeader(ByVal names As Variant, ByRef out As Variant, ByRef col As Long)
    Dim x As Variant: For Each x In names: col = col + 1: out(1, col) = x: Next x
End Sub

Private Sub AddGroupHeaders(ByVal prefix As String, ByVal groups As Variant, ByRef out As Variant, ByRef col As Long)
    Dim i As Long
    If GroupArrayCount(groups) = 0 Then Exit Sub
    For i = LBound(groups) To UBound(groups): col = col + 1: out(1, col) = prefix & groups(i): Next i
End Sub

Private Function GroupMatch(ByVal typ As Long, ByVal indexCode As Long, ByVal crossValue As Variant, ByRef lng As BSLong, ByVal row As Long) As Boolean
    If Not lng.Active(row) Or lng.Excluded(row) Then Exit Function
    Select Case typ
        Case 0: GroupMatch = True
        Case 1: GroupMatch = (lng.IndexCode(row) = indexCode)
        Case 2: GroupMatch = (CStr(lng.CrossValue(row)) = CStr(crossValue))
        Case 3: GroupMatch = (lng.IndexCode(row) = indexCode And CStr(lng.CrossValue(row)) = CStr(crossValue))
    End Select
End Function

Private Sub ComputeSummaryRow(ByVal ws As Worksheet, ByVal outRow As Long, ByVal filterName As String, ByVal typ As Long, ByVal indexCode As Long, ByVal crossValue As Variant, ByVal raw As Variant, ByVal headers As Object, ByRef st As BSSetting, ByRef lng As BSLong, ByVal safeMode As Boolean, ByVal sGroups As Variant, ByVal pGroups As Variant, ByVal cGroups As Variant, ByVal eGroups As Variant)
    Dim resp As Object, regN As Long, n As Long, i As Long, j As Long, col As Long, label As String
    Dim sMeans As Variant, pMeans As Variant, cMeans As Variant, eMeans As Variant
    Dim sMain As Double, pMain As Double, cMain As Double, eMain As Double, aMean As Double, zaMean As Double
    Dim x As Variant, y As Variant, rr As Long, ratios As Variant, collision As String, idxScore As Double
    mStage = "Group scan - " & filterName
    Set resp = CreateObject("Scripting.Dictionary")
    For i = 1 To lng.Count
        If GroupMatch(typ, indexCode, crossValue, lng, i) Then
            If Not resp.Exists(CStr(lng.Respondent(i))) Then resp.Add CStr(lng.Respondent(i)), True
            If lng.NS(i) <> MISS And lng.NP(i) <> MISS And lng.NC(i) <> MISS And lng.NE(i) <> MISS And lng.ZA(i) <> MISS Then regN = regN + 1
        End If
    Next i
    If regN > 0 Then x = BSMat(regN, 4): y = BSMat(regN, 1)
    For i = 1 To lng.Count
        If GroupMatch(typ, indexCode, crossValue, lng, i) And lng.NS(i) <> MISS And lng.NP(i) <> MISS And lng.NC(i) <> MISS And lng.NE(i) <> MISS And lng.ZA(i) <> MISS Then
            rr = rr + 1: x(rr, 1) = lng.NS(i): x(rr, 2) = lng.NP(i): x(rr, 3) = lng.NC(i): x(rr, 4) = lng.NE(i): y(rr, 1) = lng.ZA(i)
        End If
    Next i
    mStage = "Factor/Regression - " & filterName
    On Error Resume Next
    If regN >= 4 Then
        ratios = BSFactorRatios(x, y, safeMode, collision)
    Else
        ReDim ratios(1 To 4)
    End If
    If Err.Number <> 0 Then Err.Clear: ReDim ratios(1 To 4)
    On Error GoTo 0
    mStage = "Category means - " & filterName
    sMeans = CategoryMeansSelected(lng.SVal, sGroups, typ, indexCode, crossValue, lng)
    pMeans = CategoryMeansSelected(lng.PVal, pGroups, typ, indexCode, crossValue, lng)
    cMeans = CategoryMeansSelected(lng.CVal, cGroups, typ, indexCode, crossValue, lng)
    eMeans = CategoryMeansSelected(lng.EVal, eGroups, typ, indexCode, crossValue, lng)
    sMain = MeanArray(sMeans) * 100#: pMain = MeanArray(pMeans) * 100#: cMain = MeanArray(cMeans) * 100#: eMain = MeanArray(eMeans) * 100#
    aMean = ScalarMean(lng.A, typ, indexCode, crossValue, lng): zaMean = ScalarMean(lng.ZA, typ, indexCode, crossValue, lng) * 100#
    idxScore = (sMain * ratios(1) + pMain * ratios(2) + cMain * ratios(3) + eMain * ratios(4)) / 100#
    label = GroupDisplayLabel(typ, indexCode, crossValue, st)
    ws.Cells(outRow, 1).Resize(1, 15).Value = Array(IIf(typ = 1 Or typ = 3, indexCode, 0), label, resp.Count & " / Reg=" & regN, filterName, sMain, pMain, aMean, zaMean, idxScore, cMain, eMain, ratios(1), ratios(2), ratios(3), ratios(4))
    col = 15
    WriteArrayRow ws, outRow, col, sMeans, 100#: WriteArrayRow ws, outRow, col, pMeans, 100#
    WriteArrayRow ws, outRow, col, cMeans, 100#: WriteArrayRow ws, outRow, col, eMeans, 100#
    mStage = "Correlations - " & filterName
    WriteCorrelationSelected ws, outRow, col, lng.EVal, eGroups, lng.A, typ, indexCode, crossValue, lng
    WriteCorrelationSelected ws, outRow, col, lng.SVal, sGroups, lng.A, typ, indexCode, crossValue, lng
    WriteCorrelationSelected ws, outRow, col, lng.PVal, pGroups, lng.A, typ, indexCode, crossValue, lng
    mStage = "Agree/T2B - " & filterName
    WriteAgreeRow ws, outRow, col, st.AgreeS, GroupArrayCount(sGroups), raw, headers, resp, st.T2BChoice
    WriteAgreeRow ws, outRow, col, st.AgreeP, GroupArrayCount(pGroups), raw, headers, resp, st.T2BChoice
    If safeMode And regN > 0 And regN < LOW_MODEL_N Then ws.Cells(outRow, 3).Interior.Color = RGB(255, 235, 156)
End Sub

Private Function CategoryMeansSelected(ByRef values() As Double, ByVal groups As Variant, ByVal typ As Long, ByVal indexCode As Long, ByVal crossValue As Variant, ByRef lng As BSLong) As Variant
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

Private Function MeanArray(ByVal a As Variant) As Double
    Dim i As Long, s As Double, n As Long
    If UBound(a) = 0 Then Exit Function
    For i = LBound(a) To UBound(a): If a(i) <> MISS Then s = s + a(i): n = n + 1
    Next i
    If n > 0 Then MeanArray = s / n
End Function

Private Function ScalarMean(ByRef values() As Double, ByVal typ As Long, ByVal indexCode As Long, ByVal crossValue As Variant, ByRef lng As BSLong) As Double
    Dim i As Long, s As Double, n As Long
    For i = 1 To lng.Count
        If GroupMatch(typ, indexCode, crossValue, lng, i) And values(i) <> MISS Then s = s + values(i): n = n + 1
    Next i
    If n > 0 Then ScalarMean = s / n
End Function

Private Sub WriteArrayRow(ByVal ws As Worksheet, ByVal r As Long, ByRef col As Long, ByVal a As Variant, ByVal multiplier As Double)
    Dim i As Long: If UBound(a) = 0 Then Exit Sub
    For i = LBound(a) To UBound(a): col = col + 1: If a(i) <> MISS Then ws.Cells(r, col).Value = a(i) * multiplier
    Next i
End Sub

Private Sub WriteCorrelationSelected(ByVal ws As Worksheet, ByVal outRow As Long, ByRef col As Long, ByRef values() As Double, ByVal groups As Variant, ByRef target() As Double, ByVal typ As Long, ByVal indexCode As Long, ByVal crossValue As Variant, ByRef lng As BSLong)
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

Private Sub WriteAgreeRow(ByVal ws As Worksheet, ByVal outRow As Long, ByRef col As Long, ByRef vars() As String, ByVal groups As Long, ByVal raw As Variant, ByVal headers As Object, ByVal resp As Object, ByVal t2bChoice As String)
    Dim g As Long, key As Variant, countGood As Long, total As Long, v As Double, rawRow As Long
    For g = 1 To groups
        col = col + 1: countGood = 0: total = resp.Count
        If g >= LBound(vars) And g <= UBound(vars) Then
            If Len(vars(g)) > 0 And headers.Exists(vars(g)) And total > 0 Then
            For Each key In resp.Keys
                rawRow = CLng(key) + 1: v = NumericOrMiss(raw(rawRow, headers(vars(g))))
                If v <> MISS Then
                    If t2bChoice = "5+4" Then
                        If v = 4 Or v = 5 Then countGood = countGood + 1
                    Else
                        If v = 1 Or v = 2 Then countGood = countGood + 1
                    End If
                End If
            Next key
            ws.Cells(outRow, col).Value = countGood / total * 100#
            End If
        End If
    Next g
End Sub

Private Function GroupDisplayLabel(ByVal typ As Long, ByVal indexCode As Long, ByVal crossValue As Variant, ByRef st As BSSetting) As String
    Dim i As Long, idxLabel As String, crossLabelText As String
    For i = 1 To UBound(st.IndexCodes): If st.IndexCodes(i) = indexCode Then idxLabel = st.IndexLabels(i): Exit For
    Next i
    crossLabelText = CrossLabel(st, crossValue)
    Select Case typ
        Case 0: GroupDisplayLabel = "Overall"
        Case 1: GroupDisplayLabel = idxLabel
        Case 2: GroupDisplayLabel = crossLabelText
        Case 3: GroupDisplayLabel = idxLabel & " - " & crossLabelText
    End Select
End Function

Private Sub StyleSummary(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    StyleTable ws, lastRow, lastCol: ws.Rows(1).RowHeight = 34
    ws.Columns(1).ColumnWidth = 12: ws.Columns(2).ColumnWidth = 34: ws.Columns(3).ColumnWidth = 17: ws.Columns(4).ColumnWidth = 28
    ws.Range(ws.Cells(2, 5), ws.Cells(lastRow, lastCol)).NumberFormat = "0.00"
    Dim c As Long, h As String
    For c = 1 To lastCol
        h = CStr(ws.Cells(1, c).Value2)
        If Left$(h, 4) = "cor_" Or Left$(h, 5) = "CorE_" Then ws.Range(ws.Cells(2, c), ws.Cells(lastRow, c)).NumberFormat = "0.000"
        If Left$(h, 6) = "agree_" Then ws.Range(ws.Cells(2, c), ws.Cells(lastRow, c)).NumberFormat = "0.0"
    Next c
    AddColorScale ws, "Index", lastRow: AddColorScale ws, "B.S", lastRow: AddColorScale ws, "B.P", lastRow: AddColorScale ws, "B.C", lastRow: AddColorScale ws, "B.E", lastRow
    ws.Activate: ActiveWindow.SplitRow = 1: ActiveWindow.SplitColumn = 1: ActiveWindow.FreezePanes = True
    ws.Tab.Color = RGB(0, 176, 240)
End Sub

Private Sub AddColorScale(ByVal ws As Worksheet, ByVal header As String, ByVal lastRow As Long)
    Dim f As Range, rng As Range
    Set f = ws.Rows(1).Find(What:=header, LookAt:=xlWhole)
    If f Is Nothing Or lastRow < 2 Then Exit Sub
    Set rng = ws.Range(ws.Cells(2, f.Column), ws.Cells(lastRow, f.Column)): rng.FormatConditions.Delete
    rng.FormatConditions.AddColorScale ColorScaleType:=3
    With rng.FormatConditions(1)
        .ColorScaleCriteria(1).FormatColor.Color = RGB(248, 105, 107)
        .ColorScaleCriteria(2).Type = xlConditionValuePercentile: .ColorScaleCriteria(2).Value = 50: .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 235, 132)
        .ColorScaleCriteria(3).FormatColor.Color = RGB(99, 190, 123)
    End With
End Sub

Private Sub WriteSandP(ByRef st As BSSetting)
    Dim ws As Worksheet, sN As Long, pN As Long, r As Long, i As Long
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("Summary")): ws.Name = "SandP"
    ws.Cells(1, 1).Resize(1, 6).Value = Array("Variable", "DescriptionTH", "DescriptionEN", "Rank_list", "Spcode", "Important")
    sN = MaxGroup(st.S): pN = MaxGroup(st.P): r = 1
    For i = 1 To sN: r = r + 1: ws.Cells(r, 1).Value = "S_" & i: ws.Cells(r, 3).Value = "S_" & i: Next i
    For i = 1 To pN: r = r + 1: ws.Cells(r, 1).Value = "P_" & i: ws.Cells(r, 3).Value = "P_" & i: Next i
    StyleTable ws, r, 6: ws.Columns("A:F").AutoFit: ws.Columns("B:C").ColumnWidth = 34: ws.Tab.Color = RGB(0, 176, 240)
End Sub

Private Sub WriteCorrespondence(ByRef st As BSSetting, ByRef lng As BSLong, ByVal sheetName As String, ByVal prefix As String, ByVal groups As Variant)
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

Private Sub StyleTable(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Interior.Color = RGB(79, 129, 189): .Font.Color = vbWhite: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True
    End With
    With ws.Range(ws.Cells(1, 1), ws.Cells(IIf(lastRow > 1, lastRow, 1), lastCol)).Borders
        .LineStyle = xlContinuous: .Color = RGB(210, 210, 210): .Weight = xlThin
    End With
    ws.Cells.Font.Name = "Calibri": ws.Cells.Font.Size = 10: ws.Rows(1).AutoFilter
End Sub

Private Sub ReorderSheets()
    ThisWorkbook.Worksheets("Control").Move Before:=ThisWorkbook.Worksheets(1)
    ThisWorkbook.Worksheets("Rawdata").Move After:=ThisWorkbook.Worksheets("Control")
    ThisWorkbook.Worksheets("Setting Variables").Move After:=ThisWorkbook.Worksheets("Rawdata")
    ThisWorkbook.Worksheets("Setting Labels").Move After:=ThisWorkbook.Worksheets("Setting Variables")
    ThisWorkbook.Worksheets("Summary").Move After:=ThisWorkbook.Worksheets("Setting Labels")
    ThisWorkbook.Worksheets("SandP").Move After:=ThisWorkbook.Worksheets("Summary")
    ThisWorkbook.Worksheets("Correspondence(S)").Move After:=ThisWorkbook.Worksheets("SandP")
    ThisWorkbook.Worksheets("Correspondence(P)").Move After:=ThisWorkbook.Worksheets("Correspondence(S)")
    ThisWorkbook.Worksheets("QC Excluded").Move After:=ThisWorkbook.Worksheets("Correspondence(P)")
    ThisWorkbook.Worksheets("Control").Activate
End Sub

Private Sub DeleteSheetIfExists(ByVal sheetName As String)
    Dim ws As Worksheet: On Error Resume Next: Set ws = ThisWorkbook.Worksheets(sheetName): On Error GoTo 0
    If Not ws Is Nothing Then ws.Delete
End Sub

Private Sub UpdateStatus(ByVal textValue As String, ByVal colorValue As Long)
    With ThisWorkbook.Worksheets(CONTROL_SHEET).Range("B10:G10")
        .Cells(1, 1).Value = textValue: .Interior.Color = colorValue: .Font.Bold = True: .WrapText = True
    End With
    DoEvents
End Sub
