from pathlib import Path

p = Path("tmp_bs_xlsm_build/BrandSenseVBA.bas")
s = p.read_text(encoding="utf-8")
s = s.replace(
    'Dim v As Double, sum As Double, cnt As Long, attrCount As Long, pCount As Long, attrTotal As Long, pTotal As Long',
    'Dim v As Double, aQC As Double, sum As Double, cnt As Long, attrCount As Long, pCount As Long, attrTotal As Long, pTotal As Long',
)
s = s.replace(
    'rowL = rowL + 1: lng.Respondent(rowL) = r - 1: lng.IndexCode(rowL) = st.IndexCodes(i)',
    'rowL = rowL + 1: aQC = MISS: lng.Respondent(rowL) = r - 1: lng.IndexCode(rowL) = st.IndexCodes(i)',
)
s = s.replace(
    '''lng.A(rowL) = NumericOrMiss(raw(r, mapA(i)))
                If lng.A(rowL) = 9 Then lng.A(rowL) = 0''',
    '''aQC = NumericOrMiss(raw(r, mapA(i)))
                lng.A(rowL) = aQC
                If lng.A(rowL) = 9 Then lng.A(rowL) = 0''',
)
s = s.replace(
    '''If lng.A(rowL) <> MISS Then
                    hard = ((lng.A(rowL) = 1 And attrCount > 0) Or (lng.A(rowL) = 2 And attrCount >= Application.WorksheetFunction.RoundUp(attrTotal * 0.85, 0)) Or (lng.A(rowL) >= 6 And lng.A(rowL) <= 8 And attrCount = 0))
                    review = ((lng.A(rowL) <= 3 And pCount >= Application.WorksheetFunction.RoundUp(pTotal * 0.85, 0)) Or (lng.A(rowL) = 3 And attrCount >= Application.WorksheetFunction.RoundUp(attrTotal * 0.9, 0)))
                End If
                If hard Or review Then
                    reason = QCReason(lng.A(rowL), attrCount, attrTotal, pCount, pTotal)''',
    '''If aQC <> MISS Then
                    hard = ((aQC = 1 And attrCount > 0) Or (aQC = 2 And attrCount >= Application.WorksheetFunction.RoundUp(attrTotal * 0.85, 0)) Or (aQC >= 6 And aQC <= 8 And attrCount = 0))
                    review = ((aQC <= 3 And pCount >= Application.WorksheetFunction.RoundUp(pTotal * 0.85, 0)) Or (aQC = 3 And attrCount >= Application.WorksheetFunction.RoundUp(attrTotal * 0.9, 0)))
                End If
                If hard Or review Then
                    reason = QCReason(aQC, attrCount, attrTotal, pCount, pTotal)''',
)
p.write_text(s, encoding="utf-8")
