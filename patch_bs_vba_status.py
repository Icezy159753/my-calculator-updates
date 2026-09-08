from pathlib import Path


path = Path(__file__).resolve().parent / "tmp_bs_xlsm_build" / "BrandSenseRunner.bas"
text = path.read_text(encoding="utf-8-sig")

old = '''    ThisWorkbook.Worksheets(CONTROL_SHEET).Range("B10").Value = _
        IIf(runMode = "safe_all", _
            "QC All + Safe Mapping", "Normal / Legacy")
    ThisWorkbook.Worksheets(CONTROL_SHEET).Range("B11").Value = Now
    ThisWorkbook.Worksheets(CONTROL_SHEET).Range("B11").NumberFormat = _
        "yyyy-mm-dd hh:mm:ss"
    UpdateStatus "Completed successfully", RGB(198, 239, 206)'''
new = '''    ThisWorkbook.Worksheets(CONTROL_SHEET).Range("B11").Value = _
        IIf(runMode = "safe_all", _
            "QC All + Safe Mapping", "Normal / Legacy")
    ThisWorkbook.Worksheets(CONTROL_SHEET).Range("B12").Value = Now
    ThisWorkbook.Worksheets(CONTROL_SHEET).Range("B12").NumberFormat = _
        "yyyy-mm-dd hh:mm:ss"
    ThisWorkbook.Worksheets(CONTROL_SHEET).Range("B13").Value = _
        JsonStringValue(ReadTextFile(workFolder & "\\result.json"), "setting")
    UpdateStatus "Completed successfully", RGB(198, 239, 206)'''
assert text.count(old) == 1
text = text.replace(old, new, 1)

old = '    With ThisWorkbook.Worksheets(CONTROL_SHEET).Range("B8")'
new = '    With ThisWorkbook.Worksheets(CONTROL_SHEET).Range("B10")'
assert text.count(old) == 1
text = text.replace(old, new, 1)

marker = '''Private Function ReadTextFile(ByVal filePath As String) As String
'''
helper = '''Private Function JsonStringValue(ByVal jsonText As String, _
                                 ByVal keyName As String) As String
    Dim marker As String
    Dim startPosition As Long
    Dim endPosition As Long
    marker = Chr$(34) & keyName & Chr$(34) & ":"
    startPosition = InStr(1, jsonText, marker, vbTextCompare)
    If startPosition = 0 Then
        JsonStringValue = "-"
        Exit Function
    End If
    startPosition = InStr(startPosition + Len(marker), jsonText, Chr$(34)) + 1
    endPosition = InStr(startPosition, jsonText, Chr$(34))
    If startPosition <= 1 Or endPosition <= startPosition Then
        JsonStringValue = "-"
    Else
        JsonStringValue = Mid$(jsonText, startPosition, endPosition - startPosition)
    End If
End Function

Private Function ReadTextFile(ByVal filePath As String) As String
'''
assert text.count(marker) == 1
text = text.replace(marker, helper, 1)
path.write_text(text, encoding="utf-8-sig")
print("patched VBA status cells and setting display")
