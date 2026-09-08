Attribute VB_Name = "BrandSenseRunner"
Option Explicit

Private Const CONTROL_SHEET As String = "Control"
Private Const RAW_SHEET As String = "Rawdata"
Private Const ENGINE_SHEET As String = "__Engine"

Public Sub RunBrandSenseSafe()
    RunBrandSenseEngine "safe_all"
End Sub

Public Sub RunBrandSenseNormal()
    RunBrandSenseEngine "normal"
End Sub

Private Sub RunBrandSenseEngine(ByVal runMode As String)
    Dim workFolder As String
    Dim rawCsv As String
    Dim outputPath As String
    Dim commandLine As String
    Dim exitCode As Long
    Dim shell As Object
    Dim priorCalc As XlCalculation

    On Error GoTo Failed
    priorCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    If Not RawdataIsReady() Then
        Err.Raise vbObjectError + 101, , _
            "Rawdata must have a header row and at least one data row."
    End If
    ThisWorkbook.Save

    workFolder = Environ$("TEMP") & "\BrandSenseExcel_" & _
        Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Rnd() * 100000))
    MkDir workFolder

    UpdateStatus "Preparing embedded engine...", RGB(255, 242, 204)
    ExtractAsset "program.py", workFolder & "\program.py"
    ExtractAsset "runner.py", workFolder & "\runner.py"
    ExtractAsset "metadata.sav", workFolder & "\metadata.sav"

    rawCsv = workFolder & "\rawdata.csv"
    ExportRawdataCsv rawCsv

    UpdateStatus "Running BrandSense model...", RGB(221, 235, 247)
    commandLine = QuoteText("python") & " " & _
        QuoteText(workFolder & "\runner.py") & " " & _
        QuoteText(workFolder & "\program.py") & " " & _
        QuoteText(ThisWorkbook.FullName) & " " & _
        QuoteText(workFolder & "\metadata.sav") & " " & _
        QuoteText(rawCsv) & " " & QuoteText(runMode) & " " & _
        QuoteText(workFolder)

    Set shell = CreateObject("WScript.Shell")
    exitCode = shell.Run(commandLine, 0, True)
    If exitCode <> 0 Then
        Err.Raise vbObjectError + 102, , _
            "BrandSense engine failed:" & vbCrLf & _
            ReadTextFile(workFolder & "\error.txt")
    End If

    outputPath = workFolder & "\BrandSense_Engine_Input BS Output.xlsx"
    If Dir$(outputPath) = vbNullString Then
        Err.Raise vbObjectError + 103, , _
            "The engine finished but no output workbook was found."
    End If

    UpdateStatus "Importing results...", RGB(226, 239, 218)
    ImportOutputSheets outputPath
    ThisWorkbook.Worksheets(CONTROL_SHEET).Range("B11").Value = _
        IIf(runMode = "safe_all", _
            "QC All + Safe Mapping", "Normal / Legacy")
    ThisWorkbook.Worksheets(CONTROL_SHEET).Range("B12").Value = Now
    ThisWorkbook.Worksheets(CONTROL_SHEET).Range("B12").NumberFormat = _
        "yyyy-mm-dd hh:mm:ss"
    ThisWorkbook.Worksheets(CONTROL_SHEET).Range("B13").Value = _
        JsonStringValue(ReadTextFile(workFolder & "\result.json"), "setting")
    UpdateStatus "Completed successfully", RGB(198, 239, 206)
    ThisWorkbook.Save

CleanExit:
    Application.Calculation = priorCalc
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

Failed:
    UpdateStatus "Failed: " & Err.Description, RGB(255, 199, 206)
    If Application.Visible Then
        MsgBox Err.Description, vbCritical, "BrandSense Excel Runner"
    End If
    Resume CleanExit
End Sub

Private Function RawdataIsReady() As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Set ws = ThisWorkbook.Worksheets(RAW_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    RawdataIsReady = (lastRow >= 2 And lastCol >= 2)
End Function

Private Sub ExportRawdataCsv(ByVal csvPath As String)
    Dim sourceSheet As Worksheet
    Dim temporaryBook As Workbook
    Set sourceSheet = ThisWorkbook.Worksheets(RAW_SHEET)
    sourceSheet.Copy
    Set temporaryBook = ActiveWorkbook
    temporaryBook.SaveAs Filename:=csvPath, FileFormat:=62, _
        CreateBackup:=False, Local:=True
    temporaryBook.Close SaveChanges:=False
End Sub

Private Sub ExtractAsset(ByVal assetName As String, ByVal outputPath As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNumber As Long
    Dim encoded As String
    Dim xmlDocument As Object
    Dim xmlNode As Object
    Dim stream As Object

    Set ws = ThisWorkbook.Worksheets(ENGINE_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowNumber = 2 To lastRow
        If CStr(ws.Cells(rowNumber, 1).Value2) = assetName Then
            encoded = encoded & CStr(ws.Cells(rowNumber, 3).Value2)
        End If
    Next rowNumber
    If Len(encoded) = 0 Then
        Err.Raise vbObjectError + 104, , _
            "Embedded asset is missing: " & assetName
    End If

    Set xmlDocument = CreateObject("Msxml2.DOMDocument.6.0")
    Set xmlNode = xmlDocument.createElement("base64")
    xmlNode.DataType = "bin.base64"
    xmlNode.Text = encoded

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write xmlNode.nodeTypedValue
    stream.SaveToFile outputPath, 2
    stream.Close
End Sub

Private Sub ImportOutputSheets(ByVal outputPath As String)
    Dim sourceBook As Workbook
    Dim sheetNames As Variant
    Dim sheetName As Variant

    sheetNames = Array( _
        "Summary", "SandP", "Correspondence(S)", _
        "Correspondence(P)", "QC Excluded")

    For Each sheetName In sheetNames
        DeleteSheetIfExists CStr(sheetName)
    Next sheetName

    Set sourceBook = Application.Workbooks.Open( _
        Filename:=outputPath, ReadOnly:=True)
    For Each sheetName In sheetNames
        If SheetExists(sourceBook, CStr(sheetName)) Then
            sourceBook.Worksheets(CStr(sheetName)).Copy _
                Before:=ThisWorkbook.Worksheets(RAW_SHEET)
        End If
    Next sheetName
    sourceBook.Close SaveChanges:=False
    ThisWorkbook.Worksheets(CONTROL_SHEET).Move _
        Before:=ThisWorkbook.Worksheets(1)
    ThisWorkbook.Worksheets(RAW_SHEET).Move _
        After:=ThisWorkbook.Worksheets(CONTROL_SHEET)
    ThisWorkbook.Worksheets("Setting Variables").Move _
        After:=ThisWorkbook.Worksheets(RAW_SHEET)
    ThisWorkbook.Worksheets("Setting Labels").Move _
        After:=ThisWorkbook.Worksheets("Setting Variables")
    ThisWorkbook.Worksheets(CONTROL_SHEET).Activate
End Sub

Private Sub DeleteSheetIfExists(ByVal sheetName As String)
    If SheetExists(ThisWorkbook, sheetName) Then
        ThisWorkbook.Worksheets(sheetName).Delete
    End If
End Sub

Private Function SheetExists(ByVal workbookObject As Workbook, _
                             ByVal sheetName As String) As Boolean
    Dim worksheetObject As Worksheet
    On Error Resume Next
    Set worksheetObject = workbookObject.Worksheets(sheetName)
    SheetExists = Not worksheetObject Is Nothing
    Set worksheetObject = Nothing
    On Error GoTo 0
End Function

Private Sub UpdateStatus(ByVal statusText As String, ByVal fillColor As Long)
    With ThisWorkbook.Worksheets(CONTROL_SHEET).Range("B10")
        .Value = statusText
        .Interior.Color = fillColor
        .Font.Bold = True
        .WrapText = True
    End With
    DoEvents
End Sub

Private Function QuoteText(ByVal value As String) As String
    QuoteText = Chr$(34) & value & Chr$(34)
End Function

Private Function JsonStringValue(ByVal jsonText As String, _
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
    Dim stream As Object
    If Dir$(filePath) = vbNullString Then
        ReadTextFile = "No error detail was produced."
        Exit Function
    End If
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadTextFile = stream.ReadText
    stream.Close
End Function

