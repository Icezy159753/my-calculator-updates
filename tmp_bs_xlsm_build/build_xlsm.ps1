$ErrorActionPreference = 'Stop'

function OleColor([int]$r, [int]$g, [int]$b) {
    return $r + (256 * $g) + (65536 * $b)
}

function Release-ComObject($object) {
    if ($null -ne $object) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($object)
    }
}

$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$outputDirectory = Join-Path $root 'outputs\brandsense_rawdata_xlsm'
[System.IO.Directory]::CreateDirectory($outputDirectory) | Out-Null
$outputPath = Join-Path $outputDirectory 'BrandSense_Rawdata_Auto_Set_1_4_5_6.xlsm'

$programPath = Join-Path $root 'All_Programs\123_Program_Run_Brandsence2026.py'
$runnerPath = Join-Path $PSScriptRoot 'runner.py'
$modulePath = Join-Path $PSScriptRoot 'BrandSenseRunner.bas'
$metadataPath = Join-Path $root 'data\BS\SPSS_preserved_utf8_Final.sav'
$rawCsvPath = Join-Path $PSScriptRoot 'test_runtime\rawdata.csv'

$assets = [ordered]@{
    'program.py' = $programPath
    'runner.py' = $runnerPath
    'metadata.sav' = $metadataPath
    'setting_1.xlsx' = (Join-Path $root 'data\BS\1_Setting BS Set1 Bangkok Hospital Headquarters.xlsx')
    'setting_4.xlsx' = (Join-Path $root 'data\BS\4_Setting BS Set4 Bangkok Hospital Pattaya.xlsx')
    'setting_5.xlsx' = (Join-Path $root 'data\BS\5_Setting BS Set5 Bangkok Hospital Chanthaburi.xlsx')
    'setting_6.xlsx' = (Join-Path $root 'data\BS\6_Setting BS Set6 Bangkok Hospital Rayong.xlsx')
}

foreach ($path in @($programPath, $runnerPath, $modulePath, $metadataPath, $rawCsvPath) + @($assets.Values)) {
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Required build input is missing: $path"
    }
}

$excel = $null
$workbook = $null
$csvBook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false

    $workbook = $excel.Workbooks.Add()
    while ($workbook.Worksheets.Count -gt 1) {
        $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
    }

    $control = $workbook.Worksheets.Item(1)
    $control.Name = 'Control'
    $raw = $workbook.Worksheets.Add([System.Type]::Missing, $control)
    $raw.Name = 'Rawdata'
    $engine = $workbook.Worksheets.Add([System.Type]::Missing, $raw)
    $engine.Name = '__Engine'

    $control.Move($workbook.Worksheets.Item(1))
    $raw.Move([System.Type]::Missing, $control)

    $red = OleColor 198 40 45
    $darkRed = OleColor 139 24 28
    $blue = OleColor 79 129 189
    $lightBlue = OleColor 221 235 247
    $lightYellow = OleColor 255 242 204
    $lightGreen = OleColor 226 239 218
    $darkText = OleColor 48 48 48
    $white = OleColor 255 255 255
    $softGray = OleColor 242 242 242

    $control.Cells.Font.Name = 'Leelawadee UI'
    $control.Cells.Font.Size = 11
    $control.Columns('A').ColumnWidth = 4
    $control.Columns('B:G').ColumnWidth = 16
    $control.Columns('H').ColumnWidth = 4
    $control.Rows('1:30').RowHeight = 24
    $control.Rows('1:2').RowHeight = 34
    $control.Range('A1:H2').Merge()
    $control.Range('A1').Value2 = 'BrandSense — Rawdata Runner'
    $control.Range('A1').Interior.Color = $red
    $control.Range('A1').Font.Color = $white
    $control.Range('A1').Font.Bold = $true
    $control.Range('A1').Font.Size = 22
    $control.Range('A1').HorizontalAlignment = -4108
    $control.Range('A1').VerticalAlignment = -4108

    $control.Range('B4:G4').Merge()
    $control.Range('B4').Value2 = 'วาง Rawdata แล้วกดรัน — ระบบเลือก Setting 1 / 4 / 5 / 6 ให้อัตโนมัติ'
    $control.Range('B4').Font.Bold = $true
    $control.Range('B4').Font.Size = 13
    $control.Range('B4').Font.Color = $darkRed
    $control.Range('B4').HorizontalAlignment = -4108

    $instructions = @(
        'ไปที่ชีท Rawdata แล้วลบข้อมูลเดิมทั้งหมด',
        'วางหัวคอลัมน์และข้อมูล ตั้งแต่เซลล์ A1 (ห้ามเปลี่ยนชื่อตัวแปร)',
        'กลับมาชีท Control แล้วเลือกปุ่มรันด้านล่าง'
    )
    for ($i = 0; $i -lt $instructions.Count; $i++) {
        $row = 6 + $i
        $control.Range("A$row").Value2 = [string]($i + 1)
        $control.Range("A$row").Interior.Color = $blue
        $control.Range("A$row").Font.Color = $white
        $control.Range("A$row").Font.Bold = $true
        $control.Range("A$row").HorizontalAlignment = -4108
        $control.Range("B$row:G$row").Merge()
        $control.Range("B$row").Value2 = $instructions[$i]
        $control.Range("B$row").Interior.Color = $lightBlue
    }

    $control.Range('A10').Value2 = 'สถานะ'
    $control.Range('A10').Font.Bold = $true
    $control.Range('B10:G10').Merge()
    $control.Range('B10').Value2 = 'พร้อมรัน'
    $control.Range('B10').Interior.Color = $lightGreen
    $control.Range('A11').Value2 = 'โหมดล่าสุด'
    $control.Range('B11:G11').Merge()
    $control.Range('B11').Value2 = '-'
    $control.Range('A12').Value2 = 'เวลาล่าสุด'
    $control.Range('B12:G12').Merge()
    $control.Range('B12').Value2 = '-'
    $control.Range('A13').Value2 = 'Setting'
    $control.Range('B13:G13').Merge()
    $control.Range('B13').Value2 = 'ตรวจจับจากชื่อคอลัมน์เมื่อกดรัน'
    $control.Range('A10:A13').Font.Bold = $true
    $control.Range('A10:G13').Borders.LineStyle = 1
    $control.Range('A10:G13').Borders.Color = OleColor 217 217 217

    $safeButton = $control.Shapes.AddShape(5, 62, 345, 655, 62)
    $safeButton.Name = 'btnRunSafe'
    $safeButton.TextFrame2.TextRange.Text = 'รัน: ตัดเคส QC ทั้งหมด + Safe Mapping'
    $safeButton.TextFrame2.TextRange.Font.Name = 'Leelawadee UI'
    $safeButton.TextFrame2.TextRange.Font.Size = 15
    $safeButton.TextFrame2.TextRange.Font.Bold = -1
    $safeButton.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $white
    $safeButton.Fill.ForeColor.RGB = $red
    $safeButton.Line.ForeColor.RGB = $red
    $safeButton.OnAction = 'RunBrandSenseSafe'

    $normalButton = $control.Shapes.AddShape(5, 62, 420, 655, 46)
    $normalButton.Name = 'btnRunNormal'
    $normalButton.TextFrame2.TextRange.Text = 'รันปกติ (ไม่ตัด QC / Legacy Mapping)'
    $normalButton.TextFrame2.TextRange.Font.Name = 'Leelawadee UI'
    $normalButton.TextFrame2.TextRange.Font.Size = 12
    $normalButton.TextFrame2.TextRange.Font.Bold = -1
    $normalButton.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $darkRed
    $normalButton.Fill.ForeColor.RGB = $white
    $normalButton.Line.ForeColor.RGB = $red
    $normalButton.Line.Weight = 1.5
    $normalButton.OnAction = 'RunBrandSenseNormal'

    $control.Range('B22:G25').Merge()
    $control.Range('B22').Value2 = "หมายเหตุ: ปุ่มแรกจะคัดเคสคำตอบคุณภาพต่ำออกจาก Long Format และใช้การจับ Factor แบบปลอดภัยตามโปรแกรมปัจจุบัน ส่วนปุ่มรันปกติจะไม่ตัดเคสและคง Legacy Mapping ไว้`nเครื่องนี้ต้องมี Python และแพ็กเกจเดียวกับโปรแกรม BrandSense"
    $control.Range('B22').WrapText = $true
    $control.Range('B22').Interior.Color = $lightYellow
    $control.Range('B22').Font.Color = $darkText
    $control.Range('B22').VerticalAlignment = -4108
    $control.Rows('22:25').RowHeight = 28

    $csvBook = $excel.Workbooks.Open($rawCsvPath)
    $csvSheet = $csvBook.Worksheets.Item(1)
    $used = $csvSheet.UsedRange
    $used.Copy($raw.Range('A1'))
    $csvBook.Close($false)
    $csvBook = $null
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($used)
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($csvSheet)

    $raw.Cells.Font.Name = 'Calibri'
    $raw.Cells.Font.Size = 10
    $lastCol = $raw.Cells.Item(1, $raw.Columns.Count).End(-4159).Column
    $lastRow = $raw.Cells.Item($raw.Rows.Count, 1).End(-4162).Row
    $header = $raw.Range($raw.Cells.Item(1, 1), $raw.Cells.Item(1, $lastCol))
    $header.Interior.Color = $blue
    $header.Font.Color = $white
    $header.Font.Bold = $true
    $header.WrapText = $true
    $header.HorizontalAlignment = -4108
    $header.AutoFilter() | Out-Null
    $raw.Rows.Item(1).RowHeight = 34
    $raw.Range($raw.Cells.Item(1, 1), $raw.Cells.Item($lastRow, $lastCol)).Borders.Color = OleColor 225 225 225
    $raw.Range($raw.Cells.Item(1, 1), $raw.Cells.Item($lastRow, $lastCol)).Borders.LineStyle = 1
    $raw.Columns.ColumnWidth = 12
    $raw.Columns.Item(1).ColumnWidth = 12
    $raw.Columns.Item(2).ColumnWidth = 12
    $raw.Activate()
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.FreezePanes = $true

    $engine.Cells.Item(1, 1).Value2 = 'Asset'
    $engine.Cells.Item(1, 2).Value2 = 'Sequence'
    $engine.Cells.Item(1, 3).Value2 = 'Base64'
    $engineRow = 2
    foreach ($entry in $assets.GetEnumerator()) {
        $bytes = [System.IO.File]::ReadAllBytes($entry.Value)
        $encoded = [Convert]::ToBase64String($bytes)
        $chunkSize = 24000
        $sequence = 1
        for ($offset = 0; $offset -lt $encoded.Length; $offset += $chunkSize) {
            $length = [Math]::Min($chunkSize, $encoded.Length - $offset)
            $engine.Cells.Item($engineRow, 1).Value2 = $entry.Key
            $engine.Cells.Item($engineRow, 2).Value2 = $sequence
            $engine.Cells.Item($engineRow, 3).Value2 = $encoded.Substring($offset, $length)
            $engineRow++
            $sequence++
        }
    }
    $engine.Visible = 2

    $component = $workbook.VBProject.VBComponents.Import($modulePath)
    Release-ComObject $component

    if (Test-Path -LiteralPath $outputPath) {
        Remove-Item -LiteralPath $outputPath -Force
    }
    $workbook.SaveAs($outputPath, 52)

    $workbook.Worksheets.Item('Control').Activate()
    $excel.ActiveWindow.DisplayGridlines = $false
    $excel.ActiveWindow.Zoom = 90
    $workbook.Save()
    Write-Output $outputPath
}
finally {
    if ($null -ne $csvBook) {
        try { $csvBook.Close($false) } catch {}
    }
    if ($null -ne $workbook) {
        try { $workbook.Close($true) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
    }
    Release-ComObject $csvBook
    Release-ComObject $workbook
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
