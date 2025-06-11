# For Word
$word = New-Object -ComObject Word.Application
$word.Visible = $false

$folderPath = "C:\Path\To\Your\WordFiles"  # ← Change this
Get-ChildItem -Path $folderPath -Filter *.docx | ForEach-Object {
    $doc = $word.Documents.Open($_.FullName)
    $doc.PrintOut() # For multiple copies use $doc.PrintOut(Copies=3)
    $doc.Close($false)
}

$word.Quit()

# For excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$folderPath = "C:\Path\To\Your\ExcelFiles"  # ← Change this path
Get-ChildItem -Path $folderPath -Filter *.xlsx | ForEach-Object {
    $workbook = $excel.Workbooks.Open($_.FullName)
    $workbook.PrintOut() # For multiple copies use $workbook.PrintOut(Copies=3)
    $workbook.Close($false)
}

$excel.Quit()

