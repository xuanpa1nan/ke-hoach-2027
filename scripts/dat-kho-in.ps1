param(
    [Parameter(Mandatory, Position = 0, ValueFromRemainingArguments)]
    [string[]]$WorkbookPath
)

$ErrorActionPreference = 'Stop'
$workbooks = foreach ($path in $WorkbookPath) {
    $item = Get-Item -LiteralPath $path
    if ($item.PSIsContainer -or $item.Extension -ne '.xlsx') {
        throw "Cần đường dẫn tệp .xlsx: $path"
    }
    $item.FullName
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    foreach ($path in $workbooks) {
        $book = $excel.Workbooks.Open($path, 0, $false)
        try {
            if ($book.ReadOnly) { throw "Workbook đang mở hoặc chỉ đọc: $path" }
            foreach ($sheet in $book.Worksheets) {
                if ($sheet.Name -notlike 'In *') { continue }
                $setup = $sheet.PageSetup
                $setup.PaperSize = 9  # A4
                $setup.Orientation = 2  # Landscape
                $setup.Zoom = $false
                $setup.FitToPagesWide = 1
                $setup.FitToPagesTall = 1
                $setup.PrintArea = $sheet.UsedRange.Address()
                $setup.PrintTitleRows = '$1:$2'
                $setup.LeftMargin = $excel.InchesToPoints(0.28)
                $setup.RightMargin = $excel.InchesToPoints(0.28)
                $setup.TopMargin = $excel.InchesToPoints(0.33)
                $setup.BottomMargin = $excel.InchesToPoints(0.33)
                $setup.CenterHorizontally = $true
                $setup.PrintGridlines = $false
            }
            $book.Save()
        } finally {
            $book.Close($false)
        }
    }
} finally {
    $excel.Quit()
}
