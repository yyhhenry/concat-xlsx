param (
    [Parameter(Mandatory = $true)]
    [string]$Path
)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$xlOpenXMLWorkbook = 51


if (-not (Test-Path $Path)) {
    Write-Error "The specified path does not exist: $Path"
    exit 1
}

Get-ChildItem -Path $Path | Where-Object { $_.Extension -eq ".xls" } | ForEach-Object {
    $workbook = $excel.Workbooks.Open($_.FullName)
    $newName = $_.FullName -replace '\.xls$', '.xlsx'
    Write-Output "Generate $newName"
    $workbook.SaveAs($newName, $xlOpenXMLWorkbook)
    $workbook.Close()
}

$excel.Quit()
$result = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
if ($result -ne 0) {
    Write-Output "Failed to release Excel Com Object"
}
Write-Output "Done"
