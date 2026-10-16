param(
    [Parameter(Mandatory=$true)][string]$InputDocx,
    [Parameter(Mandatory=$true)][string]$OutputPdf
)
$ErrorActionPreference = 'Stop'
$sourcePath = (Resolve-Path -LiteralPath $InputDocx).Path
$destinationPath = [IO.Path]::GetFullPath($OutputPdf)
if ([IO.Path]::GetExtension($destinationPath) -ne '.pdf') { throw 'La salida debe ser PDF.' }
if (Test-Path -LiteralPath $destinationPath) { throw 'La salida ya existe.' }
$destinationDirectory = [IO.Path]::GetDirectoryName($destinationPath)
[IO.Directory]::CreateDirectory($destinationDirectory) | Out-Null
$sourceHash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
$wordApp = $null
$wordDocument = $null
try {
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $false
    $wordApp.DisplayAlerts = 0
    $wordApp.AutomationSecurity = 3
    $wordDocument = $wordApp.Documents.Open($sourcePath, $false, $true, $false)
    $wordDocument.Repaginate()
    $pages = $wordDocument.ComputeStatistics(2)
    $controls = $wordDocument.ContentControls.Count
    $wordDocument.ExportAsFixedFormat($destinationPath, 17, $false)
    [PSCustomObject]@{Input=$sourcePath; PDF=$destinationPath; Pages=$pages; ContentControls=$controls} | ConvertTo-Json
}
finally {
    $noSave = 0
    if ($null -ne $wordDocument) {
        $wordDocument.Close([ref]$noSave)
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordDocument) | Out-Null
    }
    if ($null -ne $wordApp) {
        $wordApp.Quit([ref]$noSave)
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordApp) | Out-Null
    }
}
if ((Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash -ne $sourceHash) { throw 'El archivo de entrada cambió durante la exportación.' }
if (-not (Test-Path -LiteralPath $destinationPath)) { throw 'Word no produjo el PDF.' }
