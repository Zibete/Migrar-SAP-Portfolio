$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$repoRoot = $repoRoot.Path

$harnessDir = Join-Path $repoRoot "portfolio"
$harnessPath = Join-Path $harnessDir "CoreTestsHarness.xlsm"
$coreDir = Join-Path $repoRoot "src\\vba\\core"
$testsDir = Join-Path $repoRoot "src\\vba\\tests"

New-Item -ItemType Directory -Force $harnessDir | Out-Null

if (Test-Path $harnessPath) {
    Remove-Item -Force $harnessPath
}

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()

    try {
        $null = $workbook.VBProject
    } catch {
        throw "VBProject access is not trusted. Enable 'Trust access to the VBA project object model' in Excel and see docs/TESTING.md."
    }

    $modules = @()
    $modules += Get-ChildItem -Path $coreDir -Filter *.bas | Sort-Object Name
    $modules += Get-ChildItem -Path $testsDir -Filter *.bas | Sort-Object Name

    foreach ($module in $modules) {
        $workbook.VBProject.VBComponents.Import($module.FullName) | Out-Null
    }

    $xlOpenXMLWorkbookMacroEnabled = 52
    $workbook.SaveAs($harnessPath, $xlOpenXMLWorkbookMacroEnabled)
    $workbook.Close($true)
} catch {
    Write-Error $_.Exception.Message
    exit 1
} finally {
    if ($workbook -ne $null) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel -ne $null) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

exit 0
