$ErrorActionPreference = "Stop"

function Get-LatestHarnessSourceTime {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,
        [Parameter(Mandatory = $true)]
        [string]$BuildScriptPath
    )

    $items = @()
    $items += Get-ChildItem -Path (Join-Path $RepoRoot "src\vba\core\*.bas") -ErrorAction SilentlyContinue
    $items += Get-ChildItem -Path (Join-Path $RepoRoot "src\vba\tests\*.bas") -ErrorAction SilentlyContinue
    $items += Get-Item -Path $BuildScriptPath -ErrorAction SilentlyContinue

    if (-not $items -or $items.Count -eq 0) {
        return Get-Date "2000-01-01"
    }

    return ($items | Sort-Object LastWriteTime -Descending | Select-Object -First 1).LastWriteTime
}

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$repoRoot = $repoRoot.Path

$harnessPath = Join-Path $repoRoot "portfolio\CoreTestsHarness.xlsm"
$artifactsDir = Join-Path $repoRoot "artifacts"
$summaryLogPath = Join-Path $artifactsDir "core-tests.txt"
$detailsLogPath = Join-Path $artifactsDir "core-tests-details.txt"
$buildScript = Join-Path $PSScriptRoot "build_core_harness.ps1"

New-Item -ItemType Directory -Force $artifactsDir | Out-Null

$logLines = @()
$logLines += "Core tests run: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$logLines += "Harness: $harnessPath"
$logLines += "Details log: $detailsLogPath"

$needsBuild = -not (Test-Path $harnessPath)
if (-not $needsBuild) {
    $harnessTime = (Get-Item -Path $harnessPath).LastWriteTime
    $latestSourceTime = Get-LatestHarnessSourceTime -RepoRoot $repoRoot -BuildScriptPath $buildScript
    if ($latestSourceTime -gt $harnessTime) {
        $needsBuild = $true
        $logLines += "Harness rebuild required (sources newer than harness)."
    }
}

if ($needsBuild) {
    if (-not ($logLines -contains "Harness rebuild required (sources newer than harness).")) {
        $logLines += "Harness not found. Building."
    } else {
        $logLines += "Rebuilding harness."
    }

    & $buildScript
    if ($LASTEXITCODE -ne 0) {
        $logLines += "Result: FAIL"
        $logLines += "Build failed. Check build script output and docs/TESTING.md."
        Set-Content -Path $summaryLogPath -Value ($logLines -join [Environment]::NewLine)
        exit 1
    }
}

$excel = $null
$workbook = $null
$exitCode = 0
$failedCount = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open($harnessPath, $null, $true)
    $failedCount = [int]($excel.Run("RunCoreTests"))

    if ($failedCount -gt 0) {
        $exitCode = 1
        $logLines += "Result: FAIL"
        $logLines += "Failed tests: $failedCount"
    } else {
        $logLines += "Result: PASS"
        $logLines += "Failed tests: 0"
    }
} catch {
    $exitCode = 1
    $logLines += "Result: FAIL"
    $logLines += "Error: $($_.Exception.Message)"
    if (Test-Path $detailsLogPath) {
        $logLines += "Check details log (if generated before exception): $detailsLogPath"
    } else {
        $logLines += "Details log not found. This usually means the VBA runner did not start (compile/macro/COM issue)."
    }
} finally {
    if ($workbook -ne $null) {
        $workbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel -ne $null) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Set-Content -Path $summaryLogPath -Value ($logLines -join [Environment]::NewLine)
exit $exitCode
