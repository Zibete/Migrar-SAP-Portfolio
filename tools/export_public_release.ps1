function Copy-PathIfExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    if (-not (Test-Path $SourcePath)) {
        return
    }

    $destParent = Split-Path -Parent $DestinationPath
    if ($destParent) {
        New-Item -ItemType Directory -Force $destParent | Out-Null
    }

    Copy-Item -Recurse -Force $SourcePath $DestinationPath
}

function Remove-IfExists {
    param([string]$PathToRemove)
    if (Test-Path $PathToRemove) {
        Remove-Item -Recurse -Force $PathToRemove
    }
}

function Remove-FilesByPattern {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath,
        [Parameter(Mandatory = $true)]
        [string]$Filter
    )

    if (-not (Test-Path $RootPath)) {
        return
    }

    Get-ChildItem -Path $RootPath -Recurse -File -Filter $Filter | ForEach-Object {
        Remove-Item -Force $_.FullName
    }
}

function Get-PythonLauncher {
    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) {
        return [object[]]@($python.Source)
    }

    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) {
        return [object[]]@($py.Source, "-3")
    }

    return $null
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$distDir = Join-Path $repoRoot "dist"
$releaseDir = Join-Path $distDir "public_release"

New-Item -ItemType Directory -Force $distDir | Out-Null
Remove-IfExists -PathToRemove $releaseDir
New-Item -ItemType Directory -Force $releaseDir | Out-Null

# Core source/docs/tooling for the public technical annex
Copy-PathIfExists -SourcePath (Join-Path $repoRoot "src") -DestinationPath (Join-Path $releaseDir "src")
Copy-PathIfExists -SourcePath (Join-Path $repoRoot "tools") -DestinationPath (Join-Path $releaseDir "tools")
Copy-PathIfExists -SourcePath (Join-Path $repoRoot "docs") -DestinationPath (Join-Path $releaseDir "docs")
Copy-PathIfExists -SourcePath (Join-Path $repoRoot "scripts") -DestinationPath (Join-Path $releaseDir "scripts")
Copy-PathIfExists -SourcePath (Join-Path $repoRoot "assets") -DestinationPath (Join-Path $releaseDir "assets")

Copy-PathIfExists -SourcePath (Join-Path $repoRoot "README.md") -DestinationPath (Join-Path $releaseDir "README.md")
Copy-PathIfExists -SourcePath (Join-Path $repoRoot "LICENSE") -DestinationPath (Join-Path $releaseDir "LICENSE")
Copy-PathIfExists -SourcePath (Join-Path $repoRoot ".gitignore") -DestinationPath (Join-Path $releaseDir ".gitignore")
Copy-PathIfExists -SourcePath (Join-Path $repoRoot ".gitattributes") -DestinationPath (Join-Path $releaseDir ".gitattributes")
Copy-PathIfExists -SourcePath (Join-Path $repoRoot "portfolio\\.gitignore") -DestinationPath (Join-Path $releaseDir "portfolio\\.gitignore")
Copy-PathIfExists -SourcePath (Join-Path $repoRoot "artifacts\\.gitignore") -DestinationPath (Join-Path $releaseDir "artifacts\\.gitignore")

# Exclude internal-only docs and local/runtime artifacts from exported tree
Remove-IfExists -PathToRemove (Join-Path $releaseDir "docs\\AUDIT_PORTFOLIO.md")
Remove-IfExists -PathToRemove (Join-Path $releaseDir "docs\\AUDIT_PUBLIC_RELEASE_DIST.md")
Remove-IfExists -PathToRemove (Join-Path $releaseDir "docs\\PUBLIC_RELEASE_CHECKLIST.md")
Remove-IfExists -PathToRemove (Join-Path $releaseDir "docs\\SCRIPTS.md")
Remove-FilesByPattern -RootPath $releaseDir -Filter "*.frx"
Remove-IfExists -PathToRemove (Join-Path $releaseDir "scripts\\__pycache__")
Remove-IfExists -PathToRemove (Join-Path $releaseDir "scripts\\archivo_index.json")
Remove-IfExists -PathToRemove (Join-Path $releaseDir "portfolio\\CoreTestsHarness.xlsm")
Remove-IfExists -PathToRemove (Join-Path $releaseDir "artifacts\\core-tests.txt")
Remove-IfExists -PathToRemove (Join-Path $releaseDir "artifacts\\core-tests-details.txt")
Remove-IfExists -PathToRemove (Join-Path $releaseDir "artifacts\\prepublish-scan.txt")
Remove-IfExists -PathToRemove (Join-Path $releaseDir "dist")

$fixMojibakeScript = Join-Path $PSScriptRoot "fix_mojibake.py"
$pythonLauncher = @(Get-PythonLauncher)
if ($pythonLauncher.Count -eq 0) {
    Write-Error "No se encontro Python para ejecutar tools/fix_mojibake.py."
    exit 1
}

$pythonArgs = @()
if ($pythonLauncher.Length -gt 1) {
    $pythonArgs = $pythonLauncher[1..($pythonLauncher.Length - 1)]
}

& $pythonLauncher[0] @pythonArgs $fixMojibakeScript $releaseDir
if ($LASTEXITCODE -ne 0) {
    Write-Error "La normalizacion de encoding (mojibake) fallo durante el export."
    exit 1
}

$scanScript = Join-Path $PSScriptRoot "prepublish_scan.ps1"
& $scanScript -TargetRoot $releaseDir
if ($LASTEXITCODE -ne 0) {
    Write-Error "Public release export failed the prepublish scan. Fix findings and rerun."
    exit 1
}

Write-Host ""
Write-Host "Public release prepared at: $releaseDir" -ForegroundColor Green
Write-Host ""
Write-Host "Suggested next steps (new repo, no history):"
Write-Host "  cd `"$releaseDir`""
Write-Host "  git init"
Write-Host "  git add ."
Write-Host "  git commit -m `"Initial public release`""
Write-Host "  # create remote and push manually"

exit 0
