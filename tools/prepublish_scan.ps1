param(
    [string]$TargetRoot,
    [string]$ExtraTermsFile
)

$ErrorActionPreference = "Stop"

function Get-DefaultTargetRoot {
    return (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}

function Get-RelativePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BasePath,
        [Parameter(Mandatory = $true)]
        [string]$FullPath
    )

    $baseAbs = [IO.Path]::GetFullPath($BasePath)
    $fullAbs = [IO.Path]::GetFullPath($FullPath)
    $baseUri = [Uri]($baseAbs + [IO.Path]::DirectorySeparatorChar)
    $fullUri = [Uri]$fullAbs
    return [Uri]::UnescapeDataString($baseUri.MakeRelativeUri($fullUri).ToString()).Replace("/", "\")
}

function Is-ExcludedPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RelativePath,
        [Parameter(Mandatory = $true)]
        [string[]]$ExcludedPrefixes,
        [Parameter(Mandatory = $true)]
        [string[]]$ExcludedFiles
    )

    foreach ($prefix in $ExcludedPrefixes) {
        if ($RelativePath -ieq $prefix) { return $true }
        if ($RelativePath.StartsWith($prefix + "\", [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }

    foreach ($file in $ExcludedFiles) {
        if ($RelativePath -ieq $file) { return $true }
    }

    return $false
}

function Get-ScanFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath
    )

    $allowedNames = @(".gitignore", ".gitattributes", "README.md", "LICENSE")
    $allowedExts = @(".bas", ".cls", ".frm", ".md", ".ps1", ".py", ".txt", ".json", ".yml", ".yaml")
    $excludedPrefixes = @(".git", ".idea", "portfolio_pack", "portfolio", "artifacts", "dist")
    $excludedFiles = @(
        "docs\AUDIT_PORTFOLIO.md",
        "tools\prepublish_scan.ps1",
        "tools\prepublish_terms.local.txt",
        "Migrar SAP.txt"
    )

    $files = Get-ChildItem -Path $RootPath -Recurse -File
    $scanFiles = New-Object System.Collections.Generic.List[object]

    foreach ($file in $files) {
        $rel = Get-RelativePath -BasePath $RootPath -FullPath $file.FullName
        if (Is-ExcludedPath -RelativePath $rel -ExcludedPrefixes $excludedPrefixes -ExcludedFiles $excludedFiles) { continue }

        $ext = $file.Extension.ToLowerInvariant()
        if (($allowedNames -icontains $file.Name) -or ($allowedExts -contains $ext)) {
            $scanFiles.Add([PSCustomObject]@{
                FullPath = $file.FullName
                RelativePath = $rel
            })
        }
    }

    return $scanFiles
}

function Get-LocalExtraTerms {
    param(
        [Parameter(Mandatory = $false)]
        [string]$FilePath
    )

    $terms = New-Object System.Collections.Generic.List[string]

    if ([string]::IsNullOrWhiteSpace($FilePath)) {
        return $terms
    }

    $candidatePath = $FilePath
    if (-not (Test-Path $candidatePath)) {
        return $terms
    }

    foreach ($line in (Get-Content -Path $candidatePath -ErrorAction Stop)) {
        $trimmed = $line.Trim()
        if ($trimmed -eq "") { continue }
        if ($trimmed.StartsWith("#")) { continue }
        $terms.Add($trimmed) | Out-Null
    }

    return $terms
}

function Add-MatchRecords {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Category,
        [Parameter(Mandatory = $true)]
        [string]$RegexPattern,
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$Files,
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$Records,
        [switch]$CaseSensitive
    )

    foreach ($file in $Files) {
        $params = @{
            Path = $file.FullPath
            Pattern = $RegexPattern
            AllMatches = $true
        }
        if (-not $CaseSensitive) {
            $params.CaseSensitive = $false
        } else {
            $params.CaseSensitive = $true
        }

        $matches = Select-String @params
        foreach ($m in $matches) {
            $lineText = ($m.Line -replace "\s+", " ").Trim()
            foreach ($subMatch in $m.Matches) {
                $Records.Add([PSCustomObject]@{
                    Category = $Category
                    File = $file.RelativePath
                    Line = $m.LineNumber
                    Match = $subMatch.Value
                    Preview = $lineText
                }) | Out-Null
            }
        }
    }
}

function Add-UrlRecords {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$Files,
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$Records,
        [Parameter(Mandatory = $true)]
        [string[]]$AllowedHosts
    )

    $urlRegex = "https?://[^\s'""`<>]+"

    foreach ($file in $Files) {
        $matches = Select-String -Path $file.FullPath -Pattern $urlRegex -AllMatches
        foreach ($m in $matches) {
            $lineText = ($m.Line -replace "\s+", " ").Trim()
            foreach ($subMatch in $m.Matches) {
                $urlValue = $subMatch.Value
                $urlHost = ""
                try {
                    $urlHost = ([Uri]$urlValue).Host.ToLowerInvariant()
                } catch {
                    $urlHost = ""
                }

                if ($urlHost -and ($AllowedHosts -contains $urlHost)) {
                    continue
                }

                $Records.Add([PSCustomObject]@{
                    Category = "url_not_whitelisted"
                    File = $file.RelativePath
                    Line = $m.LineNumber
                    Match = $urlValue
                    Preview = $lineText
                }) | Out-Null
            }
        }
    }
}

function Write-ScanReport {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetRootPath,
        [Parameter(Mandatory = $true)]
        [string]$LogPath,
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$Records,
        [Parameter(Mandatory = $true)]
        [string[]]$PatternNames
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("Prepublish scan run: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')") | Out-Null
    $lines.Add("Target root: $TargetRootPath") | Out-Null
    $lines.Add("Allowed URL hosts: servicioscf.afip.gob.ar, learn.microsoft.com") | Out-Null
    $lines.Add("Pattern categories: " + ($PatternNames -join ", ")) | Out-Null
    $lines.Add("") | Out-Null

    if ($Records.Count -eq 0) {
        $lines.Add("Result: PASS") | Out-Null
        $lines.Add("No matches found for blocked patterns.") | Out-Null
    } else {
        $lines.Add("Result: FAIL") | Out-Null
        $lines.Add("Matches found: $($Records.Count)") | Out-Null
        $lines.Add("") | Out-Null

        $grouped = $Records | Group-Object Category | Sort-Object Name
        foreach ($group in $grouped) {
            $lines.Add("[" + $group.Name + "] " + $group.Count + " match(es)") | Out-Null
            $fileGroups = $group.Group | Group-Object File | Sort-Object Name
            foreach ($fg in $fileGroups) {
                $lines.Add(" - " + $fg.Name + " (" + $fg.Count + ")") | Out-Null
            }

            foreach ($item in ($group.Group | Select-Object -First 10)) {
                $lines.Add("   " + $item.File + ":" + $item.Line + " :: " + $item.Preview) | Out-Null
            }
            if ($group.Count -gt 10) {
                $lines.Add("   ...") | Out-Null
            }
            $lines.Add("") | Out-Null
        }
    }

    [IO.File]::WriteAllLines($LogPath, $lines, (New-Object System.Text.UTF8Encoding($false)))
}

$resolvedTargetRoot = if ([string]::IsNullOrWhiteSpace($TargetRoot)) { Get-DefaultTargetRoot } else { (Resolve-Path $TargetRoot).Path }
$artifactsDir = Join-Path $resolvedTargetRoot "artifacts"
New-Item -ItemType Directory -Force $artifactsDir | Out-Null
$logPath = Join-Path $artifactsDir "prepublish-scan.txt"

$files = Get-ScanFiles -RootPath $resolvedTargetRoot
$records = New-Object System.Collections.Generic.List[object]

$forbiddenPortalNames = @(
    ("Store" + "Book")
)

$forbiddenAcronyms = @(
    # Intencionalmente vacio por defecto para evitar falsos positivos.
    # Ejemplo local (no versionado): RW
)

$forbiddenRegexDefs = @(
    @{ Name = "corp_email"; Regex = "[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}" },
    @{ Name = "windows_user_path"; Regex = "[A-Za-z]:\\Users\\[^\\]+" },
    @{ Name = "obvious_secret_tokens"; Regex = "(?i)\b(x-api-key|api[_-]?key|client[_-]?secret|authorization\s*:\s*bearer)\b" },
    @{ Name = "hardcoded_sheet_password"; Regex = 'Password:=\s*"1234"' },
    @{ Name = "forbidden_parser_prefix"; Regex = "\b" + [regex]::Escape(("mod" + "Leer")) + "[A-Za-z0-9_]*\b" }
)

$patternDefs = New-Object System.Collections.Generic.List[object]

foreach ($def in $forbiddenRegexDefs) {
    $patternDefs.Add([PSCustomObject]@{
        Name = $def.Name
        Regex = $def.Regex
    }) | Out-Null
}

foreach ($term in $forbiddenPortalNames) {
    $patternDefs.Add([PSCustomObject]@{
        Name = "forbidden_portal_name"
        Regex = [regex]::Escape($term)
    }) | Out-Null
}

foreach ($acronym in $forbiddenAcronyms) {
    $patternDefs.Add([PSCustomObject]@{
        Name = "forbidden_acronyms"
        Regex = "\b" + [regex]::Escape($acronym) + "\b"
    }) | Out-Null
}

$localExtraTerms = Get-LocalExtraTerms -FilePath $ExtraTermsFile
foreach ($term in $localExtraTerms) {
    $patternDefs.Add([PSCustomObject]@{
        Name = "forbidden_terms_from_file"
        Regex = [regex]::Escape($term)
    }) | Out-Null
}

foreach ($def in $patternDefs) {
    Add-MatchRecords -Category $def.Name -RegexPattern $def.Regex -Files $files -Records $records
}

Add-UrlRecords -Files $files -Records $records -AllowedHosts @("servicioscf.afip.gob.ar", "learn.microsoft.com", "img.shields.io")

Write-ScanReport `
    -TargetRootPath $resolvedTargetRoot `
    -LogPath $logPath `
    -Records $records `
    -PatternNames (@($patternDefs | ForEach-Object { $_.Name } | Select-Object -Unique) + "url_not_whitelisted")

if ($records.Count -gt 0) {
    Write-Host "Prepublish scan failed. Review: $logPath" -ForegroundColor Red
    exit 1
}

Write-Host "Prepublish scan passed. Report: $logPath" -ForegroundColor Green
exit 0
