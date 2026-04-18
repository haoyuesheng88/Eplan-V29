Set-StrictMode -Version 2

function Get-VersionSortValue {
    param(
        [string]$Text
    )

    if ($Text -match '\d+(?:\.\d+)+') {
        return [version]$matches[0]
    }

    return [version]'0.0'
}

function Get-PreferredDirectory {
    param(
        [string[]]$Paths,
        [string]$ExecutableName
    )

    $candidates = foreach ($path in $Paths) {
        if ([string]::IsNullOrWhiteSpace($path)) {
            continue
        }

        if (Test-Path -LiteralPath (Join-Path $path $ExecutableName)) {
            [pscustomobject]@{
                Path = $path
                Version = Get-VersionSortValue -Text $path
            }
        }
    }

    return $candidates |
        Sort-Object Version -Descending |
        Select-Object -ExpandProperty Path -First 1
}

function Get-BinDirectoriesUnderRoot {
    param(
        [string[]]$Roots,
        [string]$ExecutableName
    )

    $results = New-Object System.Collections.Generic.List[string]

    foreach ($root in $Roots) {
        if (-not (Test-Path -LiteralPath $root)) {
            continue
        }

        Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $binPath = Join-Path $_.FullName 'Bin'
            if (Test-Path -LiteralPath (Join-Path $binPath $ExecutableName)) {
                $results.Add($binPath)
            }
        }
    }

    return $results
}

function Resolve-EplanInstallation {
    $platformBins = New-Object System.Collections.Generic.List[string]
    $offlineBins = New-Object System.Collections.Generic.List[string]

    $eplanProcesses = Get-Process -Name 'EPLAN' -ErrorAction SilentlyContinue
    foreach ($process in $eplanProcesses) {
        if ([string]::IsNullOrWhiteSpace($process.Path)) {
            continue
        }
        $platformBins.Add((Split-Path -Parent $process.Path))
    }

    $w3uProcesses = Get-Process -Name 'W3u' -ErrorAction SilentlyContinue
    foreach ($process in $w3uProcesses) {
        if ([string]::IsNullOrWhiteSpace($process.Path)) {
            continue
        }
        $offlineBins.Add((Split-Path -Parent $process.Path))
    }

    foreach ($path in @(Get-BinDirectoriesUnderRoot -Roots @(
        'C:\Program Files\EPLAN\Platform',
        'C:\Program Files (x86)\EPLAN\Platform'
    ) -ExecutableName 'EPLAN.exe')) {
        $platformBins.Add($path)
    }

    foreach ($path in @(Get-BinDirectoriesUnderRoot -Roots @(
        'C:\Program Files\EPLAN\Electric P8',
        'C:\Program Files (x86)\EPLAN\Electric P8'
    ) -ExecutableName 'W3u.exe')) {
        $offlineBins.Add($path)
    }

    $platformBin = Get-PreferredDirectory -Paths $platformBins -ExecutableName 'EPLAN.exe'
    $offlineBin = Get-PreferredDirectory -Paths $offlineBins -ExecutableName 'W3u.exe'

    if (-not $offlineBin -and $platformBin) {
        $version = Split-Path -Leaf (Split-Path -Parent $platformBin)
        $versionMatch = @(
            "C:\Program Files\EPLAN\Electric P8\$version\Bin",
            "C:\Program Files (x86)\EPLAN\Electric P8\$version\Bin"
        )
        $offlineBin = Get-PreferredDirectory -Paths $versionMatch -ExecutableName 'W3u.exe'
    }

    if (-not $platformBin) {
        throw 'Could not find EPLAN.exe. Start EPLAN first or install EPLAN Electric P8 in a standard location.'
    }

    if (-not $offlineBin) {
        throw 'Could not find W3u.exe. Start EPLAN first or install EPLAN Electric P8 in a standard location.'
    }

    [pscustomobject]@{
        PlatformBin = $platformBin
        OfflineBin = $offlineBin
        EplanExe = Join-Path $platformBin 'EPLAN.exe'
        W3uExe = Join-Path $offlineBin 'W3u.exe'
    }
}

function Resolve-OpenEplanProjectPath {
    $candidateSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    $processes = Get-Process -Name 'EPLAN' -ErrorAction SilentlyContinue |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_.MainWindowTitle) }

    foreach ($process in $processes) {
        $title = $process.MainWindowTitle
        $candidates = New-Object System.Collections.Generic.List[string]

        foreach ($match in [regex]::Matches($title, '(?<path>[A-Za-z]:\\[^<>:"|?*\r\n]+?\.elk)')) {
            $candidate = $match.Groups['path'].Value.Trim(' ', '"')
            if (-not [string]::IsNullOrWhiteSpace($candidate)) {
                $candidates.Add($candidate)
            }
        }

        foreach ($match in [regex]::Matches($title, '(?<path>[A-Za-z]:\\[^<>:"|?*\r\n]+?)(?=\s+-\s+|$)')) {
            $candidate = $match.Groups['path'].Value.Trim(' ', '"')
            if (-not [string]::IsNullOrWhiteSpace($candidate)) {
                $candidates.Add($candidate)
            }
        }

        foreach ($segment in ($title -split '\s+-\s+')) {
            $candidate = $segment.Trim(' ', '"')
            if ($candidate -match '^[A-Za-z]:\\') {
                $candidates.Add($candidate)
            }
        }

        foreach ($basePath in $candidates) {
            if ([string]::IsNullOrWhiteSpace($basePath)) {
                continue
            }

            if ($candidateSet.Contains($basePath)) {
                continue
            }

            $candidateSet.Add($basePath) | Out-Null

            if (Test-Path -LiteralPath $basePath) {
                return $basePath
            }

            if (-not [IO.Path]::HasExtension($basePath)) {
                $elkCandidate = $basePath + '.elk'
                if ($candidateSet.Add($elkCandidate) -and (Test-Path -LiteralPath $elkCandidate)) {
                    return $elkCandidate
                }

                $directory = Split-Path -Parent $basePath
                $leaf = Split-Path -Leaf $basePath
                if (Test-Path -LiteralPath $directory) {
                    $match = Get-ChildItem -LiteralPath $directory -Filter '*.elk' -File -ErrorAction SilentlyContinue |
                        Where-Object { $_.BaseName -eq $leaf } |
                        Select-Object -First 1

                    if ($match -ne $null) {
                        return $match.FullName
                    }
                }
            }
        }
    }

    throw 'Could not resolve an .elk project path from the open EPLAN window title. Pass -ProjectPath explicitly if the title shows only the project name or another non-path format.'
}
