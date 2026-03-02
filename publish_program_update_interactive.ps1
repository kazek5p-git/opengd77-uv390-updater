param(
    [string]$Version = "",
    [switch]$SkipPause
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$publishScript = Join-Path $root "publish_program_update_release.ps1"
$configPath = Join-Path $root "program_update_config.json"

function Get-SuggestedVersion {
    param([string]$CurrentVersion)

    if ([string]::IsNullOrWhiteSpace($CurrentVersion)) {
        return "2026.03.01.1"
    }

    $parts = @()
    foreach ($chunk in ($CurrentVersion -split "[^0-9]+")) {
        if ($chunk -match "^[0-9]+$") {
            $parts += [int]$chunk
        }
    }

    if ($parts.Count -eq 0) {
        return $CurrentVersion
    }

    while ($parts.Count -lt 4) { $parts += 0 }
    if ($parts.Count -gt 4) { $parts = $parts[0..3] }
    $parts[3] += 1
    return ("{0}.{1}.{2}.{3}" -f $parts[0], $parts[1], $parts[2], $parts[3])
}

try {
    if (-not (Test-Path $publishScript)) {
        throw ("Brak skryptu publikacji: " + $publishScript)
    }

    $currentVersion = ""
    if (Test-Path $configPath) {
        try {
            $cfg = Get-Content -Raw -Path $configPath -Encoding UTF8 | ConvertFrom-Json
            if ($cfg.app_version) {
                $currentVersion = [string]$cfg.app_version
            }
        } catch {
            Write-Host ("Uwaga: nie udalo sie odczytac config: " + $_.Exception.Message) -ForegroundColor Yellow
        }
    }

    if ([string]::IsNullOrWhiteSpace($Version)) {
        $suggested = Get-SuggestedVersion -CurrentVersion $currentVersion
        Write-Host "Aktualna wersja (config): $currentVersion"
        Write-Host "Podaj wersje do publikacji (Enter = $suggested): " -NoNewline
        $typed = Read-Host
        if ([string]::IsNullOrWhiteSpace($typed)) {
            $Version = $suggested
        } else {
            $Version = $typed.Trim()
        }
    }

    Write-Host ("Publikacja wersji: " + $Version) -ForegroundColor Cyan
    & $publishScript -Version $Version -UpdateLocalConfig
    Write-Host "Publikacja zakonczona poprawnie." -ForegroundColor Green
} catch {
    Write-Host ("BLAD publikacji: " + $_.Exception.Message) -ForegroundColor Red
    exit 1
} finally {
    if (-not $SkipPause) {
        Write-Host ""
        Read-Host "Nacisnij Enter, aby zamknac"
    }
}
