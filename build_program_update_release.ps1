param(
    [string]$Version = "",
    [string]$BaseUrl = "https://kazpar.pl/opengd77-updater",
    [switch]$UpdateLocalConfig
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$configPath = Join-Path $root "program_update_config.json"
$autoCheckOnStart = $true
$uiLanguage = "pl"

if ([string]::IsNullOrWhiteSpace($Version) -and (Test-Path $configPath)) {
    try {
        $cfg = Get-Content -Raw -Path $configPath -Encoding UTF8 | ConvertFrom-Json
        if ($cfg.app_version) {
            $Version = [string]$cfg.app_version
        }
        if ($null -ne $cfg.auto_check_on_start) {
            $autoCheckOnStart = [System.Convert]::ToBoolean($cfg.auto_check_on_start)
        }
        if ($cfg.ui_language) {
            $uiLanguage = [string]$cfg.ui_language
        }
    } catch {
        throw ("Nie mozna odczytac program_update_config.json: " + $_.Exception.Message)
    }
} elseif (Test-Path $configPath) {
    try {
        $cfg = Get-Content -Raw -Path $configPath -Encoding UTF8 | ConvertFrom-Json
        if ($null -ne $cfg.auto_check_on_start) {
            $autoCheckOnStart = [System.Convert]::ToBoolean($cfg.auto_check_on_start)
        }
        if ($cfg.ui_language) {
            $uiLanguage = [string]$cfg.ui_language
        }
    } catch {
        throw ("Nie mozna odczytac program_update_config.json: " + $_.Exception.Message)
    }
}

if ([string]::IsNullOrWhiteSpace($Version)) {
    throw "Podaj wersje parametrem -Version lub wpisz app_version w program_update_config.json"
}

$safeVersion = ($Version -replace "[^0-9A-Za-z._-]", "_")
$releaseDir = Join-Path $root ("dist\\program_updater_release\\" + $safeVersion)
$stagingDir = Join-Path $releaseDir "staging"
$packageName = "OpenGD77_UV390_A11y_" + $safeVersion + ".zip"
$packagePath = Join-Path $releaseDir $packageName
$manifestPath = Join-Path $releaseDir "latest.json"

if (Test-Path $releaseDir) {
    Remove-Item -Path $releaseDir -Recurse -Force
}
[void](New-Item -ItemType Directory -Path $stagingDir)

$itemsToPack = @(
    "OpenGD77_UV390_A11y.ps1",
    "start_OpenGD77_UV390_A11y.vbs",
    "program_update_config.json",
    "README_GUI_RUN.txt",
    "assets",
    "dist\\OpenGD77_UV390_BackendCLI.exe"
)

foreach ($relPath in $itemsToPack) {
    $srcPath = Join-Path $root $relPath
    if (-not (Test-Path $srcPath)) {
        throw ("Brak pliku/folderu do paczki: " + $srcPath)
    }

    $dstPath = Join-Path $stagingDir $relPath
    if (Test-Path $srcPath -PathType Container) {
        Copy-Item -Path $srcPath -Destination $dstPath -Recurse -Force
    } else {
        $dstDir = Split-Path -Parent $dstPath
        if (-not (Test-Path $dstDir)) {
            [void](New-Item -ItemType Directory -Path $dstDir -Force)
        }
        Copy-Item -Path $srcPath -Destination $dstPath -Force
    }
}

# Zawsze ustaw wersje i manifest URL wewnatrz paczki update, aby uniknac petli aktualizacji.
$stagedConfigPath = Join-Path $stagingDir "program_update_config.json"
$stagedConfig = [ordered]@{
    app_version = $Version
    manifest_url = ($BaseUrl.TrimEnd("/") + "/latest.json")
    auto_check_on_start = $autoCheckOnStart
    ui_language = $uiLanguage
}
$stagedConfigJson = $stagedConfig | ConvertTo-Json -Depth 3
[System.IO.File]::WriteAllText($stagedConfigPath, $stagedConfigJson, (New-Object System.Text.UTF8Encoding($false)))

Compress-Archive -Path (Join-Path $stagingDir "*") -DestinationPath $packagePath -CompressionLevel Optimal -Force
$hash = (Get-FileHash -Path $packagePath -Algorithm SHA256).Hash.ToLowerInvariant()
$sizeBytes = (Get-Item -Path $packagePath).Length
$releasedAt = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
$packageUrl = ($BaseUrl.TrimEnd("/") + "/" + $packageName)

$manifest = [ordered]@{
    app_name = "OpenGD77 UV390 Updater"
    version = $Version
    package_url = $packageUrl
    sha256 = $hash
    file_size = $sizeBytes
    released_at = $releasedAt
}

$manifestJson = $manifest | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText($manifestPath, $manifestJson, (New-Object System.Text.UTF8Encoding($false)))

if ($UpdateLocalConfig) {
    [System.IO.File]::WriteAllText($configPath, $stagedConfigJson, (New-Object System.Text.UTF8Encoding($false)))
}

[pscustomobject]@{
    Version = $Version
    ReleaseDir = $releaseDir
    PackagePath = $packagePath
    ManifestPath = $manifestPath
    PackageUrl = $packageUrl
    Sha256 = $hash
    SizeBytes = $sizeBytes
}
