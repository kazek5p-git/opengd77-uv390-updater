param(
    [string]$Version = "",
    [string]$BaseUrl = "https://kazpar.pl/opengd77-updater",
    [string]$ServerHost = "kazpar.pl",
    [int]$Port = 1024,
    [string]$User = "root",
    [string]$SshKeyPath = "C:\\Users\\Kazek\\.ssh\\kazek_server",
    [string]$RemoteDir = "/home/kazek/www/opengd77-updater",
    [switch]$UpdateLocalConfig
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$buildScript = Join-Path $root "build_program_update_release.ps1"
if (-not (Test-Path $buildScript)) {
    throw ("Brak skryptu build: " + $buildScript)
}
if (-not (Test-Path $SshKeyPath)) {
    throw ("Brak klucza SSH: " + $SshKeyPath)
}

Write-Host "[1/5] Build paczki update..." -ForegroundColor Cyan
$buildResult = & $buildScript -Version $Version -BaseUrl $BaseUrl -UpdateLocalConfig:$UpdateLocalConfig
if ($buildResult -is [Array]) {
    $buildResult = $buildResult[-1]
}

$packagePath = [string]$buildResult.PackagePath
$manifestPath = [string]$buildResult.ManifestPath
$sshTarget = ($User + "@" + $ServerHost)

Write-Host "[2/5] Tworze katalog na serwerze: $RemoteDir" -ForegroundColor Cyan
& ssh -p $Port -i $SshKeyPath $sshTarget ("mkdir -p '" + $RemoteDir + "'")
if ($LASTEXITCODE -ne 0) {
    throw ("SSH mkdir failed, code=" + $LASTEXITCODE)
}

Write-Host "[3/5] Wysylam paczke i manifest..." -ForegroundColor Cyan
$scpArgs = @(
    "-P", $Port.ToString(),
    "-i", $SshKeyPath,
    $packagePath,
    $manifestPath,
    ($sshTarget + ":" + $RemoteDir + "/")
)
& scp @scpArgs
if ($LASTEXITCODE -ne 0) {
    throw ("SCP failed, code=" + $LASTEXITCODE)
}

Write-Host "[4/5] Weryfikacja plikow na serwerze..." -ForegroundColor Cyan
$pkgName = Split-Path -Leaf $packagePath
& ssh -p $Port -i $SshKeyPath $sshTarget ("ls -lh '" + $RemoteDir + "' && sha256sum '" + $RemoteDir + "/" + $pkgName + "'")
if ($LASTEXITCODE -ne 0) {
    throw ("SSH verify failed, code=" + $LASTEXITCODE)
}

Write-Host "[5/5] Weryfikacja online manifestu..." -ForegroundColor Cyan
$onlineManifestUrl = $BaseUrl.TrimEnd("/") + "/latest.json"
$online = Invoke-WebRequest -Uri $onlineManifestUrl -UseBasicParsing -TimeoutSec 30
$onlineContent = $online.Content.TrimStart([char]0xFEFF)
$onlineObj = $onlineContent | ConvertFrom-Json

[pscustomobject]@{
    Version = [string]$buildResult.Version
    PackagePath = $packagePath
    ManifestPath = $manifestPath
    PackageUrl = [string]$buildResult.PackageUrl
    OnlineManifestUrl = $onlineManifestUrl
    OnlineVersion = [string]$onlineObj.version
    OnlinePackageUrl = [string]$onlineObj.package_url
    OnlineSha256 = [string]$onlineObj.sha256
}
