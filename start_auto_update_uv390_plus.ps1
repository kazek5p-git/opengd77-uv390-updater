param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$PassThruArgs
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$py = "python"
$target = Join-Path $scriptDir "opengd77_auto_update_uv390_plus.py"

if (-not (Test-Path $target)) {
    Write-Host "Brak skryptu: $target"
    exit 1
}

& $py $target @PassThruArgs
exit $LASTEXITCODE
