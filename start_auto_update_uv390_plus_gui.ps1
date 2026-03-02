$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$py = "C:\Python314\python.exe"
if (-not (Test-Path $py)) {
    $py = "python"
}

& $py "$scriptDir\opengd77_auto_update_uv390_plus_gui.py"
