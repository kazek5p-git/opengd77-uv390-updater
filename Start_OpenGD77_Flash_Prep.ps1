$base = 'C:\Users\Kazek\OpenGD77_UV390_Plus10W'
$cps  = 'C:\Program Files (x86)\OpenGD77CPS\OpenGD77CPS.exe'
Start-Process explorer.exe $base
if (Test-Path $cps) { Start-Process $cps }
