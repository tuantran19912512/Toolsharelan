@echo off
setlocal
set "LANSHARE_HOST_SCRIPT=%~dp0bin\May-chu.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$p=$env:LANSHARE_HOST_SCRIPT; if(-not(Test-Path -LiteralPath $p)){Write-Host 'Khong tim thay bin\May-chu.ps1. Hay giai nen va chep nguyen thu muc Tool_Share_LAN.'; exit 1}; try { $a='-NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File '+[char]34+$p+[char]34; Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList $a -Verb RunAs -WindowStyle Hidden -ErrorAction Stop; exit 0 } catch { Write-Host $_.Exception.Message; exit 1 }"
if errorlevel 1 pause
endlocal
