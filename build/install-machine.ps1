#Requires -Version 5
<#
  מריץ את ההתקנה ברמת המכונה בחלון מוגבה, ומחזיר את הפלט לקובץ.

  קודם מוסר הרישום שתחת המשתמש: Windows מעדיף את HKCU\Software\Classes
  על פני HKLM, ולכן רישום ישן ברמת המשתמש היה מסתיר את זה שברמת המכונה
  והבדיקה הייתה חוזרת על עצמה בלי שדבר השתנה.
#>
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$root = Split-Path -Parent $PSScriptRoot
$install = Join-Path $PSScriptRoot 'install.ps1'
$log = Join-Path $root 'bin\install-machine.log'

if (Test-Path $log) { Remove-Item $log -Force }

Write-Host 'מסיר רישום קודם ברמת המשתמש...' -ForegroundColor Cyan
& $install -Uninstall | Out-Null

Write-Host 'מבקש הרשאות מנהל — אשרו את חלון האישור...' -ForegroundColor Yellow

# עוטפים ב-try/catch: בלי זה חלון מוגבה שנכשל פשוט נסגר,
# והשגיאה נעלמת יחד איתו.
$cmd = "try { & '$install' -Machine *>&1 | Out-File -FilePath '$log' -Encoding utf8 } " +
       "catch { \$_ | Format-List * -Force | Out-File -FilePath '$log' -Append -Encoding utf8 }"
$p = Start-Process -FilePath 'powershell.exe' -Verb RunAs -Wait -PassThru `
    -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', $cmd

Write-Host ''
if (Test-Path $log) {
    Get-Content $log -Encoding UTF8 | ForEach-Object { Write-Host $_ }
}
else {
    Write-Host 'לא נוצר קובץ פלט — ייתכן שהאישור נדחה.' -ForegroundColor Red
}

Write-Host ''
Write-Host "קוד יציאה: $($p.ExitCode)" -ForegroundColor DarkGray
