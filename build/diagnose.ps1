#Requires -Version 5
<#
  אבחון עמוק: סורק את כל המקומות שבהם Office יכול לחסום תוסף בשקט.

  Office מנהל כמה רשימות חסימה נפרדות זו מזו, וכולן פועלות בלי
  להודיע למשתמש דבר. כשתוסף "מותקן אבל לא עובד", התשובה כמעט תמיד
  נמצאת באחת מהן.
#>
param([switch]$Fix)

$ErrorActionPreference = 'Continue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$clsid = '{7A3C1F42-9E5B-4D18-B6C4-2E9A5D3F8B71}'
$progId = 'Sadar.Connect'
$issues = @()

Write-Host 'אבחון חסימות Office' -ForegroundColor Cyan
Write-Host ('=' * 62)

# ---- 1. רשימת תאימות COM ----
# Compatibility Flags = 0x400 פירושו "Office השבית את התוסף"
Write-Host ''
Write-Host '1. רשימת תאימות COM' -ForegroundColor White
$compatRoots = @(
    'HKCU:\Software\Microsoft\Office\16.0\Common\COM Compatibility',
    'HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\COM Compatibility',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Common\COM Compatibility'
)
$compatFound = $false
foreach ($r in $compatRoots) {
    $k = Join-Path $r $clsid
    if (Test-Path $k) {
        $compatFound = $true
        $flags = (Get-ItemProperty $k).'Compatibility Flags'
        Write-Host ("   נמצא: $r") -ForegroundColor Yellow
        Write-Host ("   Compatibility Flags = 0x{0:X}" -f $flags) -ForegroundColor Yellow
        if ($flags -band 0x400) {
            $issues += 'Office השבית את התוסף ברשימת התאימות (0x400)'
            if ($Fix) {
                Set-ItemProperty $k -Name 'Compatibility Flags' -Value 0 -Type DWord
                Write-Host '   תוקן: הדגל אופס.' -ForegroundColor Green
            }
        }
    }
}
if (-not $compatFound) { Write-Host '   נקי.' -ForegroundColor Green }

# ---- 2. LoadBehavior ----
Write-Host ''
Write-Host '2. התנהגות טעינה' -ForegroundColor White
$addinKey = "HKCU:\Software\Microsoft\Office\Word\Addins\$progId"
if (Test-Path $addinKey) {
    $lb = (Get-ItemProperty $addinKey).LoadBehavior
    Write-Host "   LoadBehavior = $lb"
    if ($lb -ne 3) {
        $issues += "LoadBehavior הוא $lb — וורד כיבה את התוסף אחרי כשל טעינה"
        if ($Fix) {
            Set-ItemProperty $addinKey -Name 'LoadBehavior' -Value 3 -Type DWord
            Write-Host '   תוקן: הוחזר ל-3.' -ForegroundColor Green
        }
    }
    else { Write-Host '   תקין.' -ForegroundColor Green }
}
else { $issues += 'מפתח התוסף אינו קיים' }

# ---- 3. פריטים מושבתים ----
Write-Host ''
Write-Host '3. פריטים מושבתים' -ForegroundColor White
$res = 'HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency'
$ourDisabled = $false
foreach ($sub in @('DisabledItems', 'CrashingAddinList')) {
    $k = Join-Path $res $sub
    if (-not (Test-Path $k)) { continue }
    foreach ($p in (Get-ItemProperty $k).PSObject.Properties) {
        if ($p.Name -like 'PS*') { continue }
        if ($p.Value -is [byte[]]) {
            $t = [System.Text.Encoding]::Unicode.GetString($p.Value)
            if ($t -match 'Sadar') {
                $ourDisabled = $true
                Write-Host "   נמצא ב-$sub" -ForegroundColor Yellow
                if ($Fix) {
                    Remove-ItemProperty $k -Name $p.Name -Force
                    Write-Host '   תוקן: הוסר.' -ForegroundColor Green
                }
            }
        }
    }
}
if (-not $ourDisabled) { Write-Host '   נקי.' -ForegroundColor Green }
else { $issues += 'התוסף נמצא ברשימת הפריטים המושבתים' }

# ---- 4. מדיניות אבטחה ----
Write-Host ''
Write-Host '4. מדיניות אבטחה של Office' -ForegroundColor White
$secRoots = @(
    'HKCU:\Software\Microsoft\Office\16.0\Word\Security',
    'HKCU:\Software\Policies\Microsoft\Office\16.0\Word\Security',
    'HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Word\Security',
    'HKCU:\Software\Policies\Microsoft\Office\16.0\Common\Security'
)
$secIssue = $false
foreach ($r in $secRoots) {
    if (-not (Test-Path $r)) { continue }
    $props = Get-ItemProperty $r
    foreach ($n in @('DisableAllAddins', 'RequireAddinSig', 'BlockContentExecutionFromInternet')) {
        $v = $props.$n
        if ($null -ne $v -and $v -ne 0) {
            Write-Host "   $r -> $n = $v" -ForegroundColor Yellow
            $issues += "$n מופעל וחוסם תוספות"
            $secIssue = $true
        }
    }
}
if (-not $secIssue) { Write-Host '   נקי.' -ForegroundColor Green }

# ---- 5. שגיאות זמן ריצה ----
Write-Host ''
Write-Host '5. שגיאות טעינה אחרונות' -ForegroundColor White
$events = Get-WinEvent -FilterHashtable @{LogName = 'Application'; StartTime = (Get-Date).AddHours(-1) } -ErrorAction SilentlyContinue |
    Where-Object { $_.ProviderName -match 'NET Runtime|Application Error' }
if ($events) {
    foreach ($e in ($events | Select-Object -First 3)) {
        $m = $e.Message -replace '\s+', ' '
        Write-Host ("   {0:HH:mm:ss}  {1}" -f $e.TimeCreated, $m.Substring(0, [Math]::Min(150, $m.Length))) -ForegroundColor Yellow
    }
    $issues += 'יש שגיאות זמן ריצה ביומן האירועים'
}
else { Write-Host '   אין.' -ForegroundColor Green }

# ---- סיכום ----
Write-Host ''
Write-Host ('=' * 62)
if ($issues.Count -eq 0) {
    Write-Host 'לא נמצאו חסימות.' -ForegroundColor Green
}
else {
    Write-Host 'ממצאים:' -ForegroundColor Yellow
    foreach ($i in $issues) { Write-Host "  - $i" -ForegroundColor Yellow }
    if (-not $Fix) {
        Write-Host ''
        Write-Host 'להסרת החסימות: .\diagnose.ps1 -Fix' -ForegroundColor Cyan
    }
}
