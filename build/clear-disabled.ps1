#Requires -Version 5
<#
  מציג ומנקה את רשימת התוספות שוורד השבית.

  כשטעינה של תוסף נכשלת פעם אחת, וורד מכניס אותו לרשימה שחורה
  ולא מנסה שוב — גם אחרי שהתקלה תוקנה. זו אחת הסיבות הנפוצות
  לכך ש"התוסף מותקן אבל לא מופיע", והמשתמש אינו מקבל שום הודעה.

  שימוש:
    .\clear-disabled.ps1          הצגה בלבד
    .\clear-disabled.ps1 -Clear   הסרת הערכים ששייכים לסַדָּר
    .\clear-disabled.ps1 -Clear -All   הסרת כל הערכים
#>
param(
    [switch]$Clear,
    [switch]$All
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$found = 0
$removed = 0

$officeRoot = 'HKCU:\Software\Microsoft\Office'
if (-not (Test-Path $officeRoot)) {
    Write-Host 'לא נמצאה התקנת Office.' -ForegroundColor Yellow
    return
}

$versions = Get-ChildItem $officeRoot -ErrorAction SilentlyContinue |
    Where-Object { $_.PSChildName -match '^\d+\.\d+$' }

foreach ($v in $versions) {
    foreach ($app in @('Word', 'Common')) {
        $base = Join-Path $v.PSPath "$app\Resiliency"
        foreach ($sub in @('DisabledItems', 'CrashingAddinList', 'DisabledAddins')) {
            $key = Join-Path $base $sub
            if (-not (Test-Path $key)) { continue }

            $props = (Get-ItemProperty $key).PSObject.Properties |
                Where-Object { $_.Name -notlike 'PS*' }

            if (-not $props) { continue }

            Write-Host "=== $($v.PSChildName)\$app\Resiliency\$sub" -ForegroundColor Cyan

            foreach ($p in $props) {
                $found++

                # הערכים הם blob בינארי שמכיל את הנתיב או ה-ProgId כטקסט Unicode
                $text = ''
                if ($p.Value -is [byte[]]) {
                    $text = [System.Text.Encoding]::Unicode.GetString($p.Value)
                    $text = ($text -split "`0" | Where-Object { $_.Trim().Length -gt 2 }) -join ' | '
                }

                $isOurs = $text -match 'Sadar'
                $tag = if ($isOurs) { '  <-- סַדָּר' } else { '' }
                Write-Host ("   {0} : {1}{2}" -f $p.Name, $text, $tag) `
                    -ForegroundColor $(if ($isOurs) { 'Yellow' } else { 'Gray' })

                if ($Clear -and ($All -or $isOurs)) {
                    Remove-ItemProperty -Path $key -Name $p.Name -Force
                    $removed++
                    Write-Host '        הוסר.' -ForegroundColor Green
                }
            }
        }
    }
}

Write-Host ''
if ($found -eq 0) {
    Write-Host 'אין תוספות מושבתות.' -ForegroundColor Green
}
else {
    Write-Host "נמצאו $found ערכים, הוסרו $removed." -ForegroundColor $(if ($removed -gt 0) { 'Green' } else { 'Yellow' })
    if ($removed -gt 0) {
        Write-Host 'סגרו את וורד ופתחו אותו מחדש.' -ForegroundColor Yellow
    }
}
