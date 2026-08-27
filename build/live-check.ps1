#Requires -Version 5
<#
  בדיקה חיה: פותח את וורד כמו שמשתמש פותח אותו — לא דרך אוטומציה —
  ובודק אם התוסף נטען והלשונית נבנתה.

  ההבחנה הזו קריטית: וורד שנפתח דרך COM רץ במצב /automation ומתנהג
  אחרת עם תוספות, ולכן בדיקה דרכו אינה מעידה על מה שהמשתמש יראה.
#>
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$progId = 'Sadar.Connect'
$pass = 0; $fail = 0
function Check($name, $ok, $detail = '') {
    if ($ok) { $script:pass++; Write-Host "  עבר   $name" -ForegroundColor Green }
    else { $script:fail++; Write-Host "  נכשל  $name  $detail" -ForegroundColor Red }
}

Write-Host 'בדיקה חיה של התוסף בוורד' -ForegroundColor Cyan
Write-Host ('=' * 60)

# ----- ניקוי תהליכים תלויים -----
# תהליך וורד שנשאר מלפני הרישום לא יראה את התוסף ויזהם את הבדיקה.
$leftover = @(Get-Process WINWORD -ErrorAction SilentlyContinue)
foreach ($p in $leftover) {
    if ($p.MainWindowHandle -ne 0) {
        Write-Host "וורד פתוח עם חלון (PID $($p.Id)). סגרו אותו והריצו שוב." -ForegroundColor Yellow
        exit 3
    }
}
if ($leftover.Count -gt 0) {
    Write-Host "מסיים $($leftover.Count) תהליכי וורד תלויים ללא חלון..." -ForegroundColor DarkGray
    try {
        $w = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
        if ($w.Documents.Count -eq 0) { $w.Quit() }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null
    } catch { }
    Start-Sleep -Seconds 2
    Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -eq 0 } |
        ForEach-Object { try { $_.Kill() } catch { } }
    Start-Sleep -Seconds 2
}

Check 'אין תהליכי וורד פעילים' (@(Get-Process WINWORD -ErrorAction SilentlyContinue).Count -eq 0)

# ----- פתיחה רגילה -----
$winword = @(
    "$env:ProgramFiles\Microsoft Office\root\Office16\WINWORD.EXE",
    "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\WINWORD.EXE"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

Check 'נמצא WINWORD.EXE' ($null -ne $winword)
if (-not $winword) { exit 1 }

Write-Host 'פותח את וורד כרגיל...' -ForegroundColor DarkGray
Start-Process -FilePath $winword
Start-Sleep -Seconds 12

$word = $null
for ($i = 0; $i -lt 10; $i++) {
    try { $word = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application'); break }
    catch { Start-Sleep -Seconds 2 }
}
Check 'וורד נפתח וזמין' ($null -ne $word)
if (-not $word) { exit 1 }

try {
    $names = @(); $found = $null
    foreach ($a in $word.COMAddIns) {
        $names += $a.ProgId
        if ($a.ProgId -eq $progId) { $found = $a }
    }

    Check 'התוסף מופיע ברשימת תוספות ה-COM' ($null -ne $found) "נמצאו: $($names -join ', ')"

    if ($found) {
        Check 'התוסף מחובר' ($found.Connect -eq $true) "Connect=$($found.Connect)"
        Write-Host "        $($found.Description)" -ForegroundColor DarkGray
    }

    # ----- הלשונית -----
    # וורד מוסיף כל לשונית מותאמת ל-CommandBars בשם "Ribbon".
    # אם התוסף נטען אבל ה-XML פסול, הלשונית לא תיווצר והתוסף ייראה תקין.
    $ribbonOk = $false
    try {
        $ctl = $word.CommandBars.GetEnabledMso('SadarAnalyze')
        $ribbonOk = $true
    } catch {
        # GetEnabledMso עובד רק על פקדים מובנים; ננסה דרך אחרת
        try { $ribbonOk = ($word.CommandBars.GetVisibleMso('SadarAnalyze') -ne $null) } catch { }
    }
    if ($found -and $found.Connect) {
        Check 'הלשונית נבנתה (פקד סַדָּר נגיש)' $ribbonOk 'ייתכן שהבדיקה אינה נתמכת בגרסה זו'
    }
}
finally {
    try {
        if ($word.Documents.Count -eq 0) { $word.Quit() }
        else { Write-Host 'נשארו מסמכים פתוחים — וורד לא נסגר.' -ForegroundColor Yellow }
    } catch { }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch { }
}

Write-Host ('=' * 60)
if ($fail -eq 0) { Write-Host "עברו $pass, נכשלו 0" -ForegroundColor Green }
else { Write-Host "עברו $pass, נכשלו $fail" -ForegroundColor Red }
exit $fail
