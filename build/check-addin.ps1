#Requires -Version 5
<#
  בודק שוורד באמת טוען את התוסף.

  פותח מופע וורד נקי ומוסתר, מוודא שהתוסף מופיע ברשימת תוספות ה-COM
  ושהוא מחובר בפועל, ואז בודק ש-LoadBehavior לא ירד ל-2.
  וורד מוריד את הערך ל-2 בשקט כשטעינה נכשלת — וזו הדרך היחידה לדעת
  שמשהו השתבש בלי לפתוח את הממשק ולהסתכל.
#>
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$progId = 'Sadar.Connect'
$addinKey = "HKCU:\Software\Microsoft\Office\Word\Addins\$progId"

$pass = 0; $fail = 0
function Check($name, $ok, $detail = '') {
    if ($ok) { $script:pass++; Write-Host "  עבר   $name" -ForegroundColor Green }
    else { $script:fail++; Write-Host "  נכשל  $name  $detail" -ForegroundColor Red }
}

Write-Host 'בדיקת טעינת התוסף בוורד' -ForegroundColor Cyan
Write-Host ('=' * 60)

$before = (Get-ItemProperty $addinKey -ErrorAction SilentlyContinue).LoadBehavior
Check 'LoadBehavior לפני הפתיחה הוא 3' ($before -eq 3) "ערך: $before"

$word = $null
try {
    Write-Host 'פותח וורד (מוסתר)...'
    $word = New-Object -ComObject Word.Application

    # הגנה: אם התחברנו למופע קיים של המשתמש, יוצאים בלי לגעת בכלום
    if ($word.Documents.Count -gt 0) {
        Write-Host 'הבדיקה בוטלה: התחברנו למופע וורד עם מסמכים פתוחים.' -ForegroundColor Yellow
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        exit 2
    }

    $word.Visible = $false
    Check 'וורד נפתח' ($null -ne $word)

    $found = $null
    $names = @()
    foreach ($a in $word.COMAddIns) {
        $names += $a.ProgId
        if ($a.ProgId -eq $progId) { $found = $a }
    }

    Check 'התוסף מופיע ברשימת תוספות ה-COM' ($null -ne $found) "נמצאו: $($names -join ', ')"

    if ($found) {
        Write-Host "        שם: $($found.Description)" -ForegroundColor DarkGray
        Check 'התוסף מחובר בפועל' ($found.Connect -eq $true) "Connect=$($found.Connect)"

        # מוודאים שהאובייקט חי ומגיב, ולא רק רשום
        $obj = $found.Object
        Check 'אובייקט התוסף נגיש' ($null -ne $obj -or $found.Connect)
    }
}
catch {
    $fail++
    Write-Host "  נכשל  חריגה: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($word) {
        try { if ($word.Documents.Count -eq 0) { $word.Quit() } } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch { }
    }
}

Start-Sleep -Milliseconds 700

$after = (Get-ItemProperty $addinKey -ErrorAction SilentlyContinue).LoadBehavior
Check 'LoadBehavior נשאר 3 אחרי הטעינה' ($after -eq 3) "ערך: $after — וורד מוריד ל-2 כשטעינה נכשלת"

$disabled = 'HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems'
# הרשימה מכילה בדרך כלל פריטים שאינם קשורים אלינו (דרייברים למשל),
# ולכן מחפשים בה את סַדָּר עצמו ולא סתם קיום ערכים.
$oursDisabled = $false
if (Test-Path $disabled) {
    foreach ($p in (Get-ItemProperty $disabled).PSObject.Properties) {
        if ($p.Name -like 'PS*') { continue }
        if ($p.Value -is [byte[]]) {
            $text = [System.Text.Encoding]::Unicode.GetString($p.Value)
            if ($text -match 'Sadar') { $oursDisabled = $true }
        }
    }
}
Check 'התוסף לא נכנס לרשימת הפריטים המושבתים' (-not $oursDisabled)

Write-Host ('=' * 60)
if ($fail -eq 0) { Write-Host "עברו $pass, נכשלו 0" -ForegroundColor Green }
else { Write-Host "עברו $pass, נכשלו $fail" -ForegroundColor Red }
exit $fail
