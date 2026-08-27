#Requires -Version 5
<#
  המסך שהמשתמש רואה בלחיצה כפולה.

  כל הטקסט העברי יושב כאן ולא בקובצי ה-cmd: cmd.exe קורא כל שורה
  לפי דף הקוד שפעיל באותו רגע, ולכן עברית בתוכו נהרסת על חלק
  מהמחשבים בלי שום התראה.
#>
param(
    [ValidateSet('install', 'uninstall', 'check')]
    [string]$Action = 'install'
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = 'סַדָּר'

$here = $PSScriptRoot
$root = Split-Path -Parent $here
$install = Join-Path $here 'install.ps1'

function Line { Write-Host ('  ' + ('─' * 46)) -ForegroundColor DarkGray }

function Header([string]$subtitle) {
    Write-Host ''
    Line
    Write-Host '   סַדָּר — עימוד חכם לוורד' -ForegroundColor Cyan
    Write-Host "   $subtitle" -ForegroundColor Gray
    Line
    Write-Host ''
}

function Wait-Key {
    Write-Host ''
    Write-Host '  הקישו Enter לסגירה.' -ForegroundColor DarkGray
    [void](Read-Host)
}

function Test-WordRunning {
    $procs = @(Get-Process WINWORD -ErrorAction SilentlyContinue)
    if ($procs.Count -eq 0) { return $false }

    Write-Host '  וורד פתוח כרגע.' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '  סגרו את וורד לגמרי, ואז לחצו שוב על הקובץ הזה.'
    return $true
}

# ---- הקובץ חסום? ----
# זו נקודת הכישלון הנפוצה ביותר: Windows מסמן כל קובץ שהורד,
# והתוסף פשוט לא נטען בלי שאיש מסביר למה.
function Test-Blocked {
    $dll = Join-Path $root 'bin\Sadar.Addin.dll'
    if (-not (Test-Path $dll)) { return $false }
    return $null -ne (Get-Item $dll -Stream Zone.Identifier -ErrorAction SilentlyContinue)
}

switch ($Action) {

    'install' {
        Header 'התקנה'

        if (Test-WordRunning) { Wait-Key; exit 1 }

        if (Test-Blocked) {
            Write-Host '  שימו לב: הקבצים מסומנים כ"הגיעו מהאינטרנט".' -ForegroundColor Yellow
            Write-Host '  מסירים את הסימון אוטומטית...' -ForegroundColor DarkGray
            Write-Host ''
        }

        $failed = $false
        try { & $install }
        catch {
            $failed = $true
            Write-Host ''
            Write-Host "  ההתקנה נכשלה: $($_.Exception.Message)" -ForegroundColor Red
        }

        Write-Host ''
        if ($failed) {
            Line
            Write-Host '  מה כדאי לנסות:' -ForegroundColor Yellow
            Write-Host '   1. לחצו לחיצה ימנית על קובץ ה-ZIP שהורדתם, בחרו מאפיינים,'
            Write-Host '      סמנו "בטל חסימה", ורק אז חלצו אותו מחדש.'
            Write-Host '   2. ודאו שוורד סגור לגמרי.'
            Write-Host '   3. אם זה נמשך — הריצו את "בדיקה" ושלחו לנו את הפלט.'
        }
        else {
            Line
            Write-Host '  סיימנו.' -ForegroundColor Green
            Write-Host ''
            Write-Host '  פתחו את וורד — תופיע לשונית בשם "סַדָּר".'
            Write-Host ''
            Write-Host '  להסרה: לחצו על "הסר".' -ForegroundColor DarkGray
        }

        Wait-Key
        exit $(if ($failed) { 1 } else { 0 })
    }

    'uninstall' {
        Header 'הסרה'

        if (Test-WordRunning) { Wait-Key; exit 1 }

        try {
            & $install -Uninstall
            Write-Host ''
            Write-Host '  הגדרות הכללים שלכם נשמרו ולא נמחקו.' -ForegroundColor DarkGray
        }
        catch {
            Write-Host "  ההסרה נכשלה: $($_.Exception.Message)" -ForegroundColor Red
        }

        Wait-Key
    }

    'check' {
        Header 'בדיקת מצב'

        try { & $install -Status }
        catch { Write-Host "  בדיקת ההתקנה נכשלה: $($_.Exception.Message)" -ForegroundColor Red }

        Write-Host ''
        Write-Host '  המנוע:' -ForegroundColor White

        $exe = Join-Path $root 'bin\sadar.exe'
        if (Test-Path $exe) {
            try { & $exe doctor }
            catch { Write-Host "  בדיקת המנוע נכשלה: $($_.Exception.Message)" -ForegroundColor Red }
        }
        else {
            Write-Host '  לא נמצא bin\sadar.exe' -ForegroundColor Red
        }

        Write-Host ''
        Line
        Write-Host '  אם משהו כאן אינו תקין — שלחו לנו את כל מה שמופיע במסך.' -ForegroundColor Yellow

        Wait-Key
    }
}
