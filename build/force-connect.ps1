#Requires -Version 5
<#
  מנסה לחבר את התוסף בכוח כדי לחלץ מוורד את הודעת השגיאה המלאה.

  כשוורד נכשל בטעינת תוסף הוא לא אומר דבר וסתם מוריד את LoadBehavior ל-2.
  הצבה ידנית של Connect=$true מאלצת אותו לנסות שוב ולזרוק חריגה עם פרטים.
#>
$ErrorActionPreference = 'Continue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$progId = 'Sadar.Connect'
$addinKey = "HKCU:\Software\Microsoft\Office\Word\Addins\$progId"

Write-Host 'ניסיון חיבור מאולץ' -ForegroundColor Cyan
Write-Host ('=' * 62)

# תהליכי וורד תלויים מלפני הרישום יזהמו את הבדיקה
foreach ($p in @(Get-Process WINWORD -ErrorAction SilentlyContinue)) {
    if ($p.MainWindowHandle -ne 0) {
        Write-Host "וורד פתוח עם חלון (PID $($p.Id)). סגרו אותו והריצו שוב." -ForegroundColor Yellow
        exit 3
    }
}
Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { try { $_.Kill() } catch { } }
Start-Sleep -Seconds 2

Set-ItemProperty $addinKey -Name 'LoadBehavior' -Value 3 -Type DWord

$winword = @(
    "$env:ProgramFiles\Microsoft Office\root\Office16\WINWORD.EXE",
    "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\WINWORD.EXE"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

Write-Host 'פותח את וורד...' -ForegroundColor DarkGray
Start-Process -FilePath $winword
Start-Sleep -Seconds 12

$word = $null
for ($i = 0; $i -lt 10; $i++) {
    try { $word = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application'); break }
    catch { Start-Sleep -Seconds 2 }
}
if (-not $word) { Write-Host 'לא ניתן להתחבר לוורד.' -ForegroundColor Red; exit 1 }

try {
    $addin = $null
    foreach ($a in $word.COMAddIns) { if ($a.ProgId -eq $progId) { $addin = $a } }

    if (-not $addin) {
        Write-Host 'התוסף אינו ברשימה.' -ForegroundColor Red
        exit 1
    }

    Write-Host "נמצא: $($addin.Description)" -ForegroundColor Green
    Write-Host "Connect לפני: $($addin.Connect)"
    Write-Host ''
    Write-Host 'מנסה לחבר...' -ForegroundColor Cyan

    try {
        $addin.Connect = $true
        Start-Sleep -Seconds 3
        Write-Host "Connect אחרי: $($addin.Connect)" -ForegroundColor $(if ($addin.Connect) { 'Green' } else { 'Red' })

        if ($addin.Connect) {
            Write-Host ''
            Write-Host 'התוסף נטען.' -ForegroundColor Green

            # אם הלשונית נבנתה, וורד מכיר את מזהי הפקדים שלה
            try {
                $vis = $word.CommandBars.GetVisibleMso('SadarAnalyze')
                Write-Host "פקד הלשונית נגיש: $vis" -ForegroundColor Green
            }
            catch {
                Write-Host "בדיקת הלשונית לא נתמכת: $($_.Exception.Message)" -ForegroundColor DarkGray
            }
        }
    }
    catch {
        Write-Host ''
        Write-Host 'החיבור נכשל. הודעת וורד:' -ForegroundColor Red
        Write-Host "  $($_.Exception.Message)" -ForegroundColor Yellow
        if ($_.Exception.InnerException) {
            Write-Host "  פנימי: $($_.Exception.InnerException.Message)" -ForegroundColor Yellow
        }
        $hr = $_.Exception.HResult
        Write-Host ("  HRESULT: 0x{0:X8}" -f $hr) -ForegroundColor Yellow
    }
}
finally {
    try { if ($word.Documents.Count -eq 0) { $word.Quit() } } catch { }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch { }
}

Write-Host ''
Write-Host 'LoadBehavior אחרי הניסיון: ' -NoNewline
Write-Host (Get-ItemProperty $addinKey).LoadBehavior
