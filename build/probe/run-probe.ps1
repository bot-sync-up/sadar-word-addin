#Requires -Version 5
<#
  בונה, רושם ומריץ את תוסף הבדיקה המינימלי.
  מטרתו יחידה: להכריע אם וורד מסוגל בכלל לטעון תוסף COM מנוהל במחשב הזה.
#>
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$here = $PSScriptRoot
$fx = "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319"
$csc = Join-Path $fx 'csc.exe'
$dll = Join-Path $here 'SadarProbe.dll'
$log = Join-Path $env:TEMP 'sadar-probe.log'

$progId = 'SadarProbe.Connect'
$clsid = '{3B1C7E90-5A44-4D2B-9F71-1C4E8A6D2F03}'
$addinKey = "HKCU:\Software\Microsoft\Office\Word\Addins\$progId"

# ---- בנייה ----
Write-Host 'בונה את תוסף הבדיקה...' -ForegroundColor Cyan
$src = Join-Path $here 'Probe.cs'
$bytes = [System.IO.File]::ReadAllBytes($src)
if (-not ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)) {
    $t = [System.IO.File]::ReadAllText($src, [System.Text.Encoding]::UTF8)
    [System.IO.File]::WriteAllText($src, $t, (New-Object System.Text.UTF8Encoding $true))
}

& $csc /nologo /target:library "/out:$dll" /platform:anycpu $src
if ($LASTEXITCODE -ne 0) { throw 'הבנייה נכשלה' }
Write-Host '  נבנה.' -ForegroundColor Green

# ---- רישום ----
Write-Host 'רושם...' -ForegroundColor Cyan
$asm = [System.Reflection.Assembly]::LoadFrom($dll)
$asmName = $asm.FullName
$codeBase = ([System.Uri]$dll).AbsoluteUri
$runtime = $asm.ImageRuntimeVersion
$ver = $asm.GetName().Version.ToString()

foreach ($pair in @(
    @{ Root = 'HKCU:\Software\Classes'; Core = "$env:WINDIR\System32\mscoree.dll" },
    @{ Root = 'HKCU:\Software\Classes\Wow6432Node'; Core = "$env:WINDIR\SysWOW64\mscoree.dll" })) {

    $ck = "$($pair.Root)\CLSID\$clsid"
    New-Item -Path "$ck\InprocServer32\$ver" -Force | Out-Null
    Set-ItemProperty "$ck" -Name '(default)' -Value 'Sadar.Probe.Connect'

    foreach ($k in @("$ck\InprocServer32", "$ck\InprocServer32\$ver")) {
        New-ItemProperty $k -Name 'Class' -Value 'Sadar.Probe.Connect' -PropertyType String -Force | Out-Null
        New-ItemProperty $k -Name 'Assembly' -Value $asmName -PropertyType String -Force | Out-Null
        New-ItemProperty $k -Name 'RuntimeVersion' -Value $runtime -PropertyType String -Force | Out-Null
        New-ItemProperty $k -Name 'CodeBase' -Value $codeBase -PropertyType String -Force | Out-Null
    }
    Set-ItemProperty "$ck\InprocServer32" -Name '(default)' -Value $pair.Core
    New-ItemProperty "$ck\InprocServer32" -Name 'ThreadingModel' -Value 'Both' -PropertyType String -Force | Out-Null

    New-Item -Path "$($pair.Root)\$progId\CLSID" -Force | Out-Null
    Set-ItemProperty "$($pair.Root)\$progId\CLSID" -Name '(default)' -Value $clsid
}

New-Item -Path $addinKey -Force | Out-Null
Set-ItemProperty $addinKey -Name 'FriendlyName' -Value 'Sadar Probe'
Set-ItemProperty $addinKey -Name 'Description' -Value 'minimal load probe'
Set-ItemProperty $addinKey -Name 'LoadBehavior' -Value 3 -Type DWord
Write-Host '  נרשם.' -ForegroundColor Green

# ---- הרצה ----
if (Test-Path $log) { Remove-Item $log -Force }

foreach ($p in @(Get-Process WINWORD -ErrorAction SilentlyContinue)) {
    if ($p.MainWindowHandle -ne 0) { Write-Host 'סגרו את וורד והריצו שוב.' -ForegroundColor Yellow; exit 3 }
}
Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { try { $_.Kill() } catch { } }
Start-Sleep -Seconds 2

$winword = @(
    "$env:ProgramFiles\Microsoft Office\root\Office16\WINWORD.EXE",
    "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\WINWORD.EXE"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

Write-Host 'פותח את וורד...' -ForegroundColor Cyan
Start-Process -FilePath $winword
Start-Sleep -Seconds 14

$word = $null
for ($i = 0; $i -lt 8; $i++) {
    try { $word = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application'); break }
    catch { Start-Sleep -Seconds 2 }
}

if ($word) {
    foreach ($a in $word.COMAddIns) {
        if ($a.ProgId -eq $progId) {
            Write-Host "תוסף הבדיקה ברשימה. Connect=$($a.Connect)" -ForegroundColor $(if ($a.Connect) { 'Green' } else { 'Red' })
        }
    }
    try { if ($word.Documents.Count -eq 0) { $word.Quit() } } catch { }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch { }
}

Start-Sleep -Seconds 2
Write-Host ''
if (Test-Path $log) {
    Write-Host 'תוסף הבדיקה נטען. יומן:' -ForegroundColor Green
    Get-Content $log | ForEach-Object { Write-Host "  $_" }
    Write-Host ''
    Write-Host 'מסקנה: וורד כן מסוגל לטעון תוסף COM מנוהל. הבעיה בתלויות של סַדָּר.' -ForegroundColor Cyan
}
else {
    Write-Host 'תוסף הבדיקה לא נטען כלל.' -ForegroundColor Red
    Write-Host 'מסקנה: וורד במחשב הזה אינו טוען תוספות COM מנוהלות. הבעיה סביבתית ולא בקוד.' -ForegroundColor Cyan
}

Write-Host ''
Write-Host 'LoadBehavior של תוסף הבדיקה: ' -NoNewline
Write-Host (Get-ItemProperty $addinKey).LoadBehavior
