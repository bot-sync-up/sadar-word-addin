#Requires -Version 5
<#
  התקנת תוסף סַדָּר בוורד.

  הרישום נעשה כולו תחת המשתמש הנוכחי (HKCU) ואינו דורש הרשאות מנהל.
  זו החלטה מכוונת: כלי חינמי שדורש "הפעל כמנהל" מאבד חלק ניכר
  מהמשתמשים כבר בהתקנה, ובארגונים רבים פשוט לא ניתן להריץ אותו.

  שימוש:
    .\install.ps1              התקנה
    .\install.ps1 -Uninstall   הסרה
    .\install.ps1 -Status      בדיקת מצב בלבד
#>
param(
    [switch]$Uninstall,
    [switch]$Status,
    [switch]$Machine,
    [string]$DllPath
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$root = Split-Path -Parent $PSScriptRoot
if (-not $DllPath) { $DllPath = Join-Path $root 'bin\Sadar.Addin.dll' }

$progId = 'Sadar.Connect'
$clsid = '{7A3C1F42-9E5B-4D18-B6C4-2E9A5D3F8B71}'
$className = 'Sadar.Addin.Connect'
$friendly = 'סַדָּר — עימוד חכם'
$description = 'החלת סגנונות וניקוי מסמכים, בלי לגעת במלל'

# תוכנית 32 סיביות שקוראת מ-HKCU\Software\Classes מנותבת בשקט
# ל-Wow6432Node. וורד כאן הוא 32 סיביות ו-PowerShell הוא 64, ולכן
# רישום בתצוגה אחת בלבד נראה מוצלח לחלוטין — והתוסף פשוט לא מופיע.
# הפתרון: לרשום בשתי התצוגות.
# ברירת המחדל היא רישום לפי משתמש, שאינו דורש הרשאות.
# חלק מהתקנות Office (בעיקר Click-to-Run) אינן מצליחות לאתחל רכיב
# מנוהל שרשום רק תחת המשתמש, ונכשלות ב-E_FAIL בלי הודעה. במקרה כזה
# -Machine רושם ברמת המכונה. דורש חלון מנהל.
$classesRoots = if ($Machine) {
    @('HKLM:\SOFTWARE\Classes', 'HKLM:\SOFTWARE\Classes\Wow6432Node')
} else {
    @('HKCU:\Software\Classes', 'HKCU:\Software\Classes\Wow6432Node')
}
$clsidKeys = $classesRoots | ForEach-Object { "$_\CLSID\$clsid" }
$progIdKeys = $classesRoots | ForEach-Object { "$_\$progId" }

# מפתח התוסף של Office אינו מנותב — הוא משותף לשתי התצוגות
$addinKey = "HKCU:\Software\Microsoft\Office\Word\Addins\$progId"

# הקטגוריה שמסמנת רכיב .NET מנוהל
$managedCategory = '{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}'

function New-Key([string]$path) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
}

function Set-Default([string]$path, [string]$value) {
    New-Key $path
    Set-ItemProperty -Path $path -Name '(default)' -Value $value
}

function Set-Str([string]$path, [string]$name, [string]$value) {
    New-ItemProperty -Path $path -Name $name -Value $value -PropertyType String -Force | Out-Null
}

# מה שבאמת חוסם התקנה חוזרת הוא DLL נעול, כלומר גרסה קודמת
# של התוסף שכבר טעונה בוורד הרץ.
function Assert-DllWritable {
    if (-not (Test-Path $DllPath)) { return }
    try {
        $fs = [System.IO.File]::Open($DllPath, 'Open', 'ReadWrite', 'None')
        $fs.Close()
    }
    catch {
        throw 'קובץ התוסף נעול — גרסה קודמת טעונה בוורד. סגרו את וורד לגמרי והריצו שוב.'
    }
}

# ---------------- מצב ----------------
if ($Status) {
    Write-Host 'מצב ההתקנה של סַדָּר' -ForegroundColor Cyan
    Write-Host ('=' * 52)

    $items = [ordered]@{
        'קובץ התוסף'       = (Test-Path $DllPath)
        'רישום CLSID (64)' = (Test-Path $clsidKeys[0])
        'רישום CLSID (32)' = (Test-Path $clsidKeys[1])
        'רישום ProgId'     = (Test-Path $progIdKeys[0])
        'רשום כתוסף בוורד' = (Test-Path $addinKey)
    }
    foreach ($k in $items.Keys) {
        $ok = $items[$k]
        $mark = if ($ok) { 'קיים' } else { 'חסר' }
        Write-Host ("  {0,-20} {1}" -f $k, $mark) -ForegroundColor $(if ($ok) { 'Green' } else { 'Red' })
    }

    if (Test-Path $addinKey) {
        $lb = (Get-ItemProperty $addinKey).LoadBehavior
        $meaning = switch ($lb) {
            3 { 'נטען אוטומטית — תקין' }
            2 { 'כבוי. וורד כשל בטעינה בעבר, או שהמשתמש כיבה ידנית' }
            9 { 'נטען לפי דרישה' }
            default { "ערך לא מוכר" }
        }
        Write-Host "  LoadBehavior = $lb  ($meaning)" -ForegroundColor $(if ($lb -eq 3) { 'Green' } else { 'Yellow' })
    }

    $disabled = 'HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems'
    if (Test-Path $disabled) {
        Write-Host '  שימו לב: לוורד יש רשימת פריטים מושבתים.' -ForegroundColor Yellow
        Write-Host '  קובץ > אפשרויות > תוספות > נהל: פריטים מושבתים' -ForegroundColor DarkGray
    }
    return
}

# ---------------- הסרה ----------------
if ($Uninstall) {
    Assert-DllWritable
    Write-Host 'מסיר את סַדָּר...' -ForegroundColor Cyan

    $allScopes = @(
        'HKCU:\Software\Classes', 'HKCU:\Software\Classes\Wow6432Node',
        'HKLM:\SOFTWARE\Classes', 'HKLM:\SOFTWARE\Classes\Wow6432Node'
    )
    $toRemove = @($addinKey)
    foreach ($r in $allScopes) {
        $toRemove += "$r\CLSID\$clsid"
        $toRemove += "$r\$progId"
    }
    foreach ($k in $toRemove) {
        if (Test-Path $k) {
            try { Remove-Item $k -Recurse -Force } catch { }
        }
    }

    $installed = Join-Path $env:LOCALAPPDATA 'Sadar\Sadar.Addin.dll'
    if (Test-Path $installed) {
        try { Remove-Item $installed -Force } catch { Write-Host '  קובץ התוסף נעול ולא נמחק.' -ForegroundColor Yellow }
    }

    Write-Host 'סַדָּר הוסר.' -ForegroundColor Green
    Write-Host 'הגדרות הכללים נשמרו ולא נמחקו: ' -NoNewline
    Write-Host (Join-Path $env:APPDATA 'Sadar') -ForegroundColor DarkGray
    return
}

# ---------------- התקנה ----------------
if (-not (Test-Path $DllPath)) {
    throw "לא נמצא הקובץ $DllPath . הריצו קודם את build.ps1"
}

Assert-DllWritable

if ($Machine) {
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
        ).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    if (-not $isAdmin) {
        throw 'רישום ברמת המכונה דורש חלון PowerShell שנפתח כמנהל.'
    }
    Write-Host 'מצב: רישום ברמת המכונה' -ForegroundColor Yellow
}

# Windows חוסם DLL שהגיע מהאינטרנט. בלי הסרת הסימון הטעינה
# נכשלת בשקט ווורד לא מסביר למה.
$zone = Get-Item $DllPath -Stream Zone.Identifier -ErrorAction SilentlyContinue
if ($zone) {
    Write-Host 'מסיר סימון "קובץ מהאינטרנט"...' -ForegroundColor DarkGray
    Unblock-File -Path $DllPath
}

# מעתיקים את התוסף לתיקיית התקנה קבועה תחת המשתמש.
# הרצה ישירות מתיקיית הבנייה עובדת בפיתוח אבל שבירה בייצור:
# הנתיב משתנה, והמשתמש עלול למחוק או להזיז את התיקייה
# בלי לדעת שהוא שובר בכך את התוסף.
$installDir = Join-Path $env:LOCALAPPDATA 'Sadar'
if (-not (Test-Path $installDir -PathType Container)) { New-Item -ItemType Directory -Path $installDir -Force | Out-Null }

$installedDll = Join-Path $installDir 'Sadar.Addin.dll'
if ((Resolve-Path $DllPath).Path -ne $installedDll) {
    Copy-Item $DllPath $installedDll -Force
    Write-Host "הועתק אל: $installDir" -ForegroundColor DarkGray
}
$DllPath = $installedDll

$resolved = (Resolve-Path $DllPath).Path
$asm = [System.Reflection.Assembly]::LoadFrom($resolved)
$asmName = $asm.FullName
$codeBase = ([System.Uri]$resolved).AbsoluteUri
$asmVersion = $asm.GetName().Version.ToString()
$runtime = $asm.ImageRuntimeVersion

Write-Host "אסמבלי: $asmName" -ForegroundColor DarkGray

# ----- רישום מחלקת ה-COM תחת המשתמש -----
# regasm כותב ל-HKLM ולכן דורש הרשאות מנהל. Windows ממזג את
# HKCU\Software\Classes לתוך HKCR ובודק אותו ראשון, ולכן רישום כאן
# שקול מבחינת וורד — ולא דורש כלום.
Write-Host 'רושם את רכיב ה-COM...' -ForegroundColor Cyan

# mscoree.dll חייב להירשם בנתיב מלא ולפי הביטנס של התצוגה.
# שם קובץ בלבד נפתר לפי סדר חיפוש ה-DLL של התהליך המארח, ותהליך
# עם סדר חיפוש מוקשח פשוט לא ימצא אותו — והטעינה תיכשל בשקט.
$mscoreeByView = @{
    ($classesRoots[0]) = "$env:WINDIR\System32\mscoree.dll"
    ($classesRoots[1]) = "$env:WINDIR\SysWOW64\mscoree.dll"
}

for ($vi = 0; $vi -lt $clsidKeys.Count; $vi++) {
    $ck = $clsidKeys[$vi]
    $mscoree = $mscoreeByView[$classesRoots[$vi]]
    if (-not (Test-Path $mscoree)) { $mscoree = 'mscoree.dll' }

    Set-Default $ck $className

    $inproc = "$ck\InprocServer32"
    Set-Default $inproc $mscoree
    Set-Str $inproc 'ThreadingModel' 'Both'
    Set-Str $inproc 'Class' $className
    Set-Str $inproc 'Assembly' $asmName
    Set-Str $inproc 'RuntimeVersion' $runtime
    Set-Str $inproc 'CodeBase' $codeBase

    # .NET מצפה גם לתת-מפתח לפי מספר הגרסה
    $versioned = "$inproc\$asmVersion"
    New-Key $versioned
    Set-Str $versioned 'Class' $className
    Set-Str $versioned 'Assembly' $asmName
    Set-Str $versioned 'RuntimeVersion' $runtime
    Set-Str $versioned 'CodeBase' $codeBase

    Set-Default "$ck\ProgId" $progId
    New-Key "$ck\Implemented Categories\$managedCategory"
}

foreach ($pk in $progIdKeys) {
    Set-Default $pk $friendly
    Set-Default "$pk\CLSID" $clsid
}

# ----- רישום כתוסף של וורד -----
Write-Host 'רושם את התוסף בוורד...' -ForegroundColor Cyan
New-Key $addinKey
Set-ItemProperty -Path $addinKey -Name 'FriendlyName' -Value $friendly
Set-ItemProperty -Path $addinKey -Name 'Description' -Value $description
Set-ItemProperty -Path $addinKey -Name 'LoadBehavior' -Value 3 -Type DWord
Set-ItemProperty -Path $addinKey -Name 'CommandLineSafe' -Value 0 -Type DWord

Write-Host ''
Write-Host 'סַדָּר הותקן.' -ForegroundColor Green

if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) {
    Write-Host ''
    Write-Host 'וורד פתוח כרגע — התוסף יופיע רק אחרי שתסגרו ותפתחו אותו מחדש.' -ForegroundColor Yellow
}

Write-Host ''
Write-Host 'פתחו את וורד — תופיע לשונית בשם "סַדָּר".'
Write-Host ''
Write-Host 'אם הלשונית אינה מופיעה:' -ForegroundColor DarkGray
Write-Host '  קובץ > אפשרויות > תוספות > נהל: תוספות COM > מעבר' -ForegroundColor DarkGray
Write-Host '  וודאו שהתיבה ליד "סַדָּר" מסומנת.' -ForegroundColor DarkGray
