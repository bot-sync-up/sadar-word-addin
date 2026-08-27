<#
  בניית סַדָּר ללא Visual Studio וללא NuGet.

  משתמש ב-csc.exe של .NET Framework 4 שמותקן בכל Windows.
  אין תלות חיצונית אחת בפרויקט — זו החלטה מכוונת: המוצר חינמי,
  והקהילה צריכה להיות מסוגלת לבנות אותו על מחשב רגיל.

  שימוש:
    .\build.ps1              בונה הכל
    .\build.ps1 -Target Cli  בונה רק את כלי שורת הפקודה
    .\build.ps1 -Clean       מנקה פלטים
#>
param(
    [ValidateSet('All', 'Core', 'Cli', 'Addin')]
    [string]$Target = 'All',
    [switch]$Clean
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$src = Join-Path $root 'src'
$out = Join-Path $root 'bin'

$fx = "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319"
$csc = Join-Path $fx 'csc.exe'
if (-not (Test-Path $csc)) { throw "csc.exe לא נמצא ב-$fx" }

if ($Clean) {
    if (Test-Path $out) { Remove-Item $out -Recurse -Force }
    Write-Host 'נוקה.' -ForegroundColor Green
    return
}

New-Item -ItemType Directory -Force -Path $out | Out-Null

# csc קורא קבצים לפי דף הקוד של המערכת אלא אם יש BOM.
# בלי זה כל הטקסט העברי בקוד נהרס בשקט.
function Ensure-Utf8Bom([string]$path) {
    $bytes = [System.IO.File]::ReadAllBytes($path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { return }
    $text = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
    [System.IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding $true))
    Write-Host "  BOM נוסף: $(Split-Path -Leaf $path)" -ForegroundColor DarkGray
}

function Invoke-Csc {
    param(
        [string]$OutputFile,
        [string]$TargetType,
        [string[]]$Sources,
        [string[]]$References = @(),
        [string[]]$LinkedReferences = @()
    )

    foreach ($s in $Sources) { Ensure-Utf8Bom $s }

    $cscArgs = @(
        '/nologo'
        '/nostdlib+'
        '/noconfig'
        "/target:$TargetType"
        "/out:$OutputFile"
        '/optimize+'
        '/warn:4'
        '/langversion:5'
        '/utf8output'
        '/platform:anycpu'
    )

    $baseRefs = @(
        'mscorlib.dll', 'System.dll', 'System.Core.dll', 'System.Xml.dll',
        'System.Xml.Linq.dll', 'System.Web.Extensions.dll', 'System.Security.dll',
        'System.IO.Compression.dll', 'System.IO.Compression.FileSystem.dll'
    )
    foreach ($r in $baseRefs) { $cscArgs += "/reference:$fx\$r" }
    foreach ($r in $References) { $cscArgs += "/reference:$r" }
    foreach ($r in $LinkedReferences) { $cscArgs += "/link:$r" }

    $cscArgs += $Sources

    $rsp = Join-Path $env:TEMP "sadar-csc-$([guid]::NewGuid().ToString('N')).rsp"
    $quoted = $cscArgs | ForEach-Object { if ($_ -match '^/') { $_ } else { "`"$_`"" } }
    [System.IO.File]::WriteAllLines($rsp, $quoted, (New-Object System.Text.UTF8Encoding $true))

    try {
        $output = & $csc "@$rsp" 2>&1
        $exit = $LASTEXITCODE
        $output | ForEach-Object {
            $line = $_.ToString()
            if ($line -match 'error CS') { Write-Host $line -ForegroundColor Red }
            elseif ($line -match 'warning CS') { Write-Host $line -ForegroundColor Yellow }
            elseif ($line.Trim()) { Write-Host $line }
        }
        if ($exit -ne 0) { throw "הבנייה נכשלה: $OutputFile" }
        Write-Host "  נבנה: $(Split-Path -Leaf $OutputFile)" -ForegroundColor Green
    }
    finally {
        Remove-Item $rsp -ErrorAction SilentlyContinue
    }
}

function Get-Sources([string]$dir) {
    Get-ChildItem -Path $dir -Filter '*.cs' -Recurse | ForEach-Object { $_.FullName }
}

# ---------- Sadar.Core ----------
$coreDll = Join-Path $out 'Sadar.Core.dll'
if ($Target -in @('All', 'Core', 'Cli', 'Addin')) {
    Write-Host 'בונה Sadar.Core...' -ForegroundColor Cyan
    Invoke-Csc -OutputFile $coreDll -TargetType 'library' -Sources (Get-Sources (Join-Path $src 'Sadar.Core'))
}

# ---------- Sadar.Cli ----------
if ($Target -in @('All', 'Cli')) {
    Write-Host 'בונה Sadar.Cli...' -ForegroundColor Cyan
    Invoke-Csc -OutputFile (Join-Path $out 'sadar.exe') -TargetType 'exe' `
        -Sources (Get-Sources (Join-Path $src 'Sadar.Cli')) -References @($coreDll)
}

# ---------- Sadar.Addin ----------
if ($Target -in @('All', 'Addin')) {
    $wordPia = Get-ChildItem "$env:WINDIR\assembly\GAC_MSIL\Microsoft.Office.Interop.Word" -Filter '*.dll' -Recurse -ErrorAction SilentlyContinue |
        Select-Object -First 1 -ExpandProperty FullName

    # Microsoft.Office.Core — ממנו נלקח IRibbonExtensibility האמיתי.
    # הכרזה עצמית שלו נראית זהה בקוד אך מייצרת משטח COM שונה,
    # ווורד נכשל בטעינה בלי להסביר דבר.
    $officePia = Get-ChildItem "$env:WINDIR\assembly\GAC_MSIL\office" -Filter '*.dll' -Recurse -ErrorAction SilentlyContinue |
        Select-Object -First 1 -ExpandProperty FullName

    if (-not $wordPia -or -not $officePia) {
        Write-Host 'דילוג על Sadar.Addin: לא נמצאו רכיבי ההדדיות של Word במחשב.' -ForegroundColor Yellow
    }
    else {
        Write-Host 'בונה Sadar.Addin...' -ForegroundColor Cyan

        # הליבה נבנית לתוך אותו DLL ולא כהפניה חיצונית.
        # תוסף COM נטען על ידי וורד, ולכן .NET מחפש תלויות בתיקייה של WINWORD.EXE
        # ולא בתיקייה של התוסף — DLL נפרד פשוט לא יימצא בזמן ריצה.
        $addinSources = @(Get-Sources (Join-Path $src 'Sadar.Core')) + @(Get-Sources (Join-Path $src 'Sadar.Addin'))

        Invoke-Csc -OutputFile (Join-Path $out 'Sadar.Addin.dll') -TargetType 'library' `
            -Sources $addinSources `
            -References @("$fx\System.Windows.Forms.dll", "$fx\System.Drawing.dll") `
            -LinkedReferences @($wordPia, $officePia)
    }
}

# ---------- Sadar.WordTest ----------
# בדיקת אינטגרציה מול וורד. נבנית רק אם התוסף נבנה, כי היא משתמשת במתאם שלו.
if ($Target -in @('All', 'Addin')) {
    $addinDll = Join-Path $out 'Sadar.Addin.dll'
    $wordPia2 = Get-ChildItem "$env:WINDIR\assembly\GAC_MSIL\Microsoft.Office.Interop.Word" -Filter '*.dll' -Recurse -ErrorAction SilentlyContinue |
        Select-Object -First 1 -ExpandProperty FullName

    if ((Test-Path $addinDll) -and $wordPia2) {
        Write-Host 'בונה Sadar.WordTest...' -ForegroundColor Cyan
        Invoke-Csc -OutputFile (Join-Path $out 'wordtest.exe') -TargetType 'exe' `
            -Sources (Get-Sources (Join-Path $src 'Sadar.WordTest')) `
            -References @($addinDll, "$fx\System.Windows.Forms.dll", "$fx\System.Drawing.dll", $wordPia2)
    }
}

Write-Host ''
Write-Host "הפלט: $out" -ForegroundColor Green
Get-ChildItem $out -Filter '*.dll' | ForEach-Object { Write-Host "  $($_.Name)  ($([math]::Round($_.Length/1KB,1)) KB)" }
Get-ChildItem $out -Filter '*.exe' | ForEach-Object { Write-Host "  $($_.Name)  ($([math]::Round($_.Length/1KB,1)) KB)" }
