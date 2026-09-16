#Requires -Version 5
<#
  אורז את חבילת ההפצה שמתפרסמת ב-GitHub Releases.

  החבילה עצמאית: DLL אחד, קובצי לחיצה כפולה, סקריפטי התקנה ואבחון,
  וכלי שורת הפקודה. אין צורך ב-Visual Studio, ב-NuGet או בהרשאות מנהל.
#>
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$root = Split-Path -Parent $PSScriptRoot
$bin = Join-Path $root 'bin'
$stage = Join-Path $root 'dist\sadar'
$zip = Join-Path $root 'dist\sadar.zip'

foreach ($f in @('Sadar.Addin.dll', 'sadar.exe', 'Sadar.Core.dll')) {
    if (-not (Test-Path (Join-Path $bin $f))) { throw "חסר $f — הריצו קודם build.ps1" }
}

if (Test-Path (Split-Path $stage)) { Remove-Item (Split-Path $stage) -Recurse -Force }
New-Item -ItemType Directory -Path "$stage\bin", "$stage\build" -Force | Out-Null

Copy-Item (Join-Path $bin 'Sadar.Addin.dll') "$stage\bin\"
Copy-Item (Join-Path $bin 'sadar.exe') "$stage\bin\"
Copy-Item (Join-Path $bin 'Sadar.Core.dll') "$stage\bin\"
if (Test-Path (Join-Path $bin 'wordtest.exe')) { Copy-Item (Join-Path $bin 'wordtest.exe') "$stage\bin\" }

foreach ($s in @('install.ps1', 'launcher.ps1', 'diagnose.ps1', 'live-check.ps1',
                 'clear-disabled.ps1', 'force-connect.ps1')) {
    Copy-Item (Join-Path $PSScriptRoot $s) "$stage\build\"
}

# קובצי הלחיצה הכפולה — זה מה שרוב המשתמשים יראו וישתמשו בו.
# בלעדיהם ההתקנה דורשת הקלדת פקודה, וזה חסם אמיתי בקהל הזה.
foreach ($c in (Get-ChildItem -Path $root -Filter '*.cmd' -File)) {
    Copy-Item $c.FullName $stage
}

Copy-Item (Join-Path $root 'README.md') $stage
Copy-Item (Join-Path $root 'LICENSE') $stage

# ---- הוראות למשתמש ----
# אותו מדריך שמתפרסם בפורום ובגיטהב, כדי שמי שמוריד יקבל בדיוק את אותו טקסט
$guide = Join-Path $root 'docs\מדריך-התקנה.txt'
if (-not (Test-Path $guide)) { throw "חסר המדריך: $guide" }
Copy-Item $guide (Join-Path $stage 'קרא אותי.txt')

# אותו מדריך בפורמט להדפסה — מי שמעדיף לקרוא מסודר או להדפיס
$guidePdf = Join-Path $root 'docs\מדריך-התקנה.pdf'
if (-not (Test-Path $guidePdf)) { throw "חסר המדריך בפורמט PDF: $guidePdf" }
Copy-Item $guidePdf (Join-Path $stage 'מדריך התקנה.pdf')

Compress-Archive -Path $stage -DestinationPath $zip -Force

$size = [math]::Round((Get-Item $zip).Length / 1KB, 1)
Write-Host "נארז: $zip  ($size KB)" -ForegroundColor Green
Get-ChildItem $stage -Recurse -File | ForEach-Object {
    Write-Host ('  ' + $_.FullName.Substring($stage.Length + 1))
}
