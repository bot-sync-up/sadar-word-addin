#Requires -Version 5
<#
  בדיקת שפיות על ה-DLL של התוסף, בלי להתקין ובלי לפתוח את וורד.
  מוודא שהאסמבלי נטען, שהמחלקה גלויה ל-COM, שממשקי התוסף ממומשים,
  ושה-XML של הרצועה חוקי ומכיל את כל הכפתורים עם פונקציות מתאימות.
#>
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$root = Split-Path -Parent $PSScriptRoot
$dll = Join-Path $root 'bin\Sadar.Addin.dll'

$pass = 0; $fail = 0
function Check($name, $ok, $detail) {
    if ($ok) { $script:pass++; Write-Host "  עבר   $name" -ForegroundColor Green }
    else { $script:fail++; Write-Host "  נכשל  $name  $detail" -ForegroundColor Red }
}

Write-Host 'בדיקת אסמבלי התוסף' -ForegroundColor Cyan
Write-Host ('=' * 60)

$bytes = [System.IO.File]::ReadAllBytes($dll)
$asm = [System.Reflection.Assembly]::Load($bytes)
Check 'האסמבלי נטען' ($null -ne $asm) ''

$connect = $asm.GetType('Sadar.Addin.Connect')
Check 'המחלקה Connect קיימת' ($null -ne $connect) ''

$comVisible = $connect.GetCustomAttributes([System.Runtime.InteropServices.ComVisibleAttribute], $false)
Check 'המחלקה גלויה ל-COM' ($comVisible.Count -gt 0 -and $comVisible[0].Value) ''

$progId = $connect.GetCustomAttributes([System.Runtime.InteropServices.ProgIdAttribute], $false)
Check 'יש ProgId' ($progId.Count -gt 0 -and $progId[0].Value -eq 'Sadar.Connect') ''

$guid = $connect.GetCustomAttributes([System.Runtime.InteropServices.GuidAttribute], $false)
Check 'יש GUID קבוע' ($guid.Count -gt 0) ''

$ifaces = $connect.GetInterfaces() | ForEach-Object { $_.Name }
Check 'ממומש IDTExtensibility2' ($ifaces -contains 'IDTExtensibility2') ''
Check 'ממומש IRibbonExtensibility' ($ifaces -contains 'IRibbonExtensibility') ''

# פונקציות הרישום חייבות להתקיים, אחרת regasm ירשום בלי מפתח התוסף
$reg = $connect.GetMethod('Register', [System.Reflection.BindingFlags]::Public -bor [System.Reflection.BindingFlags]::Static)
$unreg = $connect.GetMethod('Unregister', [System.Reflection.BindingFlags]::Public -bor [System.Reflection.BindingFlags]::Static)
Check 'קיימת פונקציית רישום' ($null -ne $reg) ''
Check 'קיימת פונקציית הסרה' ($null -ne $unreg) ''

# ---- הרצועה ----
$instance = [System.Activator]::CreateInstance($connect)
$xml = $connect.GetMethod('GetCustomUI').Invoke($instance, @([string]'Microsoft.Word.Document'))
Check 'GetCustomUI מחזיר תוכן' (-not [string]::IsNullOrWhiteSpace($xml)) ''

$doc = New-Object System.Xml.XmlDocument
$xmlValid = $true
try { $doc.LoadXml($xml) } catch { $xmlValid = $false; $xmlError = $_.Exception.Message }
Check 'ה-XML של הרצועה חוקי' $xmlValid $xmlError

if ($xmlValid) {
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace('ui', 'http://schemas.microsoft.com/office/2009/07/customui')

    $tab = $doc.SelectSingleNode('//ui:tab', $ns)
    Check 'קיימת לשונית' ($null -ne $tab) ''

    $buttons = $doc.SelectNodes('//ui:button', $ns)
    Check 'קיימים כפתורים' ($buttons.Count -ge 5) "נמצאו $($buttons.Count)"

    # כל onAction חייב להצביע על פונקציה ציבורית שקיימת בפועל,
    # אחרת הכפתור פשוט לא יעשה כלום ווורד לא יתלונן
    $missing = @()
    foreach ($b in $buttons) {
        $action = $b.GetAttribute('onAction')
        if (-not $action) { $missing += "$($b.GetAttribute('id')): אין onAction"; continue }
        $m = $connect.GetMethod($action)
        if ($null -eq $m) { $missing += "$($b.GetAttribute('id')) -> $action" }
    }
    Check 'לכל כפתור יש פונקציה קיימת' ($missing.Count -eq 0) ($missing -join '; ')

    $labels = @()
    foreach ($b in $buttons) { $labels += $b.GetAttribute('label') }
    Write-Host "        כפתורים: $($labels -join ' | ')" -ForegroundColor DarkGray
}

Write-Host ('=' * 60)
if ($fail -eq 0) { Write-Host "עברו $pass, נכשלו 0" -ForegroundColor Green }
else { Write-Host "עברו $pass, נכשלו $fail" -ForegroundColor Red }
exit $fail
