#Requires -Version 5
<#
  הרצת הדגמה מקצה לקצה: יוצר ספר בגודל אמיתי, מודד את נפח השלד,
  מנתח, מחיל ומאמת. משמש גם כבדיקת קבלה ידנית.
#>
param(
    [int]$Repeat = 24,
    [string]$Engine = 'mock'
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$root = Split-Path -Parent $PSScriptRoot
$exe = Join-Path $root 'bin\sadar.exe'
$samples = Join-Path $root 'samples'
New-Item -ItemType Directory -Force -Path $samples | Out-Null

$book = Join-Path $samples 'sefer-test.docx'
$out = Join-Path $samples 'sefer-test-mesudar.docx'

Write-Host '=== 1. יצירת ספר בדיקה ===' -ForegroundColor Cyan
& $exe sample $book --repeat $Repeat

Write-Host ''
Write-Host '=== 2. מדידת נפח השלד ===' -ForegroundColor Cyan
& $exe scan $book 2>&1 1>$null | Select-Object -Last 5

Write-Host ''
Write-Host '=== 3. ניתוח והחלה ===' -ForegroundColor Cyan
& $exe apply $book --engine $Engine --out $out

Write-Host ''
Write-Host '=== 4. אימות עצמאי של הקובץ שנשמר ===' -ForegroundColor Cyan
& $exe scan $out 2>&1 1>$null | Select-Object -Last 3
