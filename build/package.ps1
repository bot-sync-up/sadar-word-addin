#Requires -Version 5
<#
  אורז חבילת הפצה לבודק חיצוני.

  החבילה עצמאית: DLL אחד, סקריפטי התקנה ואבחון, וכלי שורת הפקודה.
  אין צורך ב-Visual Studio, ב-NuGet או בהרשאות מנהל אצל הבודק.
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

# ---- הוראות לבודק ----
$instructions = @'
סַדָּר — חבילת בדיקה
====================

תודה שאתם בודקים. הבדיקה אמורה לקחת חמש דקות.

לפני הכל
---------
לחיצה ימנית על קובץ ה-ZIP  >  מאפיינים  >  סמנו "בטל חסימה"  >  אישור
רק אחר כך לחלצו אותו. בלי זה Windows חוסם את הקובץ ושום דבר לא יעבוד.

שלב 1 — התקנה
--------------
סגרו את וורד לגמרי, ולחצו לחיצה כפולה על:

    התקן

זה הכל. אין צורך בהרשאות מנהל ואין צורך להקליד שום פקודה.

שלב 2 — הבדיקה
---------------
פתחו את וורד. האם מופיעה לשונית בשם "סַדָּר"?

    כן  ->  מצוין. עברו לשלב 3.
    לא  ->  לחצו לחיצה כפולה על "בדיקה" ושלחו לנו צילום מסך.
            זה בדיוק המידע שאנחנו צריכים.

שלב 3 — ניסיון אמיתי
---------------------
פתחו מסמך שאתם עובדים עליו (עותק, לא המקור), ולחצו "סדר את המסמך".

התוסף יציג טבלת הצעה. שום דבר לא מוחל עד שתאשרו.
עברו על ההצעה, ואם היא נראית טוב — לחצו "החל את המסומנים".

בסוף תופיע הודעה שמאמתת שהמלל לא השתנה.
לביטול מלא: Ctrl+Z אחד.

לתשומת לבכם
------------
* התוסף אינו נוגע במלל. לפני כל שינוי נלקחת טביעת אצבע של הטקסט
  ואחריו היא נבדקת שוב; אם משהו השתנה הפעולה מתבטלת מאליה.
* לפני כל סידור נשמר גיבוי ליד הקובץ.
* לזיהוי הכותרות נדרש Claude Code מותקן ומחובר:
      npm install -g @anthropic-ai/claude-code
      claude      ואז      /login
  לבדיקה:  bin\sadar.exe doctor
* "ניקוי בלבד" עובד גם בלי זה — הוא רץ מקומית לגמרי, בלי אינטרנט.

להסרה
------
לחיצה כפולה על "הסר". הכללים שלכם נשמרים.

מה שהכי יעזור לנו לדעת
-----------------------
1. האם הלשונית הופיעה?
2. אם כן — האם הזיהוי היה נכון? מה הוא פספס?
3. איזו גרסת Office יש לכם? (קובץ > חשבון > אודות Word)
'@

$instructions | Out-File -FilePath (Join-Path $stage 'קרא אותי.txt') -Encoding utf8

Compress-Archive -Path $stage -DestinationPath $zip -Force

$size = [math]::Round((Get-Item $zip).Length / 1KB, 1)
Write-Host "נארז: $zip  ($size KB)" -ForegroundColor Green
Get-ChildItem $stage -Recurse -File | ForEach-Object {
    Write-Host ('  ' + $_.FullName.Substring($stage.Length + 1))
}
