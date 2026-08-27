using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

[assembly: AssemblyTitle("סַדָּר — עימוד חכם לוורד")]
[assembly: AssemblyDescription("החלת סגנונות וניקוי מסמכים, בלי לגעת במלל")]
[assembly: AssemblyProduct("Sadar")]
[assembly: AssemblyCompany("Sync Up")]

// הגרסה נכתבת לרישום ה-COM ברישום המערכת ולכן היא חייבת להיות מפורשת
// ויציבה. שינוי שלה מחייב רישום מחדש.
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

// בלי התכונה הזו מעטפת ה-CLR שנטענת לתוך וורד אינה יודעת איזו גרסת
// זמן ריצה לאתחל, והאתחול נכשל עם 0x80004005 — כשל שקט לחלוטין
// שנראה בדיוק כמו "התוסף לא מותקן".
[assembly: TargetFramework(".NETFramework,Version=v4.0", FrameworkDisplayName = ".NET Framework 4")]

// המחלקות שנחשפות ל-COM מסומנות אחת-אחת ב-ComVisible משלהן
[assembly: ComVisible(true)]
