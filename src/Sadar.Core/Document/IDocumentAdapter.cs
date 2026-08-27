using System;
using System.Collections.Generic;
using Sadar.Core.Model;

namespace Sadar.Core.Document
{
    /// <summary>
    /// הפשטה מעל המסמך. שתי מימושים:
    ///   WordDocumentAdapter   — מסמך פתוח בוורד, דרך ה-Object Model (התוסף)
    ///   OpenXmlDocumentAdapter — קובץ docx על הדיסק (כלי שורת הפקודה והבדיקות)
    ///
    /// כל הלוגיקה של סַדָּר יושבת מעל הממשק הזה בלבד, כך שהליבה נבדקת
    /// בלי וורד כלל — ומתנהגת זהה בשני המקרים.
    /// </summary>
    public interface IDocumentAdapter : IDisposable
    {
        int ParagraphCount { get; }

        /// <summary>קריאת כל הפסקאות בבת אחת. חייבת להיות יעילה — נקראת פעמיים בכל ריצה.</summary>
        List<ParagraphInfo> ReadAll();

        /// <summary>שמות הסגנונות הקיימים במסמך.</summary>
        List<string> GetStyleNames();

        /// <summary>יוצר סגנון אם אינו קיים. מחזיר false אם לא ניתן ליצור.</summary>
        bool EnsureStyleExists(string styleName);

        /// <summary>החלת סגנון על פסקה. פעולת עיצוב טהורה — אינה נוגעת בטקסט.</summary>
        void ApplyStyle(int paragraphIndex, string styleName);

        /// <summary>
        /// כתיבת טקסט מנורמל לפסקה. מותר אך ורק לניקוי רווחים —
        /// המנקה הדטרמיניסטי הוא הצרכן היחיד, ואימות ה-hash אוכף זאת.
        /// </summary>
        void SetParagraphText(int paragraphIndex, string normalizedText);

        /// <summary>מחיקת פסקאות לפי אינדקס. המימוש חייב למחוק מהסוף להתחלה.</summary>
        void DeleteParagraphs(IEnumerable<int> paragraphIndexes);

        /// <summary>פתיחת פעולה אחת שניתנת לביטול יחיד (Ctrl+Z אחד).</summary>
        void BeginBatch(string undoName);

        void EndBatch();

        /// <summary>ביטול מלא של מה שנעשה מאז BeginBatch.</summary>
        void RollbackBatch();

        void Save();
    }
}
