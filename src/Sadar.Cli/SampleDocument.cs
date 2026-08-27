using System.Collections.Generic;

namespace Sadar.Cli
{
    /// <summary>
    /// מסמך בדיקה שמדמה ספר שאלות ותשובות כפי שהוא מגיע ממקליד:
    /// בלי סגנונות, עם רווחים כפולים, עם שורות ריקות עודפות,
    /// ועם מלכודות שנועדו לתפוס באגים — ראשי תיבות בגרשיים,
    /// פסקה מודגשת שאינה כותרת, ומראי מקומות.
    ///
    /// אורכי הפסקאות מכוונים לפרופורציות של ספר אמיתי (מאות תווים לפסקה),
    /// כי מדידת נפח על פסקאות קצרות מדי נותנת תוצאה מטעה.
    /// </summary>
    public static class SampleDocument
    {
        private const string Tail =
            " וכבר עמדו בזה גדולי האחרונים והאריכו בזה טובא, ומה שנראה לעניות דעתי " +
            "ביישוב הדברים הוא על פי מה שיסדו הראשונים דיש לחלק בין עיקר המצוה לבין " +
            "הכשר מצוה, ולפי זה אתי שפיר כל מה שהקשו. ועיין היטב בדברי השו\"ת שהביא " +
            "ראיות לזה מכמה מקומות בש\"ס ובראשונים, ואין כאן מקום להאריך יותר. " +
            "והנה גם מדברי הפוסקים האחרונים משמע כדברינו, ורק שהם לא פירשו כן להדיא, " +
            "ומכל מקום העיקר להלכה כמו שכתבנו, וכן הורו רבותינו שליט\"א למעשה.";

        public static void Create(string path)
        {
            Create(path, 1);
        }

        /// <summary>מייצר ספר בגודל מבוקש למדידת נפח השלד על היקף אמיתי.</summary>
        public static void Create(string path, int repeat)
        {
            var p = new List<DocxBuilder.Para>();

            p.Add(DocxBuilder.Para.Of("בס\"ד", false, 12, "center"));
            p.Add(DocxBuilder.Para.Of(""));

            AddSiman(p, "סימן א",
                "בדין נטילת ידים שחרית",
                new[]
                {
                    "כתב הרמב\"ם בהלכות תפילה פרק ד' הלכה ב' וז\"ל, כל הנוטל ידיו צריך שיברך  על נטילת ידים, ומשמע מדבריו דהברכה היא חלק מעצם המצוה ולא רק הכשר בעלמא." + Tail,
                    "ויש לעיין בדברי הטור סימן ד' שכתב דבר זה בלשון אחרת, ומשמע מיניה דהברכה  היא רשות, ולכאורה דבריו סותרים למה שכתב במקום אחר." + Tail,
                    "ולענ\"ד נראה ליישב, דהנה יש לחלק בין נטילה שחרית לנטילה לסעודה, ובזה יתיישבו כל הקושיות שהקשו הראשונים." + Tail
                });

            AddSiman(p, "סימן ב",
                "בענין הפסק בברכת המזון",
                new[]
                {
                    "שאלה: מי שהתחיל לברך ברכת המזון ובאמצע נזכר שלא נטל מים אחרונים, מה דינו.",
                    "תשובה: הנה מבואר בשו\"ע או\"ח סימן קפ\"א דמים אחרונים חובה, ומ\"מ אם כבר התחיל  בברכה אין לו להפסיק." + Tail,
                    "וטעם הדבר, דהפסק בברכת המזון חמור טפי מהפסק בשאר ברכות, וכמו שכתב המשנה ברורה שם ס\"ק כ\"ב ." + Tail
                });

            // מלכודת: פסקה מודגשת שאינה כותרת אלא הדגשה בתוך גוף הטקסט
            p.Add(DocxBuilder.Para.Of(
                "והעיקר למעשה כמו שכתבנו לעיל, וכן הורה לי מורי ורבי שליט\"א.",
                true, 12, "both"));
            p.Add(DocxBuilder.Para.Of(""));

            AddSiman(p, "סימן ג",
                "בדיני קריאת שמע על המטה",
                new[]
                {
                    "כתבו הפוסקים דקריאת שמע שעל המטה אינה חובה גמורה אלא מנהג טוב, ומ\"מ נהגו בה כל ישראל." + Tail,
                    "ויש שכתבו דהיא תקנת חכמים ממש, ועיין בזה באריכות בספרי האחרונים ובמה שכתבו בזה בשם הגאונים." + Tail
                });

            // שורות ריקות עודפות בסוף המסמך
            p.Add(DocxBuilder.Para.Of(""));
            p.Add(DocxBuilder.Para.Of(""));
            p.Add(DocxBuilder.Para.Of(""));

            if (repeat > 1)
            {
                var big = new List<DocxBuilder.Para>();
                for (int r = 0; r < repeat; r++) big.AddRange(p);
                p = big;
            }

            DocxBuilder.Create(path, p);
        }

        private static void AddSiman(List<DocxBuilder.Para> p, string siman, string subject, string[] body)
        {
            p.Add(DocxBuilder.Para.Of(siman, true, 16, "center"));
            p.Add(DocxBuilder.Para.Of(subject, true, 14, "center"));
            p.Add(DocxBuilder.Para.Of(""));

            foreach (string b in body)
            {
                p.Add(DocxBuilder.Para.Of(b, false, 12, "both"));
                p.Add(DocxBuilder.Para.Of(""));
            }

            p.Add(DocxBuilder.Para.Of(""));
        }
    }
}
