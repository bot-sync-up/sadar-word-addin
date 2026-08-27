using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using Sadar.Core.Rules;

namespace Sadar.Core.Engine
{
    /// <summary>
    /// בסיס משותף למנועים שעובדים מול API עם מפתח.
    ///
    /// עובד ב-HTTP גולמי ולא ב-SDK של אף ספק: הפרויקט נבנה עם csc.exe בלבד,
    /// בלי NuGet ובלי תלות חיצונית אחת, כדי שכל אחד בקהילה יוכל לבנות אותו
    /// על מחשב רגיל. שלוש בקשות POST פשוטות אינן מצדיקות לוותר על זה.
    /// </summary>
    public abstract class HttpEngineBase : IDecisionEngine
    {
        protected readonly EngineSettings Settings;
        protected readonly EngineInfo Info;

        protected HttpEngineBase(EngineSettings settings, EngineInfo info)
        {
            Settings = settings;
            Info = info;
            EnsureModernTls();
        }

        public string Name { get { return Info.DisplayName; } }

        /// <summary>
        /// .NET Framework 4 מגיע עם TLS 1.0 כברירת מחדל, וכל שרתי ה-API
        /// דוחים אותו היום. בלי השורה הזו כל קריאה נכשלת ב"החיבור נסגר",
        /// שגיאה שנראית בדיוק כמו תקלת רשת ושולחת לחפש במקום הלא נכון.
        /// </summary>
        private static bool _tlsReady;
        private static readonly object TlsLock = new object();

        protected static void EnsureModernTls()
        {
            if (_tlsReady) return;
            lock (TlsLock)
            {
                if (_tlsReady) return;
                try
                {
                    // TLS 1.2 בלבד.
                    //
                    // הוספת TLS 1.3 (12288) נראית כמו שיפור אבל היא מלכודת:
                    // .NET Framework 4.8 אינו תומך בו במלואו, ומול שרת שמציע
                    // אותו לחיצת היד נתקעת עד ל-timeout — שנראה בדיוק כמו
                    // שרת שאינו זמין. זה בדיוק מה שקרה מול נקודת הקצה של גוגל,
                    // בזמן שספקים שהסתפקו ב-1.2 עבדו כרגיל.
                    ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;

                    // ‏.NET שולח Expect: 100-continue ומחכה לתשובת ביניים.
                    // חלק מהשערים אינם עונים עליה, וההמתנה נראית כתקלת רשת.
                    ServicePointManager.Expect100Continue = false;
                }
                catch { }
                _tlsReady = true;
            }
        }

        // ---------- מה שכל מנוע ממלא ----------

        // בניית הבקשה ופענוח התשובה חשופים בכוונה: הם חוזה החיווט מול הספק,
        // וזה מה שמאפשר לבדוק את הפורמט של כל ספק בלי גישה לרשת ובלי מפתח.
        public abstract string Endpoint { get; }
        protected abstract void AddHeaders(HttpWebRequest request, string apiKey);
        public abstract string BuildRequestBody(string systemPrompt, string userPrompt, string model);
        public abstract string ExtractText(string responseJson);

        // ---------- זמינות ----------

        public virtual string CheckAvailability()
        {
            string key = Settings.CurrentKey;
            if (string.IsNullOrEmpty(key))
                return "לא הוגדר מפתח API עבור " + Info.DisplayName + ". " + Info.HowToConnect;

            try
            {
                string reply = Post("החזר בדיוק את המילה: תקין", "אתה עונה בקצרה.", 40);
                if (string.IsNullOrEmpty(reply))
                    return "השירות החזיר תשובה ריקה. בדקו את המפתח ואת החיבור.";
                return null;
            }
            catch (EngineException ex)
            {
                return ex.Message;
            }
            catch (Exception ex)
            {
                return "בדיקת החיבור נכשלה: " + ex.Message;
            }
        }

        // ---------- הרצה ----------

        public string Decide(string skeletonJson, RulesFile rules)
        {
            string raw = Post(skeletonJson, ClaudeCodeEngine.SystemPrompt, Settings.TimeoutSeconds);

            string json = Json.Json.ExtractObject(raw);
            if (json == null)
                throw new EngineException("התשובה מ" + Info.DisplayName + " אינה מכילה JSON תקין.");

            return json;
        }

        protected string Post(string userPrompt, string systemPrompt, int timeoutSeconds)
        {
            string key = Settings.CurrentKey;
            if (Info.NeedsApiKey && string.IsNullOrEmpty(key))
                throw new EngineException("לא הוגדר מפתח API עבור " + Info.DisplayName + ".");

            string body = BuildRequestBody(systemPrompt, userPrompt, Settings.EffectiveModel);

            var request = (HttpWebRequest)WebRequest.Create(Endpoint);
            request.Method = "POST";
            request.ContentType = "application/json";
            request.Timeout = timeoutSeconds * 1000;
            request.ReadWriteTimeout = timeoutSeconds * 1000;
            request.UserAgent = "Sadar/1.0";
            AddHeaders(request, key);

            byte[] payload = Encoding.UTF8.GetBytes(body);
            request.ContentLength = payload.Length;

            try
            {
                using (var stream = request.GetRequestStream())
                    stream.Write(payload, 0, payload.Length);

                using (var response = (HttpWebResponse)request.GetResponse())
                using (var reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                {
                    string text = reader.ReadToEnd();
                    return ExtractText(text);
                }
            }
            catch (WebException ex)
            {
                throw Translate(ex);
            }
        }

        /// <summary>
        /// הופך שגיאת רשת להודעה שאפשר לפעול לפיה.
        ///
        /// ההבחנה בין חסימת סינון לבין שגיאת מפתח חשובה: הן דורשות
        /// פעולה אחרת לגמרי מהמשתמש, ושתיהן מגיעות כ"החיבור נכשל".
        /// </summary>
        protected EngineException Translate(WebException ex)
        {
            string detail = "";
            var response = ex.Response as HttpWebResponse;

            if (response != null)
            {
                try
                {
                    using (var reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                        detail = reader.ReadToEnd();
                }
                catch { }

                int code = (int)response.StatusCode;
                string friendly = FriendlyMessage(code, detail);
                if (friendly != null) throw new EngineException(friendly);

                return new EngineException(string.Format(
                    "{0} החזיר שגיאה {1}. {2}", Info.DisplayName, code, Trim(detail, 300)));
            }

            if (ex.Status == WebExceptionStatus.TrustFailure ||
                ex.Status == WebExceptionStatus.SecureChannelFailure ||
                ClaudeCodeEngine.LooksLikeFilter(ex.Message))
            {
                return new EngineException(
                    "החיבור נחסם על ידי הסינון ברשת. נסו להפעיל 'מצב מסונן' בהגדרות, " +
                    "או מנוע אחר שאינו חסום.", true);
            }

            if (ex.Status == WebExceptionStatus.Timeout)
                return new EngineException(
                    Info.Vendor + " לא השיב בזמן. המסמך לא נגע.");

            return new EngineException("החיבור ל" + Info.Vendor + " נכשל (" + ex.Status + "): " + ex.Message);
        }

        private string FriendlyMessage(int code, string detail)
        {
            // גוגל מחזירה 400 ולא 401 על מפתח שגוי. בלי הבדיקה הזו
            // המשתמש מקבל JSON גולמי במקום משפט שאפשר לפעול לפיו.
            if (code == 400 && LooksLikeBadKey(detail))
                return "המפתח של " + Info.Vendor + " נדחה. בדקו שהוא הועתק במלואו ושהוא פעיל.";

            switch (code)
            {
                case 400:
                    return "הבקשה נדחתה על ידי " + Info.Vendor + ". " +
                           "לרוב זה מודל שאינו קיים או שאין לו גישה: \"" +
                           Settings.EffectiveModel + "\".";
                case 401:
                case 403:
                    return "המפתח של " + Info.Vendor + " נדחה. בדקו שהוא הועתק במלואו ושהוא פעיל.";
                case 404:
                    return "המודל \"" + Settings.EffectiveModel + "\" אינו קיים אצל " + Info.Vendor +
                           ", או שאין למפתח גישה אליו.";
                case 429:
                    return Info.Vendor + " הגביל את קצב הבקשות. המתינו מעט ונסו שוב.";
                case 402:
                    return "אין יתרה בחשבון של " + Info.Vendor + ".";
                case 413:
                    return "הבקשה גדולה מדי. הקטינו את גודל החלון בהגדרות הכללים.";
                case 500:
                case 502:
                case 503:
                case 529:
                    return "השירות של " + Info.Vendor + " אינו זמין כרגע. נסו שוב בעוד כמה דקות.";
                default:
                    return null;
            }
        }

        /// <summary>מזהה תשובת "מפתח לא תקין" שהגיעה עם קוד סטטוס שאינו 401.</summary>
        private static bool LooksLikeBadKey(string detail)
        {
            if (string.IsNullOrEmpty(detail)) return false;
            string d = detail.ToLowerInvariant();
            return d.Contains("api_key_invalid") ||
                   d.Contains("api key not valid") ||
                   d.Contains("invalid api key") ||
                   d.Contains("invalid_api_key") ||
                   d.Contains("incorrect api key") ||
                   d.Contains("authentication_error");
        }

        protected static string Trim(string s, int max)
        {
            if (string.IsNullOrEmpty(s)) return "";
            s = s.Trim();
            return s.Length <= max ? s : s.Substring(0, max) + "...";
        }

        /// <summary>עוזר: שליפת מחרוזת מנתיב מקונן בתשובה.</summary>
        protected static string DigString(object node, params object[] path)
        {
            foreach (object step in path)
            {
                if (node == null) return null;

                if (step is int)
                {
                    var list = node as IList<object>;
                    if (list == null)
                    {
                        var arr = node as System.Collections.IList;
                        if (arr == null || (int)step >= arr.Count) return null;
                        node = arr[(int)step];
                        continue;
                    }
                    if ((int)step >= list.Count) return null;
                    node = list[(int)step];
                }
                else
                {
                    var map = node as System.Collections.IDictionary;
                    if (map == null) return null;
                    string k = (string)step;
                    if (!map.Contains(k)) return null;
                    node = map[k];
                }
            }
            return node as string ?? (node == null ? null : Convert.ToString(node));
        }
    }
}
