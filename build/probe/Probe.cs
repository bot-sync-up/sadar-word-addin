using System;
using System.IO;
using System.Runtime.InteropServices;

namespace Sadar.Probe
{
    /// <summary>
    /// תוסף מינימלי לבידוד תקלות טעינה.
    ///
    /// אין לו שום תלות: לא Word interop, לא WinForms, לא הליבה.
    /// רק IDTExtensibility2 והכרזות COM. אם וורד טוען אותו והתוסף
    /// האמיתי נכשל — הבעיה בתלויות שלנו. אם גם הוא נכשל — הבעיה
    /// בסביבה, ואין טעם לחפש אותה בקוד.
    ///
    /// כשהוא נטען הוא כותב קובץ עדות. זו הדרך היחידה לדעת ש-OnConnection
    /// באמת רץ בתוך וורד, בלי לפתוח ממשק ובלי להסתכל על המסך.
    /// </summary>
    public enum ext_ConnectMode { ext_cm_AfterStartup = 0 }
    public enum ext_DisconnectMode { ext_dm_HostShutdown = 0 }

    // בלי [ComImport]: הוא אומר ל-CLR שהממשק מוגדר בטיפוסייה חיצונית,
    // ולכן לא נוצר גשר IDispatch עבורו. וורד יוצר את האובייקט בהצלחה
    // ואז לא מצליח לקרוא לאף שיטה.
    [ComVisible(true)]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection([MarshalAs(UnmanagedType.IDispatch)] object application,
                          ext_ConnectMode connectMode,
                          [MarshalAs(UnmanagedType.IDispatch)] object addInInst,
                          ref object custom);
        [DispId(2)] void OnDisconnection(ext_DisconnectMode removeMode, ref object custom);
        [DispId(3)] void OnAddInsUpdate(ref object custom);
        [DispId(4)] void OnStartupComplete(ref object custom);
        [DispId(5)] void OnBeginShutdown(ref object custom);
    }

    [ComVisible(true)]
    [Guid("3B1C7E90-5A44-4D2B-9F71-1C4E8A6D2F03")]
    [ProgId("SadarProbe.Connect")]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class Connect : IDTExtensibility2
    {
        private static readonly string LogFile =
            Path.Combine(Path.GetTempPath(), "sadar-probe.log");

        public Connect()
        {
            Write("ctor");
        }

        public void OnConnection(object application, ext_ConnectMode connectMode,
                                 object addInInst, ref object custom)
        {
            Write("OnConnection  host=" + (application == null ? "null" : "ok"));
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref object custom) { Write("OnDisconnection"); }
        public void OnAddInsUpdate(ref object custom) { }
        public void OnStartupComplete(ref object custom) { Write("OnStartupComplete"); }
        public void OnBeginShutdown(ref object custom) { }

        private static void Write(string message)
        {
            try
            {
                File.AppendAllText(LogFile,
                    DateTime.Now.ToString("HH:mm:ss") + "  " + message + Environment.NewLine);
            }
            catch { }
        }
    }
}
