using System;
using System.Runtime.InteropServices;

namespace Sadar.Addin
{
    /// <summary>
    /// הכרזות COM שנדרשות לתוסף.
    ///
    /// הן מוכרזות כאן במקום להסתמך על Extensibility.dll ועל office.dll,
    /// כדי שהפרויקט ייבנה על כל מחשב עם Windows בלבד — בלי Visual Studio,
    /// בלי VSTO Runtime ובלי חבילות Primary Interop.
    /// זו הסיבה שהתוסף מתקין בקובץ אחד.
    ///
    /// הממשקים של הרצועה דווקא לא מוכרזים כאן אלא נלקחים מ-office.dll:
    /// הכרזה עצמית של IRibbonExtensibility מייצרת משטח שאינו תואם
    /// למה שוורד קורא דרכו, והתוצאה היא כשל טעינה אילם.
    /// </summary>
    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3,
        ext_cm_Solution = 4,
        ext_cm_UISetup = 5
    }

    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1,
        ext_dm_UISetupComplete = 2,
        ext_dm_SolutionClosed = 3
    }

    [ComVisible(true)]
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection(
            [MarshalAs(UnmanagedType.IDispatch)] object application,
            ext_ConnectMode connectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object addInInst,
            ref Array custom);

        [DispId(2)]
        void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate(ref Array custom);

        [DispId(4)]
        void OnStartupComplete(ref Array custom);

        [DispId(5)]
        void OnBeginShutdown(ref Array custom);
    }
}
