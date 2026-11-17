using System;
using System.Runtime.InteropServices;

// Excel SDK: xlGetFmlaInfo is an LPenHelper operation, not an Excel4/Excel12 call.
static class ExcelEditState {
    [StructLayout(LayoutKind.Sequential)]
    struct FormulaInfo {
        public int Mode, CharacterCount;
        public IntPtr Text;
        public int SelectionStart, SelectionEnd, Caret;
    }
    [DllImport("XLCALL32.DLL", CallingConvention=CallingConvention.StdCall)]
    static extern int LPenHelper(int code, ref FormulaInfo info);

#if BOUNDARY_LAB
    // -2 uses Excel's actual state; the other values exercise the timer's edit gate.
    internal static int ModeOverride = -2;
#endif
    public static int ReadMode() {
#if BOUNDARY_LAB
        if (ModeOverride != -2) return ModeOverride;
#endif
        var info = new FormulaInfo();
        return LPenHelper(0x400e, ref info) == 0 ? info.Mode : -1;
    }
}
