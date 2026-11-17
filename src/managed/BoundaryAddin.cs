using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using Excel = Microsoft.Office.Interop.Excel;

[assembly: AssemblyVersion("0.1.2.0")]
[assembly: AssemblyFileVersion("0.1.2.0")]
[assembly: ComVisible(false)]

[ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IDTExtensibility2 {
    [DispId(1)] void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object application, [In] int mode, [In, MarshalAs(UnmanagedType.IDispatch)] object addin, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
    [DispId(2)] void OnDisconnection([In] int mode, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
    [DispId(3)] void OnAddInsUpdate([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
    [DispId(4)] void OnStartupComplete([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
    [DispId(5)] void OnBeginShutdown([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
}
[ComVisible(true), Guid("4CF9614F-D9B0-4688-9F28-24874CC8FDB8"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IBoundaryStatus {
    [DispId(1)] string Status();
#if BOUNDARY_LAB
    [DispId(101)] void LabSetState(int editMode, int emptyFind);
#endif
}

// Keep the original registration identity so existing installations can be upgraded in place.
#if BOUNDARY_LAB
[ComVisible(true), Guid("E510E056-1108-46AA-9F13-3A461F0AE244"), ProgId("ExcelBoundary.Lab")]
#else
[ComVisible(true), Guid("6C31AE20-5B57-43E6-8984-790CFE5A5FE3"), ProgId("Codex.ReadingSmooth")]
#endif
[ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(IBoundaryStatus))]
public sealed class ReadingAddin : IDTExtensibility2, IBoundaryStatus {
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
    [StructLayout(LayoutKind.Sequential)] struct LastInput { public uint Size, Tick; }
    [DllImport("user32.dll")] static extern bool GetLastInputInfo(ref LastInput value);

    Excel.Application events;
    dynamic excel;
    DynamicExtentController controller;
    System.Windows.Forms.Timer timer;
    BoundarySettings settings;
    MaintenanceSchedule schedule = new MaintenanceSchedule();
    string folder, logs, error = "", activeSheet = "";
    bool nativeReady, active, pending, working, disposed, waitingForReady;
    long selectAfter;
    uint pid;
    int ticks, maintenanceCount, skippedMotion, skippedEdit, editMode = -1;
    double maxMaintenanceMs, totalMaintenanceMs;
#if BOUNDARY_LAB
    const bool background = true;
#else
    const bool background = false;
#endif

    void Log(string message, bool force = false) {
        if (logs == null || (!force && (settings == null || !settings.Diagnostics))) return;
        File.AppendAllText(Path.Combine(logs, "addin-events.log"), DateTime.Now.ToString("o") + " pid=" + pid + " " + message + Environment.NewLine);
    }
    void RecordError(Exception exception) {
        if (error == exception.Message) return;
        error = exception.Message;
        Log(exception.ToString(), true);
    }
    public void OnConnection(object application, int mode, object addin, ref Array custom) {
        try {
            excel = application; events = (Excel.Application)application;
            pid = (uint)Process.GetCurrentProcess().Id;
            folder = Path.GetDirectoryName(typeof(ReadingAddin).Assembly.Location);
            logs = Environment.GetEnvironmentVariable("EXCEL_BOUNDARY_LOG_DIR") ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CodexExcelBoundary", "logs");
            Directory.CreateDirectory(logs);
            ((dynamic)addin).Object = this;
            settings = BoundarySettings.Load(folder);
            timer = new System.Windows.Forms.Timer { Interval = 200 };
            timer.Tick += Tick;
            events.WorkbookOpen += OnWorkbookOpen;
            events.WorkbookActivate += OnWorkbookActivate;
            events.WorkbookDeactivate += OnWorkbookDeactivate;
            events.SheetActivate += OnSheetActivate;
            events.WorkbookBeforeClose += OnBeforeClose;
            InitializeNative();
            RequestSelection();
            Log("connected mode=" + mode + " native=" + nativeReady);
        } catch (Exception exception) { RecordError(exception); }
    }
    void InitializeNative() {
        if (nativeReady || settings == null) return;
        string[] keys = { "CODEX_EXCEL_ADDIN_BOOTSTRAP", "CODEX_EXCEL_DYNAMIC_BOUNDARY", "CODEX_EXCEL_BOUNDARY_LOG_DIR", "CODEX_EXCEL_BOUNDARY_VERBOSE" };
        string[] values = { "1", "1", logs, settings.Diagnostics ? "1" : "0" }, saved = new string[4];
        try {
            for (int i = 0; i < keys.Length; ++i) { saved[i] = Environment.GetEnvironmentVariable(keys[i]); Environment.SetEnvironmentVariable(keys[i], values[i]); }
            nativeReady = (bool)excel.RegisterXLL(Path.Combine(folder, "ReadingBoundary.xll"));
            if (!nativeReady) error = "Native scrolling is unavailable for this Excel build. Supported: 64-bit 16.0.20527.20196.";
        } finally { for (int i = 0; i < keys.Length; ++i) Environment.SetEnvironmentVariable(keys[i], saved[i]); }
    }
    void RequestSelection() {
        pending = true; selectAfter = MaintenanceSchedule.Now + 200;
        if (timer != null && nativeReady) timer.Start();
    }
    void OnWorkbookOpen(Excel.Workbook book) { RequestSelection(); }
    void OnWorkbookActivate(Excel.Workbook book) { RequestSelection(); }
    void OnSheetActivate(object sheet) { RequestSelection(); }
    void OnBeforeClose(Excel.Workbook book, ref bool cancel) { RequestSelection(); }
    void OnWorkbookDeactivate(Excel.Workbook book) { Deactivate(); RequestSelection(); }
    void Deactivate() {
        if (timer != null) timer.Stop();
        if (controller != null) { controller.Suspend(); controller.Dispose(); controller = null; }
        if (active) Log("suspended");
        active = false; activeSheet = "";
    }
    void SelectWorkbook() {
        dynamic book = excel.ActiveWorkbook;
        Excel.Worksheet sheet = excel.ActiveSheet as Excel.Worksheet;
        string name = sheet == null ? "" : sheet.Name;
        bool matches = book != null && sheet != null && String.Equals((string)book.FullName, settings.WorkbookPath, StringComparison.OrdinalIgnoreCase) &&
            (settings.WorksheetName.Length == 0 || name == settings.WorksheetName);
        if (!matches) { Deactivate(); pending = false; return; }
        if (controller != null && name != activeSheet) Deactivate();
        if (controller == null) {
            controller = new DynamicExtentController((object)excel, (object)book, pid, logs, settings.IgnoredHeaders, settings.Diagnostics);
            controller.Refresh(); active = true; activeSheet = name;
            schedule = new MaintenanceSchedule(); timer.Start();
            Log("active sheet=" + name + " blank=" + controller.BoundaryRow);
        }
        pending = false;
    }
    bool CanReadWorkbook() {
        editMode = ExcelEditState.ReadMode();
        if (editMode != 0) {
            waitingForReady = true; ++skippedEdit;
            return false;
        }
        if (waitingForReady) {
            waitingForReady = false;
            if (controller != null) { controller.Dirty = true; controller.GeometryDirty = true; }
        }
        return true;
    }
    void Tick(object sender, EventArgs args) {
        if (working || disposed) return;
        working = true; ++ticks;
        try {
            if (!nativeReady) { timer.Stop(); return; }
            long now = MaintenanceSchedule.Now;
            if (pending) { if (now < selectAfter || !CanReadWorkbook()) return; SelectWorkbook(); }
            if (!active || controller == null) return;
            uint foregroundPid; GetWindowThreadProcessId(GetForegroundWindow(), out foregroundPid);
            if (!background && foregroundPid != pid) return;
            if (controller.ScrollingBusy) { schedule.ObserveMotion(now); ++skippedMotion; return; }
            if (!schedule.Settled(now)) return;
            var input = new LastInput { Size = 8 }; GetLastInputInfo(ref input);
            uint idle = unchecked((uint)Environment.TickCount - input.Tick);
            if (!background && idle < 180) return;
            bool entry = schedule.EntryDue(now);
            bool geometry = schedule.GeometryDue(now, background ? UInt32.MaxValue : idle, input.Tick, controller.GeometryDirty);
            if (!controller.Dirty && !entry && !geometry) return;
            if (!CanReadWorkbook()) return;
            var clock = Stopwatch.StartNew();
            if (controller.Dirty) { controller.Refresh(); if (controller.Dirty) return; }
            if (entry) { controller.PollEntryRow(); schedule.EntryDone(now); }
            if (geometry) { controller.CheckGeometry(); schedule.GeometryDone(now, input.Tick); }
            clock.Stop(); ++maintenanceCount;
            totalMaintenanceMs += clock.Elapsed.TotalMilliseconds;
            maxMaintenanceMs = Math.Max(maxMaintenanceMs, clock.Elapsed.TotalMilliseconds);
        } catch (COMException exception) {
            if (exception.HResult != unchecked((int)0x800AC472) && exception.HResult != unchecked((int)0x80010001)) RecordError(exception);
        } catch (Exception exception) { RecordError(exception); }
        finally { working = false; }
    }
    public string Status() {
        return new JavaScriptSerializer().Serialize(new {
            pid, nativeReady, active, timerRunning = timer != null && timer.Enabled, ticks,
            blankRow = controller == null ? 0 : controller.BoundaryRow,
            lastRow = controller == null ? 0 : controller.LastDataRow,
            lastColumn = controller == null ? 0 : controller.LastColumn,
            revision = controller == null ? 0 : controller.Revision,
            extentPending = controller != null && controller.Dirty,
            nativeBindings = controller == null ? 0 : controller.NativeBindings,
            busy = controller != null && controller.ScrollingBusy,
            editMode, skippedEdit, maintenanceCount, skippedMotion, maxMaintenanceMs, totalMaintenanceMs, error
        });
    }
#if BOUNDARY_LAB
    public void LabSetState(int mode, int emptyFind) {
        ExcelEditState.ModeOverride = mode;
        DynamicExtentController.EmptyFindOverride = emptyFind;
        if (controller != null) controller.Dirty = true;
    }
#endif
    void Stop() {
        if (disposed) return;
        disposed = true;
        try { Deactivate(); } catch (COMException exception) { RecordError(exception); }
        if (timer != null) { timer.Dispose(); timer = null; }
        if (events != null) {
            events.WorkbookOpen -= OnWorkbookOpen; events.WorkbookActivate -= OnWorkbookActivate;
            events.WorkbookDeactivate -= OnWorkbookDeactivate; events.SheetActivate -= OnSheetActivate;
            events.WorkbookBeforeClose -= OnBeforeClose;
        }
        Log("disconnected"); events = null; excel = null;
    }
    public void OnDisconnection(int mode, ref Array custom) { Stop(); }
    public void OnStartupComplete(ref Array custom) { if (!nativeReady) InitializeNative(); RequestSelection(); }
    public void OnAddInsUpdate(ref Array custom) { }
    public void OnBeginShutdown(ref Array custom) { Stop(); }
}
