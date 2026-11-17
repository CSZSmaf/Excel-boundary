using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using Microsoft.Win32;

class ExcelIntegration {
    [ComImport,Guid("00000016-0000-0000-C000-000000000046"),InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IMessageFilter {
        [PreserveSig]int HandleInComingCall(int type,IntPtr caller,int tick,IntPtr info);
        [PreserveSig]int RetryRejectedCall(IntPtr callee,int tick,int rejection);
        [PreserveSig]int MessagePending(IntPtr callee,int tick,int type);
    }
    sealed class RetryFilter:IMessageFilter {
        public int HandleInComingCall(int t,IntPtr c,int k,IntPtr i){return 0;}
        public int RetryRejectedCall(IntPtr c,int tick,int r){return tick<5000?100:-1;}
        public int MessagePending(IntPtr c,int tick,int t){return 2;}
    }
    [DllImport("ole32.dll")]static extern int CoRegisterMessageFilter(IMessageFilter value,out IMessageFilter old);
    [DllImport("user32.dll")]static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
    [DllImport("user32.dll")]static extern IntPtr GetThreadDesktop(uint tid);
    [DllImport("kernel32.dll")]static extern uint GetCurrentThreadId();
    [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern bool GetUserObjectInformation(IntPtr handle,int index,StringBuilder name,int size,out int needed);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern IntPtr FindWindowEx(IntPtr parent,IntPtr after,string cls,string title);
    [DllImport("user32.dll")]static extern IntPtr SendMessage(IntPtr hwnd,uint message,IntPtr wp,IntPtr lp);
    static string Desktop(uint tid){var name=new StringBuilder(256);int n;GetUserObjectInformation(GetThreadDesktop(tid),2,name,512,out n);return name.ToString();}
    static void Pump(int ms){var clock=Stopwatch.StartNew();while(clock.ElapsedMilliseconds<ms){System.Windows.Forms.Application.DoEvents();Thread.Sleep(15);}}
    static Dictionary<string,object> Status(dynamic addin){return new JavaScriptSerializer().Deserialize<Dictionary<string,object>>((string)addin.Object.Status());}
    static void Check(bool value,string message){if(!value)throw new Exception(message);}
    static void RestoreProductionRegistration(){
        string previous=Environment.GetEnvironmentVariable("EXCEL_BOUNDARY_PREVIOUS_LOAD");
        if(previous==null)return;
        using(var key=Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\Excel\Addins\Codex.ReadingSmooth",true))
            if(key!=null)key.SetValue("LoadBehavior",Int32.Parse(previous),RegistryValueKind.DWord);
    }
    [STAThread]static int Main(){
        string folder=AppDomain.CurrentDomain.BaseDirectory,logs=Environment.GetEnvironmentVariable("EXCEL_BOUNDARY_LOG_DIR"),target=Path.Combine(folder,"Fixture.xlsx");
        Directory.CreateDirectory(logs);
        string oldError=Path.Combine(logs,"error.txt");if(File.Exists(oldError))File.Delete(oldError);
        string config=Path.Combine(folder,"settings.json");
        File.WriteAllText(config,new JavaScriptSerializer().Serialize(new {WorkbookPath=target,WorksheetName="",IgnoredHeaders=new[]{"Read count","Record ID"},Diagnostics=true}));
        Environment.SetEnvironmentVariable("EXCEL_BOUNDARY_CONFIG",config);
        dynamic app=null,book=null,other=null;uint pid=0;bool passed=false;
        IMessageFilter oldFilter;CoRegisterMessageFilter(new RetryFilter(),out oldFilter);
        try {
            string desktop=Desktop(GetCurrentThreadId());Check(desktop.StartsWith("ExcelBoundaryLab_"),"Private desktop required");
            app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            uint tid=GetWindowThreadProcessId(new IntPtr((int)app.Hwnd),out pid);
            RestoreProductionRegistration();Check(Desktop(tid)==desktop,"Excel stayed on its private desktop");
            app.DisplayAlerts=false;app.AutomationSecurity=3;app.Visible=true;app.WindowState=-4137;
            dynamic addin=app.COMAddIns.Item("ExcelBoundary.Lab");Check((bool)addin.Connect,"COM add-in automatically loaded");
            Pump(600);
            var idle1=Status(addin);Pump(500);var idle2=Status(addin);
            Check(!(bool)idle2["timerRunning"]&&Object.Equals(idle1["ticks"],idle2["ticks"]),"No-workbook timer stopped");
            book=app.Workbooks.Add();dynamic sheet=book.Worksheets[1];sheet.Name="Records";
            sheet.Range["A:O"].ColumnWidth=28;sheet.Range["1:101"].RowHeight=42;sheet.Range["1:1"].RowHeight=27;
            var values=new object[101,15];
            for(int column=0;column<15;++column)values[0,column]="Column "+(column+1);
            values[0,0]="Title";values[0,11]="Read count";values[0,14]="Record ID";
            for(int row=1;row<101;++row){values[row,11]=0;values[row,14]=row;if(row<=35){values[row,0]="Sample "+row;values[row,3]="Synthetic sample text for scrolling.";}}
            sheet.Range["A1:O101"].Value2=values;
            sheet.Range["A1:N101"].WrapText=true;
            app.ActiveWindow.Zoom=90;app.ActiveWindow.SplitColumn=0;app.ActiveWindow.SplitRow=1;app.ActiveWindow.FreezePanes=true;
            book.SaveAs(target,51);book.Close(false);book=app.Workbooks.Open(target,0,false);sheet=book.Worksheets[1];
            app.Visible=true;app.WindowState=-4137;Pump(1600);
            var initial=Status(addin);Check((bool)initial["nativeReady"]&&(bool)initial["active"]&&Convert.ToInt32(initial["nativeBindings"])>0,"Target native scope active");
            Check(Convert.ToInt32(initial["lastRow"])==36&&Convert.ToInt32(initial["blankRow"])==37,"Helper columns excluded from content extent");
            Check(Convert.ToInt32(initial["editMode"])==0,"Native edit-state query reports Ready");
            // Fault injection stays inside the add-in; no keyboard, mouse or desktop capture.
            foreach(int mode in new[]{1,2,4,-1}) {
                var before=Status(addin);addin.Object.LabSetState(mode,3);Pump(450);var editing=Status(addin);
                Check(Object.Equals(before["revision"],editing["revision"])&&Object.Equals(before["maintenanceCount"],editing["maintenanceCount"]),"Editing defers all COM maintenance: "+mode);
                Check(Convert.ToInt32(editing["skippedEdit"])>Convert.ToInt32(before["skippedEdit"]),"Edit gate observed the mode: "+mode);
                Check(Convert.ToInt32(editing["blankRow"])==37&&Convert.ToInt32(editing["lastColumn"])==15,"Editing preserves the published extent");
            }
            foreach(int missingFind in new[]{1,2}) {
                addin.Object.LabSetState(-2,missingFind);Pump(450);var missing=Status(addin);
                Check(Convert.ToInt32(missing["blankRow"])==37&&Convert.ToInt32(missing["lastColumn"])==15&&(bool)missing["extentPending"],"Empty search preserves last good extent: "+missingFind);
            }
            addin.Object.LabSetState(-2,0);Pump(650);
            Check(!(bool)Status(addin)["extentPending"],"Maintenance resumes after edit/search recovery");
            var receiver=FindWindowEx(new IntPtr(-3),IntPtr.Zero,"CodexExcelDynamicBoundary10","CodexExcelBoundary_"+pid);
            SendMessage(receiver,0x8055,IntPtr.Zero,IntPtr.Zero);
            for(int phase=0;phase<6;++phase){SendMessage(receiver,0x8056,new IntPtr(phase),IntPtr.Zero);for(int sample=0;sample<8;++sample){Pump(50);SendMessage(receiver,0x8052,new IntPtr(phase*100+sample),IntPtr.Zero);}}
            SendMessage(receiver,0x8055,new IntPtr(1),IntPtr.Zero);
            Pump(1200);sheet.Cells[37,1].Select();app.ActiveWindow.ScrollRow=27;
            app.EnableEvents=false;sheet.Cells[37,1].Value2="Appended with events disabled";app.EnableEvents=true;Pump(2300);
            var appended=Status(addin);Check(Convert.ToInt32(appended["blankRow"])==38,"Event-disabled paste detected");
            dynamic visible=app.ActiveWindow.ActivePane.VisibleRange;
            double fraction=((double)sheet.Cells[38,1].Top+(double)sheet.Cells[38,1].Height/2-(double)visible.Top)/(double)visible.Height;
            Check(fraction>0.35&&fraction<0.65,"Entry row centered after append");
            sheet.Cells[37,1].ClearContents();Pump(900);
            Check(Convert.ToInt32(Status(addin)["blankRow"])==37,"Deleting the last record still contracts the boundary");
            other=app.Workbooks.Add();Pump(650);var suspended=Status(addin);Pump(500);var stillSuspended=Status(addin);
            Check(!(bool)suspended["active"]&&!(bool)suspended["timerRunning"]&&(int)SendMessage(receiver,0x8058,IntPtr.Zero,IntPtr.Zero)==0,"Other workbook restores native boundaries");
            Check(Object.Equals(suspended["ticks"],stillSuspended["ticks"]),"No timer callbacks while other workbook active");
            addin.Object.LabSetState(4,3);book.Activate();Pump(450);
            Check(!(bool)Status(addin)["active"],"Initial activation waits for reference selection to end");
            addin.Object.LabSetState(-2,0);Pump(700);Check((bool)Status(addin)["active"],"Returning to target resumes boundaries");
            dynamic single=book.Worksheets.Add();single.Name="Single";single.Cells[1,1].Value2="Title";single.Cells[2,1].Value2="One item";Pump(1200);
            Check(Convert.ToInt32(Status(addin)["blankRow"])==3,"Single-column header is supported");
            single.Range["A1:A2"].ClearContents();Pump(800);
            Check(Convert.ToInt32(Status(addin)["blankRow"])==2&&Convert.ToInt32(Status(addin)["lastColumn"])==1,"A truly empty sheet can use A1:A2");
            single.Cells[3,2].Value2="New content";Pump(800);
            Check(Convert.ToInt32(Status(addin)["blankRow"])==4&&Convert.ToInt32(Status(addin)["lastColumn"])==2,"An emptied sheet can expand again");
            sheet.Activate();Pump(700);var final=Status(addin);Check(Convert.ToString(final["error"])=="","No add-in errors");
            passed=true;
            File.WriteAllText(Path.Combine(logs,"integration.json"),new JavaScriptSerializer().Serialize(new {passed,pid,automaticLoad=true,ignoredHeaders=true,editStateQuery=true,simulatedEditModes=new[]{1,2,4,-1},emptySearchRecovery=true,deleteAndClear=true,eventDisabledAppend=true,entryVerticalFraction=fraction,inactiveTimerCallbacks=0,singleColumn=true,initial,final}));
            return 0;
        }catch(Exception exception){File.WriteAllText(Path.Combine(logs,"error.txt"),exception.ToString());return 1;}
        finally {
            RestoreProductionRegistration();
            try{app.DisplayAlerts=false;if(other!=null)other.Close(false);if(book!=null)book.Close(false);app.Quit();}catch{}
            book=null;other=null;app=null;GC.Collect();GC.WaitForPendingFinalizers();
            if(pid>0)try{var process=Process.GetProcessById((int)pid);if(!process.WaitForExit(2000))process.Kill();}catch{}
        }
    }
}
