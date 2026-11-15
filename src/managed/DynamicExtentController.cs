using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel=Microsoft.Office.Interop.Excel;

sealed class DynamicExtentController:IDisposable {
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern IntPtr FindWindowEx(IntPtr parent,IntPtr after,string cls,string title);
    [DllImport("user32.dll")] static extern uint GetDpiForWindow(IntPtr window);
    [StructLayout(LayoutKind.Sequential)] struct CopyData {public UIntPtr Kind;public int Length;public IntPtr Data;}
    [DllImport("user32.dll",SetLastError=true)] static extern IntPtr SendMessageTimeout(IntPtr window,uint message,UIntPtr wp,ref CopyData data,uint flags,uint timeout,out UIntPtr result);
    [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr window,uint message,IntPtr wp,IntPtr lp);
    readonly Excel.Application events;
    readonly dynamic excel;
    readonly IntPtr receiver;
    readonly string log;
    readonly string workbookPath;
    readonly HashSet<string> ignoredHeaders;
    readonly bool diagnostics;
    readonly string worksheetName;
    readonly List<string> dataRegions=new List<string>();
    bool[] dataColumns;
    dynamic entryRange;
    dynamic layoutRange,layoutWindow;
    double layoutWidth,layoutHeight;
    sealed class DimensionRun {
        public dynamic Range;
        public bool Columns;
        public double Points;
        public object Uniform;
        public bool Unchanged() {
            object value=Columns?Range.ColumnWidth:Range.RowHeight;
            if(Points==0)return (double)(Columns?Range.Width:Range.Height)==0;
            return value is double && Uniform is double && Math.Abs((double)value-(double)Uniform)<0.001;
        }
    }
    readonly List<DimensionRun> dimensionRuns=new List<DimensionRun>();
    double[] previous;
    uint previousDpi;
    uint previousFrozenColumns,previousFrozenRows;
    readonly IntPtr excelWindow;
    bool appendedAtEnd;
    public int LastDataRow {get;private set;}
    public int BoundaryRow {get;private set;}
    public int LastColumn {get;private set;}
    public uint Revision {get;private set;}
    public bool Dirty=true;
    public bool GeometryDirty=true;
    public bool ScrollingBusy {get{return SendMessage(receiver,0x8057,IntPtr.Zero,IntPtr.Zero)!=IntPtr.Zero;}}
    public int NativeBindings {get{return (int)SendMessage(receiver,0x8058,IntPtr.Zero,IntPtr.Zero);}}
    public DynamicExtentController(object application,object workbook,uint pid,string folder,string[] excludedHeaders,bool diagnosticLogging=false) {
        excel=application;events=(Excel.Application)application;
        ignoredHeaders=new HashSet<string>(excludedHeaders,StringComparer.OrdinalIgnoreCase);
        diagnostics=diagnosticLogging;
        workbookPath=(string)((dynamic)workbook).FullName;
        worksheetName=(string)excel.ActiveSheet.Name;
        excelWindow=new IntPtr((int)excel.Hwnd);
        receiver=FindWindowEx(new IntPtr(-3),IntPtr.Zero,"CodexExcelDynamicBoundary10","CodexExcelBoundary_"+pid);
        if(receiver==IntPtr.Zero)throw new Exception("动态边界服务未启动。");
        log=Path.Combine(folder,"dynamic-extent-events.jsonl");
        events.SheetChange+=OnSheetChange;
        events.WindowResize+=OnWindowResize;
    }
    void OnSheetChange(object sheet,Excel.Range changed){
        if((string)((dynamic)sheet).Name!=worksheetName)return;
        Dirty=true;GeometryDirty=true;appendedAtEnd=true;
        if(changed.Row==1)dataColumns=null;
    }
    void OnWindowResize(Excel.Workbook changed,Excel.Window window){GeometryDirty=true;}
    static string ColumnName(int column){string name="";while(column>0){--column;name=(char)('A'+column%26)+name;column/=26;}return name;}
    static void Dimensions(dynamic sheet,bool columns,int first,int count,double[] points,int offset) {
        string address=columns?ColumnName(first)+":"+ColumnName(first+count-1):first+":"+(first+count-1);
        dynamic range=sheet.Range[address];
        double total=columns?(double)range.Width:(double)range.Height;
        object uniform=count==1||total==0?null:columns?range.ColumnWidth:range.RowHeight;
        bool equal=uniform is double;
        if(equal&&count>1) {
            if(!columns)equal=Math.Abs(total-(double)uniform*count)<0.001;
            else {
                dynamic firstRange=sheet.Range[ColumnName(first)+":"+ColumnName(first)];
                equal=Math.Abs(total-(double)firstRange.Width*count)<0.001;
                Marshal.ReleaseComObject(firstRange);
            }
        }
        if(count==1||total==0||equal) {
            for(int i=0;i<count;++i)points[offset+i]=total/count;
        } else {
            int half=count/2;
            Dimensions(sheet,columns,first,half,points,offset);
            Dimensions(sheet,columns,first+half,count-half,points,offset+half);
        }
        Marshal.ReleaseComObject(range);
    }
    void ClearLayoutCache() {
        foreach(var run in dimensionRuns)Marshal.ReleaseComObject(run.Range);
        dimensionRuns.Clear();
        if(layoutRange!=null){Marshal.ReleaseComObject(layoutRange);layoutRange=null;}
        if(entryRange!=null){Marshal.ReleaseComObject(entryRange);entryRange=null;}
        if(layoutWindow!=null){Marshal.ReleaseComObject(layoutWindow);layoutWindow=null;}
    }
    void CacheLayout(dynamic sheet,int columns,int rows,double[] points) {
        ClearLayoutCache();
        layoutRange=sheet.Range["A1:"+ColumnName(columns)+rows];layoutWindow=excel.ActiveWindow;
        entryRange=sheet.Range["A"+Math.Max(1,rows-1)+":"+ColumnName(columns)+rows];
        layoutWidth=points.Take(columns).Sum();layoutHeight=points.Skip(columns).Sum();
        foreach(bool isColumn in new[]{true,false}) {
            int count=isColumn?columns:rows,offset=isColumn?0:columns;
            for(int start=0;start<count;) {
                int end=start+1;while(end<count&&Math.Abs(points[offset+end]-points[offset+start])<0.001)++end;
                string address=isColumn?ColumnName(start+1)+":"+ColumnName(end):(start+1)+":"+end;
                dynamic range=sheet.Range[address];
                var run=new DimensionRun {Range=range,Columns=isColumn,Points=points[offset+start]};
                run.Uniform=isColumn?range.ColumnWidth:(object)run.Points;
                dimensionRuns.Add(run);start=end;
            }
        }
    }
    public bool CheckGeometry() {
        if(previous==null||Dirty||layoutRange==null)return Refresh(true);
        dynamic active=excel.ActiveWorkbook;
        if(active==null)return false;
        bool matches=(string)active.FullName==workbookPath;Marshal.ReleaseComObject(active);
        if(!matches)return false;
        bool frozen=(bool)layoutWindow.FreezePanes;
        uint columns=frozen?(uint)(int)layoutWindow.SplitColumn:0,rows=frozen?(uint)(int)layoutWindow.SplitRow:0;
        if(columns!=previousFrozenColumns||rows!=previousFrozenRows||GetDpiForWindow(excelWindow)!=previousDpi||
           Math.Abs((double)layoutRange.Width-layoutWidth)>0.001||Math.Abs((double)layoutRange.Height-layoutHeight)>0.001)
            return Refresh(true);
        foreach(var run in dimensionRuns)if(!run.Unchanged())return Refresh(true);
        GeometryDirty=false;return false;
    }
    void CacheDataColumns(dynamic sheet,int columns) {
        dataRegions.Clear();dataColumns=new bool[columns];
        dynamic header=sheet.Range["A1:"+ColumnName(columns)+"1"];
        object raw=header.Value2;object[,] names=raw as object[,];Marshal.ReleaseComObject(header);
        int first=0;
        for(int column=1;column<=columns+1;++column) {
            bool included=column<=columns;
            if(included) {
                string name=Convert.ToString(names==null?raw:names[1,column]);
                included=!ignoredHeaders.Contains(name);
                dataColumns[column-1]=included;
            }
            if(included&&first==0)first=column;
            if(!included&&first>0){dataRegions.Add(ColumnName(first)+":"+ColumnName(column-1));first=0;}
        }
    }
    int LastReadingRow(dynamic sheet,int columns) {
        if(dataColumns==null||dataColumns.Length!=columns)CacheDataColumns(sheet,columns);
        int last=1;
        foreach(string address in dataRegions) {
            dynamic range=sheet.Range[address];
            dynamic found=range.Find("*",Type.Missing,-4163,Type.Missing,1,2,false,Type.Missing,false);
            if(found!=null){last=Math.Max(last,(int)found.Row);Marshal.ReleaseComObject(found);}
            Marshal.ReleaseComObject(range);
        }
        return last;
    }
    public bool PollEntryRow() {
        if(entryRange==null)return false;
        object[,] values=(object[,])entryRange.Value2;
        bool last=false,next=false;
        for(int column=1;column<=dataColumns.Length;++column)if(dataColumns[column-1]) {
            last|=values[1,column]!=null&&Convert.ToString(values[1,column]).Length>0;
            next|=values[values.GetUpperBound(0),column]!=null&&Convert.ToString(values[values.GetUpperBound(0),column]).Length>0;
        }
        if(next||(!last&&LastDataRow>1)){Dirty=true;appendedAtEnd=next;return Refresh();}
        return false;
    }
    public bool Refresh(bool fullGeometry=false) {
        if(previous!=null&&!Dirty&&!fullGeometry)return false;
        dynamic active=excel.ActiveWorkbook;
        if(active==null)return false;
        bool matches=(string)active.FullName==workbookPath;Marshal.ReleaseComObject(active);
        if(!matches)return false;
        dynamic sheet=excel.ActiveSheet;
        dynamic lastColumn=sheet.Cells.Find("*",Type.Missing,-4123,Type.Missing,2,2,false,Type.Missing,false);
        int columns=lastColumn==null?1:(int)lastColumn.Column;
        int dataRow=LastReadingRow(sheet,columns);
        int rows=Math.Min(1048576,dataRow+1);
        bool follow=false;
        if(appendedAtEnd&&rows>BoundaryRow&&BoundaryRow>0) {
            dynamic visible=excel.ActiveWindow.ActivePane.VisibleRange;
            follow=(int)visible.Row<=BoundaryRow&&(int)visible.Row+(int)visible.Rows.Count>BoundaryRow;
        }
        fullGeometry=fullGeometry||previous==null||columns!=LastColumn||rows<BoundaryRow;
        var points=new double[columns+rows];
        uint frozenColumns=previousFrozenColumns,frozenRows=previousFrozenRows;
        if(fullGeometry) {
            Dimensions(sheet,true,1,columns,points,0);
            Dimensions(sheet,false,1,rows,points,columns);
            dynamic window=excel.ActiveWindow;
            bool frozen=(bool)window.FreezePanes;
            frozenColumns=frozen?(uint)(int)window.SplitColumn:0;
            frozenRows=frozen?(uint)(int)window.SplitRow:0;
        } else {
            Array.Copy(previous,points,Math.Min(previous.Length,points.Length));
            if(rows>BoundaryRow)Dimensions(sheet,false,BoundaryRow,rows-BoundaryRow+1,points,columns+BoundaryRow-1);
        }
        uint dpi=GetDpiForWindow(excelWindow);
        bool pointsChanged=previous==null||!previous.SequenceEqual(points);
        bool changed=pointsChanged||columns!=LastColumn||rows!=BoundaryRow||dpi!=previousDpi||frozenColumns!=previousFrozenColumns||frozenRows!=previousFrozenRows;
        LastDataRow=dataRow;
        if(lastColumn!=null)Marshal.ReleaseComObject(lastColumn);
        if(pointsChanged||layoutRange==null)CacheLayout(sheet,columns,rows,points);
        if(changed) {
            var payload=new byte[32+points.Length*8];
            var header=new uint[]{12,(uint)columns,(uint)rows,Revision+1,dpi,frozenColumns,frozenRows,0};
            Buffer.BlockCopy(header,0,payload,0,32);Buffer.BlockCopy(points,0,payload,32,points.Length*8);
            var pinned=GCHandle.Alloc(payload,GCHandleType.Pinned);
            try {
                var copy=new CopyData {Kind=new UIntPtr(0x43425831),Length=payload.Length,Data=pinned.AddrOfPinnedObject()};UIntPtr result;
                if(SendMessageTimeout(receiver,0x4a,UIntPtr.Zero,ref copy,2,2000,out result)==IntPtr.Zero||result==UIntPtr.Zero)
                    throw new Exception("更新动态边界失败。");
            } finally {pinned.Free();}
            ++Revision;previous=points;previousDpi=dpi;LastColumn=columns;BoundaryRow=rows;
            previousFrozenColumns=frozenColumns;previousFrozenRows=frozenRows;
            if(diagnostics)File.AppendAllText(log,"{\"revision\":"+Revision+",\"lastDataRow\":"+dataRow+",\"boundaryRow\":"+rows+",\"lastColumn\":"+columns+",\"dpi\":"+dpi+"}"+Environment.NewLine);
        }
        if(follow)CenterInputRow();
        appendedAtEnd=false;
        if(fullGeometry)GeometryDirty=false;
        Dirty=false;return changed;
    }
    public void CenterInputRow() {
        if(previous==null)return;
        dynamic window=excel.ActiveWindow;
        dynamic visible=window.ActivePane.VisibleRange;
        double half=(double)visible.Height/2;
        int minimum=(bool)window.FreezePanes?(int)window.SplitRow+1:1;
        int first=BoundaryRow;
        double height=previous[LastColumn+BoundaryRow-1]/2;
        while(first>minimum&&height<half){--first;height+=previous[LastColumn+first-1];}
        window.ScrollRow=first;
    }
    public double InputVerticalFraction() {
        dynamic window=excel.ActiveWindow,visible=window.ActivePane.VisibleRange,sheet=excel.ActiveSheet;
        dynamic cell=sheet.Cells[BoundaryRow,1];
        return ((double)cell.Top+(double)cell.Height/2-(double)visible.Top)/(double)visible.Height;
    }
    public void Suspend() {
        var header=new uint[]{12,0,0,Revision+1,previousDpi,0,0,0};
        var pinned=GCHandle.Alloc(header,GCHandleType.Pinned);
        try {
            var copy=new CopyData{Kind=new UIntPtr(0x43425831),Length=32,Data=pinned.AddrOfPinnedObject()};UIntPtr result;
            SendMessageTimeout(receiver,0x4a,UIntPtr.Zero,ref copy,2,2000,out result);
        } finally {pinned.Free();}
        previous=null;Dirty=true;GeometryDirty=true;ClearLayoutCache();
    }
    public void Dispose(){events.SheetChange-=OnSheetChange;events.WindowResize-=OnWindowResize;ClearLayoutCache();}
}
