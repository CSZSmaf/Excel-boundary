using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;

static class HitTestChecks {
    [DllImport("user32.dll")] static extern uint GetDpiForWindow(IntPtr window);
    [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr window,uint message,IntPtr wp,IntPtr lp);
    static double Number(object value) { return Convert.ToDouble(value); }
    static double[] Vector(object value) { return ((System.Collections.IEnumerable)value).Cast<object>().Select(Number).ToArray(); }
    static Dictionary<string,object> Snapshot(IntPtr receiver,string logs,int sample) {
        SendMessage(receiver,0x8052,new IntPtr(sample),IntPtr.Zero);
        var json=new JavaScriptSerializer();
        foreach(string line in File.ReadLines(Path.Combine(logs,"boundary-native-events.jsonl")).Reverse()) {
            var value=json.Deserialize<Dictionary<string,object>>(line);
            if(!value.ContainsKey("sample")||Convert.ToInt32(value["sample"])!=sample||Convert.ToString(value["event"])!="native-bound-probe")continue;
            var viewport=Vector(value["viewport"]);var content=Vector(value["content"]);
            if(viewport[0]>500&&viewport[1]>400&&content[0]<10000)return value;
        }
        throw new Exception("Worksheet composition probe not found.");
    }
    public static void Run(dynamic app,dynamic sheet,IntPtr receiver,string logs,Action<int> pump) {
        dynamic window=app.ActiveWindow;
        double scale=GetDpiForWindow(new IntPtr((int)app.Hwnd))/72.0*(double)window.Zoom/100;
        Func<double,double> pixel=points=>Math.Floor(points*scale+0.5);
        double header=pixel((double)sheet.Rows[1].Height),rowHeight=pixel((double)sheet.Rows[2].Height);
        double colWidth=pixel((double)sheet.Columns[1].Width);
        window.ScrollColumn=1;window.ScrollRow=2;pump(500);
        var start=Snapshot(receiver,logs,700);var startOffset=Vector(start["renderedOffset"]);
        double originX=(int)window.PointsToScreenPixelsX(0)-startOffset[0];
        // Calibrate the body origin at the known top position. PointsToScreenPixelsY
        // includes Excel's document/ribbon offsets and differs with frozen panes.
        int guess=(int)window.PointsToScreenPixelsY((int)sheet.Rows[1].Height),row3Top=0;
        int anchorX=(int)Math.Round(originX+colWidth*1.5);
        for(int y=guess;y<guess+600;y+=20) {
            dynamic cell=window.RangeFromPoint(anchorX,y);
            if(cell!=null&&(int)cell.Row==3) {
                int low=y-20,high=y;
                while(high-low>1) {int mid=(low+high)/2;dynamic item=window.RangeFromPoint(anchorX,mid);if(item!=null&&(int)item.Row>=3)high=mid;else low=mid;}
                row3Top=high;break;
            }
        }
        if(row3Top==0)throw new Exception("Could not calibrate the top worksheet pane.");
        double originY=row3Top-rowHeight;
        var results=new List<object>();bool passed=true;
        int sample=701;
        foreach(int[] position in new[]{new[]{2,1},new[]{27,1},new[]{35,1},new[]{2,13}}) {
            int scrollRow=position[0],scrollColumn=position[1];
            window.ScrollRow=scrollRow;window.ScrollColumn=scrollColumn;pump(500);
            var state=Snapshot(receiver,logs,sample++);var offset=Vector(state["renderedOffset"]);var viewport=Vector(state["viewport"]);
            int expectedRow=Math.Max(2,(int)Math.Ceiling((-offset[1]-header)/rowHeight)+3);
            int expectedColumn=Math.Max(2,(int)Math.Ceiling(-offset[0]/colWidth)+2);
            double x=originX+offset[0]+colWidth*(expectedColumn-0.5);
            double y=originY+offset[1]+header+rowHeight*(expectedRow-1.5);
            dynamic hit=window.RangeFromPoint((int)Math.Round(x),(int)Math.Round(y));
            string address=hit==null?"":(string)hit.Address[false,false];
            string expected=((char)('A'+expectedColumn-1)).ToString()+expectedRow;bool matches=address==expected;passed&=matches;
            results.Add(new {scrollRow,scrollColumn,logicalScrollRow=(int)window.ScrollRow,logicalScrollColumn=(int)window.ScrollColumn,renderedOffset=offset,screen=new[]{x,y},expected,address,matches});
        }
        File.WriteAllText(Path.Combine(logs,"hit-test.json"),new JavaScriptSerializer().Serialize(new {passed,scale,originX,originY,header,rowHeight,colWidth,results}));
        if(!passed)throw new Exception("Visible cells and RangeFromPoint disagree. See hit-test.json.");
        window.ScrollColumn=1;
    }
}
