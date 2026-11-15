using System;
using System.IO;
using System.Web.Script.Serialization;

public sealed class BoundarySettings {
    public string WorkbookPath { get; set; }
    public string WorksheetName { get; set; }
    public string[] IgnoredHeaders { get; set; }
    public bool Diagnostics { get; set; }

    public BoundarySettings() {
        WorksheetName = "";
        IgnoredHeaders = new string[0];
    }

    public static BoundarySettings Load(string directory) {
        string path = Environment.GetEnvironmentVariable("EXCEL_BOUNDARY_CONFIG") ?? Path.Combine(directory, "settings.json");
        var settings = new JavaScriptSerializer().Deserialize<BoundarySettings>(File.ReadAllText(path));
        if (String.IsNullOrWhiteSpace(settings.WorkbookPath)) throw new InvalidDataException("WorkbookPath is required in settings.json.");
        settings.WorkbookPath = Path.GetFullPath(settings.WorkbookPath);
        settings.WorksheetName = settings.WorksheetName ?? "";
        settings.IgnoredHeaders = settings.IgnoredHeaders ?? new string[0];
        return settings;
    }
}
