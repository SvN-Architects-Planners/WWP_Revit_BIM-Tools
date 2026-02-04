
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Markup;
using Microsoft.Win32;
using System.Runtime.InteropServices;

public class Script
{
    public static void Execute(UIApplication uiapp)
    {
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

        Window window = UiLoader.LoadWindow("ui.xaml");
        if (window == null)
        {
            TaskDialog.Show("WWP BIM Tools", "ui.xaml not found. Ensure ui.xaml is next to script.cs.");
            return;
        }

        new PrintFromExcelController(uiapp, window);
        window.ShowDialog();
    }
}

internal static class UiLoader
{
    public static Window LoadWindow(string fileName)
    {
        string xamlPath = FindXamlPath(fileName);
        if (string.IsNullOrEmpty(xamlPath) || !File.Exists(xamlPath))
            return null;

        using (FileStream stream = new FileStream(xamlPath, FileMode.Open, FileAccess.Read))
        {
            return (Window)XamlReader.Load(stream);
        }
    }

    private static string FindXamlPath(string fileName)
    {
        List<string> candidates = new List<string>();

        try
        {
            candidates.Add(Path.Combine(Environment.CurrentDirectory, fileName));
        }
        catch { }

        try
        {
            string asmDir = Path.GetDirectoryName(typeof(UiLoader).Assembly.Location);
            if (!string.IsNullOrEmpty(asmDir))
                candidates.Add(Path.Combine(asmDir, fileName));
        }
        catch { }

        try
        {
            string procDir = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
            if (!string.IsNullOrEmpty(procDir))
                candidates.Add(Path.Combine(procDir, fileName));
        }
        catch { }

        foreach (string candidate in candidates.Distinct())
        {
            if (File.Exists(candidate))
                return candidate;
        }

        try
        {
            string cwd = Environment.CurrentDirectory;
            for (int i = 0; i < 6; i++)
            {
                string guess = Path.Combine(
                    cwd,
                    "WWP_Revit_BIM-Tools.extension",
                    "WWP BIM Tools.tab",
                    "Cloud.panel",
                    "Print Sheets from Excel.pushbutton",
                    fileName);
                if (File.Exists(guess))
                    return guess;

                DirectoryInfo parent = Directory.GetParent(cwd);
                if (parent == null)
                    break;
                cwd = parent.FullName;
            }
        }
        catch { }

        return null;
    }
}
internal class PrintFromExcelController
{
    private readonly UIApplication _uiapp;
    private readonly Window _window;

    private readonly TextBlock _docTitleText;
    private readonly TextBox _excelPathBox;
    private readonly Button _browseExcelButton;
    private readonly Button _generateExcelButton;
    private readonly TextBox _outputPathBox;
    private readonly Button _browseOutputButton;
    private readonly ToggleButton _pdfToggle;
    private readonly ToggleButton _dwgToggle;
    private readonly Button _printNowButton;
    private readonly DatePicker _scheduleDatePicker;
    private readonly TextBox _scheduleTimeBox;
    private readonly Button _scheduleButton;
    private readonly Button _cancelScheduleButton;
    private readonly TextBlock _scheduleStatusText;
    private readonly TextBox _logBox;

    public PrintFromExcelController(UIApplication uiapp, Window window)
    {
        _uiapp = uiapp;
        _window = window;

        _docTitleText = (TextBlock)window.FindName("DocTitleText");
        _excelPathBox = (TextBox)window.FindName("ExcelPathBox");
        _browseExcelButton = (Button)window.FindName("BrowseExcelButton");
        _generateExcelButton = (Button)window.FindName("GenerateExcelButton");
        _outputPathBox = (TextBox)window.FindName("OutputPathBox");
        _browseOutputButton = (Button)window.FindName("BrowseOutputButton");
        _pdfToggle = (ToggleButton)window.FindName("PdfToggle");
        _dwgToggle = (ToggleButton)window.FindName("DwgToggle");
        _printNowButton = (Button)window.FindName("PrintNowButton");
        _scheduleDatePicker = (DatePicker)window.FindName("ScheduleDatePicker");
        _scheduleTimeBox = (TextBox)window.FindName("ScheduleTimeBox");
        _scheduleButton = (Button)window.FindName("ScheduleButton");
        _cancelScheduleButton = (Button)window.FindName("CancelScheduleButton");
        _scheduleStatusText = (TextBlock)window.FindName("ScheduleStatusText");
        _logBox = (TextBox)window.FindName("LogBox");

        _browseExcelButton.Click += (s, e) => BrowseExcel();
        _generateExcelButton.Click += (s, e) => GenerateOrUpdateExcel();
        _browseOutputButton.Click += (s, e) => BrowseOutputFolder();
        _printNowButton.Click += (s, e) => PrintNow();
        _scheduleButton.Click += (s, e) => SchedulePrint();
        _cancelScheduleButton.Click += (s, e) => CancelSchedule();
        _window.Closing += OnWindowClosing;

        _scheduleDatePicker.SelectedDate = DateTime.Today;
        _scheduleTimeBox.Text = DateTime.Now.AddHours(1).ToString("HH:mm");
        _scheduleStatusText.Text = "No schedule";

        _docTitleText.Text = GetDocumentTitle();
        if (string.IsNullOrWhiteSpace(_outputPathBox.Text))
            _outputPathBox.Text = DefaultOutputFolder();

        PrintScheduler.Initialize(_uiapp, Log);
        UpdateScheduleStatus();

        Log("Ready.");
    }

    private string GetDocumentTitle()
    {
        Document doc = GetDocument();
        return doc != null ? doc.Title : "No active document";
    }

    private Document GetDocument()
    {
        if (_uiapp == null || _uiapp.ActiveUIDocument == null)
            return null;
        return _uiapp.ActiveUIDocument.Document;
    }

    private void BrowseExcel()
    {
        OpenFileDialog dialog = new OpenFileDialog();
        dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|Excel Files (*.xls)|*.xls";
        dialog.Multiselect = false;

        if (dialog.ShowDialog() != true)
            return;

        _excelPathBox.Text = dialog.FileName;
        Log("Excel selected: " + dialog.FileName);
    }

    private void GenerateOrUpdateExcel()
    {
        Document doc = GetDocument();
        if (doc == null)
        {
            TaskDialog.Show("WWP BIM Tools", "No active document.");
            return;
        }

        string path = _excelPathBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(path))
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
            dialog.FileName = "PrintDatabase.xlsx";
            if (dialog.ShowDialog() != true)
                return;
            path = dialog.FileName;
            _excelPathBox.Text = path;
        }

        try
        {
            IList<ViewSheet> sheets = SheetUtils.GetPrintableSheets(doc);
            ExcelDatabase.GenerateOrUpdate(path, sheets, Log);
            Log("Excel updated: " + path);
        }
        catch (Exception ex)
        {
            Log("Excel update failed: " + ex.Message);
            TaskDialog.Show("WWP BIM Tools", "Excel update failed: " + ex.Message);
        }
    }

    private void BrowseOutputFolder()
    {
        string selected = FolderPicker.PickFolder(_outputPathBox.Text);
        if (string.IsNullOrWhiteSpace(selected))
            return;

        _outputPathBox.Text = selected;
        Log("Output folder: " + selected);
    }

    private void PrintNow()
    {
        if (!ValidateInputs())
            return;

        Document doc = GetDocument();
        if (doc == null)
        {
            TaskDialog.Show("WWP BIM Tools", "No active document.");
            return;
        }

        PrintEngine.Run(
            doc,
            _excelPathBox.Text.Trim(),
            _outputPathBox.Text.Trim(),
            _pdfToggle.IsChecked == true,
            _dwgToggle.IsChecked == true,
            Log);
    }

    private void SchedulePrint()
    {
        if (!ValidateInputs())
            return;

        DateTime runAt;
        if (!TryGetScheduleDate(out runAt))
            return;

        Document doc = GetDocument();
        if (doc == null)
        {
            TaskDialog.Show("WWP BIM Tools", "No active document.");
            return;
        }

        ScheduledJob job = new ScheduledJob(
            doc,
            runAt,
            _excelPathBox.Text.Trim(),
            _outputPathBox.Text.Trim(),
            _pdfToggle.IsChecked == true,
            _dwgToggle.IsChecked == true);

        PrintScheduler.SetJob(job);
        UpdateScheduleStatus();
        Log("Scheduled print at " + runAt.ToString("yyyy-MM-dd HH:mm"));
    }

    private void CancelSchedule()
    {
        PrintScheduler.CancelJob();
        UpdateScheduleStatus();
    }

    private bool ValidateInputs()
    {
        if (string.IsNullOrWhiteSpace(_excelPathBox.Text))
        {
            TaskDialog.Show("WWP BIM Tools", "Select or generate an Excel file first.");
            return false;
        }

        if (string.IsNullOrWhiteSpace(_outputPathBox.Text))
        {
            TaskDialog.Show("WWP BIM Tools", "Select an output folder.");
            return false;
        }

        if (_pdfToggle.IsChecked != true && _dwgToggle.IsChecked != true)
        {
            TaskDialog.Show("WWP BIM Tools", "Enable PDF and/or DWG export.");
            return false;
        }

        return true;
    }

    private bool TryGetScheduleDate(out DateTime runAt)
    {
        runAt = DateTime.MinValue;
        DateTime? selectedDate = _scheduleDatePicker.SelectedDate;
        if (!selectedDate.HasValue)
        {
            TaskDialog.Show("WWP BIM Tools", "Select a schedule date.");
            return false;
        }

        TimeSpan time;
        if (!TimeSpan.TryParse(_scheduleTimeBox.Text.Trim(), out time))
        {
            TaskDialog.Show("WWP BIM Tools", "Enter time in HH:MM (24h) format.");
            return false;
        }

        runAt = selectedDate.Value.Date.Add(time);
        if (runAt <= DateTime.Now.AddMinutes(1))
        {
            TaskDialog.Show("WWP BIM Tools", "Scheduled time must be in the future.");
            return false;
        }

        return true;
    }

    private void UpdateScheduleStatus()
    {
        if (PrintScheduler.HasJob)
        {
            ScheduledJob job = PrintScheduler.CurrentJob;
            _scheduleStatusText.Text = "Scheduled for " + job.RunAt.ToString("yyyy-MM-dd HH:mm");
        }
        else
        {
            _scheduleStatusText.Text = "No schedule";
        }
    }

    private void OnWindowClosing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        if (!PrintScheduler.HasJob)
            return;

        MessageBoxResult res = MessageBox.Show(
            "A scheduled print is still pending. Cancel the schedule and close this window?",
            "WWP BIM Tools",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (res == MessageBoxResult.No)
        {
            e.Cancel = true;
            return;
        }

        PrintScheduler.CancelJob();
        UpdateScheduleStatus();
    }

    private string DefaultOutputFolder()
    {
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "WWP Print Folder");
    }

    private void Log(string message)
    {
        if (_logBox == null)
            return;
        _logBox.AppendText(DateTime.Now.ToString("HH:mm:ss") + "  " + message + Environment.NewLine);
        _logBox.ScrollToEnd();
    }
}
internal static class PrintScheduler
{
    private static UIApplication _uiapp;
    private static ScheduledJob _job;
    private static Action<string> _log;
    private static bool _initialized;

    public static bool HasJob { get { return _job != null; } }
    public static ScheduledJob CurrentJob { get { return _job; } }

    public static void Initialize(UIApplication uiapp, Action<string> log)
    {
        if (_initialized)
            return;

        _uiapp = uiapp;
        _log = log;
        if (_uiapp != null)
        {
            _uiapp.Idling += OnIdling;
            _uiapp.Application.ApplicationClosing += OnApplicationClosing;
        }

        _initialized = true;
    }

    public static void SetJob(ScheduledJob job)
    {
        _job = job;
    }

    public static void CancelJob()
    {
        if (_job == null)
            return;
        _log?.Invoke("Schedule canceled.");
        _job = null;
    }

    private static void OnIdling(object sender, IdlingEventArgs e)
    {
        if (_job == null)
            return;
        if (_job.IsRunning)
            return;
        if (DateTime.Now < _job.RunAt)
            return;

        _job.IsRunning = true;
        try
        {
            _job.Execute(_log);
        }
        catch (Exception ex)
        {
            _log?.Invoke("Scheduled print failed: " + ex.Message);
        }
        finally
        {
            _job = null;
        }
    }

    private static void OnApplicationClosing(object sender, ApplicationClosingEventArgs e)
    {
        if (_job == null)
            return;

        TaskDialogResult result = TaskDialog.Show(
            "WWP BIM Tools",
            "A scheduled print is pending at " + _job.RunAt.ToString("yyyy-MM-dd HH:mm") + ".\n" +
            "Close Revit and cancel the scheduled print?",
            TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

        if (result == TaskDialogResult.No)
        {
            e.Cancel = true;
            _log?.Invoke("Close canceled. Schedule still active.");
        }
        else
        {
            _log?.Invoke("Schedule canceled because Revit is closing.");
            _job = null;
        }
    }
}

internal class ScheduledJob
{
    public DateTime RunAt { get; private set; }
    public bool IsRunning { get; set; }
    public Document DocumentRef { get; private set; }
    public string ExcelPath { get; private set; }
    public string OutputFolder { get; private set; }
    public bool ExportPdf { get; private set; }
    public bool ExportDwg { get; private set; }

    public ScheduledJob(Document doc, DateTime runAt, string excelPath, string outputFolder, bool exportPdf, bool exportDwg)
    {
        DocumentRef = doc;
        RunAt = runAt;
        ExcelPath = excelPath;
        OutputFolder = outputFolder;
        ExportPdf = exportPdf;
        ExportDwg = exportDwg;
    }

    public void Execute(Action<string> log)
    {
        if (DocumentRef == null || !DocumentRef.IsValidObject)
        {
            log?.Invoke("Scheduled print failed: document is not available.");
            return;
        }

        PrintEngine.Run(DocumentRef, ExcelPath, OutputFolder, ExportPdf, ExportDwg, log);
    }
}
internal static class SheetUtils
{
    public static IList<ViewSheet> GetPrintableSheets(Document doc)
    {
        return new FilteredElementCollector(doc)
            .OfClass(typeof(ViewSheet))
            .Cast<ViewSheet>()
            .Where(sheet => sheet != null && !sheet.IsPlaceholder && sheet.CanBePrinted)
            .OrderBy(sheet => sheet.SheetNumber)
            .ToList();
    }
}

internal static class PrintEngine
{
    public static void Run(Document doc, string excelPath, string outputFolder, bool exportPdf, bool exportDwg, Action<string> log)
    {
        if (doc == null)
        {
            log?.Invoke("No document to print.");
            return;
        }

        if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
        {
            log?.Invoke("Excel file not found.");
            return;
        }

        if (string.IsNullOrWhiteSpace(outputFolder))
        {
            log?.Invoke("Output folder not set.");
            return;
        }

        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        List<ExcelPrintRow> rows = ExcelDatabase.ReadPrintRows(excelPath);
        if (rows.Count == 0)
        {
            log?.Invoke("No rows found in Excel.");
            return;
        }

        IList<ViewSheet> sheets = SheetUtils.GetPrintableSheets(doc);
        Dictionary<string, List<ViewSheet>> sheetGroups = new Dictionary<string, List<ViewSheet>>(StringComparer.OrdinalIgnoreCase);
        foreach (ViewSheet sheet in sheets)
        {
            if (!sheetGroups.ContainsKey(sheet.Name))
                sheetGroups[sheet.Name] = new List<ViewSheet>();
            sheetGroups[sheet.Name].Add(sheet);
        }

        HashSet<string> duplicates = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        Dictionary<string, ViewSheet> sheetByName = new Dictionary<string, ViewSheet>(StringComparer.OrdinalIgnoreCase);
        foreach (KeyValuePair<string, List<ViewSheet>> pair in sheetGroups)
        {
            if (pair.Value.Count > 1)
            {
                duplicates.Add(pair.Key);
                log?.Invoke("Duplicate sheet name found: " + pair.Key + ". Skipping those rows.");
            }
            else
            {
                sheetByName[pair.Key] = pair.Value[0];
            }
        }

        Dictionary<string, int> usedNames = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        int exported = 0;
        int skipped = 0;

        foreach (ExcelPrintRow row in rows)
        {
            if (string.IsNullOrWhiteSpace(row.DrawingName))
                continue;

            if (duplicates.Contains(row.DrawingName))
            {
                skipped++;
                log?.Invoke("Skipped duplicate sheet name: " + row.DrawingName);
                continue;
            }

            ViewSheet sheet;
            if (!sheetByName.TryGetValue(row.DrawingName, out sheet))
            {
                skipped++;
                log?.Invoke("Sheet not found: " + row.DrawingName);
                continue;
            }

            string baseName = row.PrintFileName;
            if (string.IsNullOrWhiteSpace(baseName))
                baseName = sheet.SheetNumber + "_" + sheet.Name;
            baseName = FileNameUtils.SanitizeFileName(baseName);
            baseName = FileNameUtils.EnsureUnique(baseName, usedNames);

            if (exportPdf)
            {
                try
                {
                    ExportPdf(doc, sheet, outputFolder, baseName);
                    exported++;
                    log?.Invoke("PDF exported: " + baseName + ".pdf");
                }
                catch (Exception ex)
                {
                    log?.Invoke("PDF export failed for " + sheet.Name + ": " + ex.Message);
                }
            }

            if (exportDwg)
            {
                try
                {
                    ExportDwg(doc, sheet, outputFolder, baseName);
                    exported++;
                    log?.Invoke("DWG exported: " + baseName + ".dwg");
                }
                catch (Exception ex)
                {
                    log?.Invoke("DWG export failed for " + sheet.Name + ": " + ex.Message);
                }
            }
        }

        log?.Invoke("Export finished. Exported: " + exported + ", Skipped: " + skipped);
    }

    private static void ExportPdf(Document doc, ViewSheet sheet, string outputFolder, string baseName)
    {
        PDFExportOptions options = new PDFExportOptions();
        options.FileName = baseName;
        IList<ElementId> ids = new List<ElementId> { sheet.Id };
        doc.Export(outputFolder, ids, options);
    }

    private static void ExportDwg(Document doc, ViewSheet sheet, string outputFolder, string baseName)
    {
        DWGExportOptions options = new DWGExportOptions();
        IList<ElementId> ids = new List<ElementId> { sheet.Id };
        doc.Export(outputFolder, baseName + ".dwg", ids, options);
    }
}

internal static class FileNameUtils
{
    public static string SanitizeFileName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return "Sheet";

        foreach (char c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');

        return name.Trim();
    }

    public static string EnsureUnique(string baseName, Dictionary<string, int> used)
    {
        int count;
        if (!used.TryGetValue(baseName, out count))
        {
            used[baseName] = 1;
            return baseName;
        }

        count++;
        used[baseName] = count;
        return baseName + "_" + count.ToString();
    }
}
internal class ExcelPrintRow
{
    public string PrintFileName { get; set; }
    public string DrawingName { get; set; }
    public string DrawingNumber { get; set; }
}

internal static class ExcelDatabase
{
    private const string HeaderFileName = "Printed File Name";
    private const string HeaderDrawingName = "Drawing Name";
    private const string HeaderDrawingNumber = "Drawing Number";

    public static void GenerateOrUpdate(string path, IList<ViewSheet> sheets, Action<string> log)
    {
        dynamic excel = null;
        dynamic workbook = null;
        dynamic sheet = null;
        dynamic usedRange = null;

        try
        {
            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
                throw new InvalidOperationException("Excel is not installed.");

            excel = Activator.CreateInstance(excelType);
            excel.Visible = false;

            if (File.Exists(path))
                workbook = excel.Workbooks.Open(path, ReadOnly: false);
            else
                workbook = excel.Workbooks.Add();

            sheet = workbook.Worksheets[1];
            usedRange = sheet.UsedRange;

            int colFileName = 1;
            int colDrawingName = 2;
            int colDrawingNumber = 3;
            EnsureHeaders(sheet, ref colFileName, ref colDrawingName, ref colDrawingNumber);

            Dictionary<string, int> rowByDrawingName = ReadExistingRows(sheet, colDrawingName, usedRange);
            int lastRow = Math.Max(2, rowByDrawingName.Count + 2);

            foreach (ViewSheet viewSheet in sheets)
            {
                string drawingName = viewSheet.Name;
                string drawingNumber = viewSheet.SheetNumber;
                string defaultFileName = FileNameUtils.SanitizeFileName(drawingNumber + "_" + drawingName);

                int rowIndex;
                if (rowByDrawingName.TryGetValue(drawingName, out rowIndex))
                {
                    object currentFileName = sheet.Cells[rowIndex, colFileName].Value2;
                    if (currentFileName == null || string.IsNullOrWhiteSpace(currentFileName.ToString()))
                        sheet.Cells[rowIndex, colFileName].Value2 = defaultFileName;

                    sheet.Cells[rowIndex, colDrawingName].Value2 = drawingName;
                    sheet.Cells[rowIndex, colDrawingNumber].Value2 = drawingNumber;
                }
                else
                {
                    sheet.Cells[lastRow, colFileName].Value2 = defaultFileName;
                    sheet.Cells[lastRow, colDrawingName].Value2 = drawingName;
                    sheet.Cells[lastRow, colDrawingNumber].Value2 = drawingNumber;
                    rowByDrawingName[drawingName] = lastRow;
                    lastRow++;
                }
            }

            if (File.Exists(path))
                workbook.Save();
            else
                workbook.SaveAs(path);
        }
        finally
        {
            if (workbook != null)
                workbook.Close(false);
            if (excel != null)
                excel.Quit();

            ReleaseComObject(usedRange);
            ReleaseComObject(sheet);
            ReleaseComObject(workbook);
            ReleaseComObject(excel);
        }
    }

    public static List<ExcelPrintRow> ReadPrintRows(string path)
    {
        List<ExcelPrintRow> result = new List<ExcelPrintRow>();

        dynamic excel = null;
        dynamic workbook = null;
        dynamic sheet = null;
        dynamic usedRange = null;

        try
        {
            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
                throw new InvalidOperationException("Excel is not installed.");

            excel = Activator.CreateInstance(excelType);
            excel.Visible = false;

            workbook = excel.Workbooks.Open(path, ReadOnly: true);
            sheet = workbook.Worksheets[1];
            usedRange = sheet.UsedRange;
            object[,] values = usedRange.Value2 as object[,];
            if (values == null)
                return result;

            int rows = values.GetLength(0);
            int cols = values.GetLength(1);

            Dictionary<string, int> headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int col = 1; col <= cols; col++)
            {
                object headerVal = values[1, col];
                if (headerVal == null)
                    continue;
                string header = headerVal.ToString().Trim();
                if (string.IsNullOrWhiteSpace(header))
                    continue;
                if (!headers.ContainsKey(header))
                    headers.Add(header, col);
            }

            int colFileName = headers.ContainsKey(HeaderFileName) ? headers[HeaderFileName] : 1;
            int colDrawingName = headers.ContainsKey(HeaderDrawingName) ? headers[HeaderDrawingName] : 2;
            int colDrawingNumber = headers.ContainsKey(HeaderDrawingNumber) ? headers[HeaderDrawingNumber] : 3;

            for (int row = 2; row <= rows; row++)
            {
                object drawingNameVal = values[row, colDrawingName];
                if (drawingNameVal == null)
                    continue;

                string drawingName = drawingNameVal.ToString().Trim();
                if (string.IsNullOrWhiteSpace(drawingName))
                    continue;

                string fileName = string.Empty;
                object fileNameVal = values[row, colFileName];
                if (fileNameVal != null)
                    fileName = fileNameVal.ToString().Trim();

                string drawingNumber = string.Empty;
                object drawingNumberVal = values[row, colDrawingNumber];
                if (drawingNumberVal != null)
                    drawingNumber = drawingNumberVal.ToString().Trim();

                result.Add(new ExcelPrintRow
                {
                    PrintFileName = fileName,
                    DrawingName = drawingName,
                    DrawingNumber = drawingNumber
                });
            }

            return result;
        }
        finally
        {
            if (workbook != null)
                workbook.Close(false);
            if (excel != null)
                excel.Quit();

            ReleaseComObject(usedRange);
            ReleaseComObject(sheet);
            ReleaseComObject(workbook);
            ReleaseComObject(excel);
        }
    }

    private static Dictionary<string, int> ReadExistingRows(dynamic sheet, int colDrawingName, dynamic usedRange)
    {
        Dictionary<string, int> rows = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        object[,] values = usedRange.Value2 as object[,];
        if (values == null)
            return rows;

        int rowCount = values.GetLength(0);
        for (int row = 2; row <= rowCount; row++)
        {
            object nameVal = values[row, colDrawingName];
            if (nameVal == null)
                continue;
            string name = nameVal.ToString().Trim();
            if (string.IsNullOrWhiteSpace(name))
                continue;
            if (!rows.ContainsKey(name))
                rows.Add(name, row);
        }

        return rows;
    }

    private static void EnsureHeaders(dynamic sheet, ref int colFileName, ref int colDrawingName, ref int colDrawingNumber)
    {
        colFileName = 1;
        colDrawingName = 2;
        colDrawingNumber = 3;

        sheet.Cells[1, colFileName].Value2 = HeaderFileName;
        sheet.Cells[1, colDrawingName].Value2 = HeaderDrawingName;
        sheet.Cells[1, colDrawingNumber].Value2 = HeaderDrawingNumber;
    }

    private static void ReleaseComObject(object obj)
    {
        try
        {
            if (obj != null && Marshal.IsComObject(obj))
                Marshal.FinalReleaseComObject(obj);
        }
        catch { }
    }
}

internal static class FolderPicker
{
    public static string PickFolder(string initialPath)
    {
        string picked = TryCommonDialog(initialPath);
        if (!string.IsNullOrWhiteSpace(picked))
            return picked;

        OpenFileDialog ofd = new OpenFileDialog();
        ofd.CheckFileExists = false;
        ofd.CheckPathExists = true;
        ofd.ValidateNames = false;
        ofd.FileName = "Select Folder";
        if (!string.IsNullOrWhiteSpace(initialPath))
            ofd.InitialDirectory = initialPath;

        if (ofd.ShowDialog() == true)
            return Path.GetDirectoryName(ofd.FileName);

        return null;
    }

    private static string TryCommonDialog(string initialPath)
    {
        try
        {
            Type dialogType = Type.GetType("Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog, Microsoft.WindowsAPICodePack");
            if (dialogType == null)
                return null;

            dynamic dialog = Activator.CreateInstance(dialogType);
            dialog.IsFolderPicker = true;
            dialog.Multiselect = false;
            dialog.Title = "Select Output Folder";
            if (!string.IsNullOrWhiteSpace(initialPath))
                dialog.InitialDirectory = initialPath;

            bool ok = dialog.ShowDialog() == 1;
            if (ok)
                return dialog.FileName as string;
        }
        catch { }

        return null;
    }
}
