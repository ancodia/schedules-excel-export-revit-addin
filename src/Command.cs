using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ClosedXML.Excel;
using ScheduleExporter;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Threading;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;

namespace SchedulesExcelExport {
    [Transaction(TransactionMode.Manual)]
    public partial class Command : IExternalCommand {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements) {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            Settings docSettings = doc.Settings;

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            List<ViewSchedule> schedules = collector
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(schedule => !schedule.Name.Contains("<Revision"))
                .OrderBy(vs => vs.Name)
                .ToList();

            List<string> scheduleNames = schedules
                .Select(schedule => schedule.Name)
                .ToList();

            try {
                ExportWindow exportWindow = new(scheduleNames);
                var result = exportWindow.ShowDialog();

                if (result == true) {
                    string filePath = exportWindow.SelectedFilePath;
                    bool writeAsString = exportWindow.WriteAsString;
                    bool excludeEmptyAllRows = exportWindow.ExcludeEmptyRows;
                    List<string> selectedScheduleNames = exportWindow.SelectedSchedules;


                    List<ViewSchedule> selectedSchedules = schedules
                        .Where(schedule => selectedScheduleNames.Contains(schedule.Name))
                        .ToList();

                    // Show loading window during export
                    LoadingWindow loadingWindow = new LoadingWindow();
                    loadingWindow.Show();

                    // Force window to render
                    loadingWindow.Dispatcher.Invoke(
                        DispatcherPriority.Background,
                        new Action(delegate { })
                    );

                    try {
                        // Extract data from Revit
                        var scheduleDataList = new List<(string name, List<List<object>> data)>();
                        int currentSchedule = 0;
                        int totalSchedules = selectedSchedules.Count;

                        foreach (var schedule in selectedSchedules) {
                            currentSchedule++;
                            loadingWindow.UpdateStatus($"Processing schedule {currentSchedule} of {totalSchedules}: {schedule.Name}");

                            // Force UI update
                            loadingWindow.Dispatcher.Invoke(
                                DispatcherPriority.Background,
                                new Action(delegate { })
                            );

                            string scheduleName = SanitizeWorksheetName(schedule.Name);
                            var data = PrepareScheduleData(schedule, writeAsString, excludeEmptyAllRows);
                            scheduleDataList.Add((scheduleName, data));
                        }

                        // Export to Excel
                        loadingWindow.UpdateStatus("Writing to Excel file...");

                        // Force UI update
                        loadingWindow.Dispatcher.Invoke(
                            DispatcherPriority.Background,
                            new Action(delegate { })
                        );

                        ExportToExcel(scheduleDataList, filePath);

                        loadingWindow.UpdateStatus("Export complete!");

                        // Force UI update
                        loadingWindow.Dispatcher.Invoke(
                            DispatcherPriority.Background,
                            new Action(delegate { })
                        );
                    }
                    finally {
                        // Close loading window
                        loadingWindow.Close();
                    }
                }
            }
            catch (Exception ex) {
                Autodesk.Revit.UI.TaskDialog.Show("Error", ex.Message);
                return Result.Failed;
            }

            return Result.Succeeded;
        }

        public void ExportToExcel(
            List<(string name, List<List<object>> data)> scheduleDataList,
            string filePath
        ) {
            FileInfo excelFile = new(filePath);

            if (excelFile.Exists && Helper.IsFileOpen(filePath)) {
                MessageBox.Show("The Excel file is currently open. Please close it and try again.", "File In Use");
                return;
            }

            XLWorkbook workbook;

            // Load existing workbook or create new one
            if (excelFile.Exists) {
                workbook = new XLWorkbook(filePath);
            }
            else {
                workbook = new XLWorkbook();
            }

            using (workbook) {
                foreach (var scheduleData in scheduleDataList) {
                    string scheduleName = scheduleData.name;
                    var data = scheduleData.data;

                    // Get existing worksheet or add new one
                    IXLWorksheet worksheet;
                    if (workbook.Worksheets.Contains(scheduleName)) {
                        worksheet = workbook.Worksheet(scheduleName);
                        worksheet.Clear(); // Clear existing content
                    }
                    else {
                        worksheet = workbook.Worksheets.Add(scheduleName);
                    }

                    // Insert data starting at cell A1
                    if (data.Count > 0) {
                        worksheet.Cell(1, 1).InsertData(data);
                    }
                }

                // Ensure directory exists
                Directory.CreateDirectory(Path.GetDirectoryName(filePath));

                // Remove the default sheet if it is present
                var defaultSheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == "Sheet1");
                if (defaultSheet != null && workbook.Worksheets.Count > 1) {
                    defaultSheet.Delete();
                }

                workbook.SaveAs(filePath);
            }

            // Open the saved Excel file
            Process.Start(new ProcessStartInfo
            {
                FileName = filePath,
                UseShellExecute = true
            });
        }

        private List<List<object>> PrepareScheduleData(
            ViewSchedule schedule,
            bool numbersAsStrings = false,
            bool excludeAllEmptyRows = false
        ) {
            TableData tableData = schedule.GetTableData();
            TableSectionData sectionData = tableData.GetSectionData(SectionType.Body);

            int rowCount = sectionData.NumberOfRows;
            int colCount = sectionData.NumberOfColumns;

            var data = new List<List<object>>(rowCount);

            for (int row = 0; row < rowCount; row++) {
                if (!IsRowEmpty(schedule, row, colCount, excludeAllEmptyRows)) {
                    var rowData = new List<object>(colCount);

                    for (int col = 0; col < colCount; col++) {
                        string cellValue = schedule.GetCellText(SectionType.Body, row, col).Trim();

                        if (!numbersAsStrings) {
                            // Attempt to convert cell value to numeric types
                            if (double.TryParse(cellValue, out double doubleValue)) {
                                rowData.Add(doubleValue);
                            }
                            else if (int.TryParse(cellValue, out int intValue)) {
                                rowData.Add(intValue);
                            }
                            else {
                                rowData.Add(cellValue);
                            }
                        }
                        else {
                            rowData.Add(cellValue);
                        }
                    }
                    data.Add(rowData);
                }
            }
            return data;
        }


        private static string SanitizeWorksheetName(string name) {
            // Replace characters that are not supported in sheet names
            string sanitized = Regex.Replace(name, @"[:\\/[\]*?]", "");

            // Excel has a 31-character limit for worksheet names
            if (sanitized.Length > 31) {
                sanitized = sanitized.Substring(0, 31);
            }

            return sanitized;
        }

        /// <summary>
        /// By default any empty header rows (i.e. no data in second row of `sectionData`) will be excluded with this method.
        /// Optionally (if `checkAllRows==true`) this method will flag all other rows with no data for exclusion.
        /// </summary>
        private static bool IsRowEmpty(
            ViewSchedule schedule,
            int rowIndex,
            int colCount,
            bool checkAllRows = false
        ) {
            if (checkAllRows || rowIndex == 1) {
                for (int colIndex = 0; colIndex < colCount; colIndex++) {
                    string cellValue = schedule.GetCellText(SectionType.Body, rowIndex, colIndex);

                    if (!string.IsNullOrWhiteSpace(cellValue)) {
                        return false;
                    }
                }
                return true;
            }
            else {
                return false;
            }
        }
    }
}