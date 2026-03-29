using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace H2M
{
    [Transaction(TransactionMode.Manual)]
    public class SheetCountCommand : IExternalCommand
    {
        private static string _lastExcelFolder = null;

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            var uiApp = commandData.Application;
            var doc = uiApp.ActiveUIDocument.Document;

            string excelPath = BrowseExcelFile();
            if (string.IsNullOrEmpty(excelPath))
            {
                TaskDialog.Show("Info", "No Excel file selected.");
                return Result.Cancelled;
            }

            List<string> excelSheetList;
            try
            {
                excelSheetList = ReadSheetNamesFromExcelInterop(excelPath);
                if (excelSheetList.Count == 0)
                {
                    TaskDialog.Show("Error", "No sheets found in Excel file.");
                    return Result.Failed;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", "Failed to read Excel file: " + ex.Message);
                return Result.Failed;
            }

            var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .ToList();

            // Filter Revit sheets to those in Excel
            var filteredCollector = collector.Where(sheet =>
                !string.IsNullOrEmpty(sheet.SheetNumber) &&
                excelSheetList.Any(excelSheet =>
                    string.Equals(sheet.SheetNumber.Trim(), excelSheet.Trim(), StringComparison.OrdinalIgnoreCase)))
                .ToList();

            // Warn about Revit sheets not in Excel, option to continue or cancel
            var unmatchedSheetsInExcel = collector
                .Where(sheet => !string.IsNullOrEmpty(sheet.SheetNumber))
                .Where(sheet => !excelSheetList.Any(excelSheet =>
                    string.Equals(sheet.SheetNumber.Trim(), excelSheet.Trim(), StringComparison.OrdinalIgnoreCase)))
                .Select(s => $"{s.SheetNumber} - {s.Name}")
                .ToList();

            if (unmatchedSheetsInExcel.Count > 0)
            {
                string list = string.Join("\n", unmatchedSheetsInExcel);

                TaskDialog mismatchDialog = new TaskDialog("Sheet Mismatch Warning")
                {
                    MainInstruction = "The following sheets were found in Revit but not in Excel. Press Ok to ignore these sheets in the sheet count, or press Cancel and contact Project Manager to update sheet list Excel.",
                    MainContent = list,
                    CommonButtons = TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel,
                    DefaultButton = TaskDialogResult.Cancel
                };

                var result = mismatchDialog.Show();

                if (result == TaskDialogResult.Cancel)
                {
                    return Result.Cancelled;
                }
                // If Ok (Continue): process as normal with filtered sheets
            }

            var sheetNumberParamNames = new List<string>
            {
                "Sheet Number Count", "SheetNumberCount", "sheetnumbercount"
            };
            var sheetTotalCountParamNames = new List<string>
            {
                "Sheet Total Project Count", "SheetTotalProjectCount", "sheettotalprojectcount"
            };

            using (Transaction t = new Transaction(doc, "Assign Sheet Numbers"))
            {
                t.Start();

                foreach (var sheet in filteredCollector)
                {
                    int index = excelSheetList.FindIndex(name =>
                        string.Equals(name.Trim(), sheet.SheetNumber.Trim(), StringComparison.OrdinalIgnoreCase));

                    if (index != -1)
                    {
                        int assignedNumber = index + 1; // numbering starts at 1
                        var param = GetParameterByFlexibleName(sheet, sheetNumberParamNames);
                        if (param != null && !param.IsReadOnly)
                        {
                            switch (param.StorageType)
                            {
                                case StorageType.String:
                                    param.Set(assignedNumber.ToString());
                                    break;
                                case StorageType.Integer:
                                    param.Set(assignedNumber);
                                    break;
                            }
                        }
                    }
                }

                // *** FIXED: Use count of Excel sheets for total parameter ***
                int totalSheets = excelSheetList.Count;

                foreach (var sheet in filteredCollector)
                {
                    var totalParam = GetParameterByFlexibleName(sheet, sheetTotalCountParamNames);
                    if (totalParam != null && !totalParam.IsReadOnly)
                    {
                        switch (totalParam.StorageType)
                        {
                            case StorageType.String:
                                totalParam.Set(totalSheets.ToString());
                                break;
                            case StorageType.Integer:
                                totalParam.Set(totalSheets);
                                break;
                        }
                    }
                }

                // Clear sheet count parameters for sheets not in Excel (skipped sheets)
                var skippedSheets = collector.Where(sheet =>
                    !string.IsNullOrEmpty(sheet.SheetNumber) &&
                    !excelSheetList.Any(excelSheet =>
                        string.Equals(sheet.SheetNumber.Trim(), excelSheet.Trim(), StringComparison.OrdinalIgnoreCase)))
                    .ToList();

                foreach (var sheet in skippedSheets)
                {
                    var sheetCountParam = GetParameterByFlexibleName(sheet, sheetNumberParamNames);
                    if (sheetCountParam != null && !sheetCountParam.IsReadOnly)
                    {
                        switch (sheetCountParam.StorageType)
                        {
                            case StorageType.String:
                                sheetCountParam.Set(string.Empty);
                                break;
                            case StorageType.Integer:
                                sheetCountParam.Set(0);
                                break;
                        }
                    }

                    var totalCountParam = GetParameterByFlexibleName(sheet, sheetTotalCountParamNames);
                    if (totalCountParam != null && !totalCountParam.IsReadOnly)
                    {
                        switch (totalCountParam.StorageType)
                        {
                            case StorageType.String:
                                totalCountParam.Set(string.Empty);
                                break;
                            case StorageType.Integer:
                                totalCountParam.Set(0);
                                break;
                        }
                    }
                }

                t.Commit();
            }

            var sheetsInExcelNotInRevit = excelSheetList.Where(excelSheet =>
                !collector.Any(sheet => string.Equals(sheet.SheetNumber.Trim(), excelSheet.Trim(), StringComparison.OrdinalIgnoreCase))
            ).ToList();

            if (sheetsInExcelNotInRevit.Any())
            {
                string errorMessage = "The following sheets are present in Excel but missing in Revit:\n" +
                    string.Join("\n", sheetsInExcelNotInRevit);

                TaskDialog.Show("Sheets in Excel Not in Revit", errorMessage);
            }

            // Finally the existing completion message
            TaskDialog.Show("Done", $"Sheet numbering complete.\nTotal sheets counted in Excel: {excelSheetList.Count}.");
            return Result.Succeeded;
        }

        private string BrowseExcelFile()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title = "Select Excel File with Sheet List",
                InitialDirectory = _lastExcelFolder ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            var result = openFileDialog.ShowDialog();
            if (result == DialogResult.OK && !string.IsNullOrEmpty(openFileDialog.FileName))
            {
                _lastExcelFolder = Path.GetDirectoryName(openFileDialog.FileName);
                return openFileDialog.FileName;
            }
            return null;
        }

        private List<string> ReadSheetNamesFromExcelInterop(string path)
        {
            var sheetNames = new List<string>();
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Open(path);
                worksheet = workbook.Sheets[1] as Excel.Worksheet;

                int row = 2; // Skips header row
                while (true)
                {
                    var cell = worksheet.Cells[row, 1] as Excel.Range;
                    if (cell == null || cell.Value2 == null)
                        break;
                    sheetNames.Add(cell.Value2.ToString().Trim());
                    row++;
                }
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                    workbook = null;
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                    excelApp = null;
                }
                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                    worksheet = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return sheetNames;
        }

        private Parameter GetParameterByFlexibleName(Element element, List<string> possibleNames)
        {
            foreach (Parameter param in element.Parameters)
            {
                string paramName = param.Definition.Name.Replace(" ", "").ToLowerInvariant();
                foreach (var name in possibleNames)
                {
                    if (paramName == name.Replace(" ", "").ToLowerInvariant())
                    {
                        return param;
                    }
                }
            }
            return null;
        }
    }
}
