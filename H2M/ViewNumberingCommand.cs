using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace H2M
{
    [Transaction(TransactionMode.Manual)]
    public class ViewNumberingCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var telemetry = new TelemetryService();
            telemetry.TrackEvent("ViewNumbering", "button_click");
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Show sheet selection dialog
            TaskDialog sheetSelectionDialog = new TaskDialog("Sheet Selection");
            sheetSelectionDialog.MainInstruction = "Select sheets for renumbering";
            sheetSelectionDialog.MainContent = "Choose which sheets you want to renumber.";
            sheetSelectionDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Select All Sheets");
            sheetSelectionDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Select Individual Sheets");
            sheetSelectionDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Select Current Sheet");
            sheetSelectionDialog.CommonButtons = TaskDialogCommonButtons.Cancel;

            TaskDialogResult result = sheetSelectionDialog.Show();

            List<ViewSheet> selectedSheets = new List<ViewSheet>();

            if (result == TaskDialogResult.CommandLink1)
            {
                selectedSheets = GetAllSheets(doc);
            }
            else if (result == TaskDialogResult.CommandLink2)
            {
                selectedSheets = SelectIndividualSheets(doc);
                if (selectedSheets == null || selectedSheets.Count == 0)
                {
                    TaskDialog.Show("No Sheets Selected", "No sheets were selected for renumbering.");
                    return Result.Cancelled;
                }
            }
            else if (result == TaskDialogResult.CommandLink3)
            {
                Autodesk.Revit.DB.View activeView = uidoc.ActiveView;
                if (activeView is ViewSheet currentSheet)
                {
                    selectedSheets.Add(currentSheet);
                }
                else
                {
                    TaskDialog.Show("Invalid Selection", "The current view is not a sheet.");
                    return Result.Cancelled;
                }
            }
            else
            {
                return Result.Cancelled;
            }

            int totalSheets = selectedSheets.Count;
            int totalViewsRenumbered = 0;

            // First Transaction: Assign Temporary Numbers for all except excluded views
            using (Transaction trans = new Transaction(doc, "Assign Temporary Numbers"))
            {
                trans.Start();
                foreach (ViewSheet sheet in selectedSheets)
                {
                    AssignTemporaryNumbers(doc, sheet);
                }
                trans.Commit();
            }

            // Second Transaction: Assign Final Numbers with rules applied
            using (Transaction trans = new Transaction(doc, "Assign Final Numbers"))
            {
                trans.Start();
                foreach (ViewSheet sheet in selectedSheets)
                {
                    totalViewsRenumbered += AssignFinalNumbers(doc, sheet);
                }
                trans.Commit();
            }

            TaskDialog.Show("Detail Renumbering Complete",
                $"Processed {totalSheets} sheets.\nRenumbered {totalViewsRenumbered} views in total.");

            telemetry.TrackEvent("ViewNumbering", "task_completed", new Dictionary<string, object>
            {
                ["sheets_processed"]  = totalSheets,
                ["views_renumbered"]  = totalViewsRenumbered
            });
            return Result.Succeeded;
        }

        private List<ViewSheet> GetAllSheets(Document doc)
        {
            FilteredElementCollector sheetCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet));

            return sheetCollector.Cast<ViewSheet>().ToList();
        }

        private List<ViewSheet> SelectIndividualSheets(Document doc)
        {
            List<ViewSheet> selectedSheets = new List<ViewSheet>();

            using (System.Windows.Forms.Form sheetSelectionForm = new System.Windows.Forms.Form())
            {
                sheetSelectionForm.Text = "Select Sheets";
                sheetSelectionForm.Size = new System.Drawing.Size(300, 400);

                CheckedListBox sheetList = new CheckedListBox();
                sheetList.Dock = DockStyle.Fill;
                sheetList.CheckOnClick = true;

                List<ViewSheet> sortedSheets = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .OrderBy(sheet => sheet.SheetNumber)
                    .ToList();

                Dictionary<string, ViewSheet> sheetDict = new Dictionary<string, ViewSheet>();

                foreach (ViewSheet sheet in sortedSheets)
                {
                    string displayName = $"{sheet.SheetNumber} - {sheet.Name}";
                    sheetList.Items.Add(displayName, false);
                    sheetDict[displayName] = sheet;
                }

                Button okButton = new Button();
                okButton.Text = "OK";
                okButton.DialogResult = DialogResult.OK;
                okButton.Dock = DockStyle.Bottom;

                Button cancelButton = new Button();
                cancelButton.Text = "Cancel";
                cancelButton.DialogResult = DialogResult.Cancel;
                cancelButton.Dock = DockStyle.Bottom;

                System.Windows.Forms.Panel buttonPanel = new System.Windows.Forms.Panel();
                buttonPanel.Dock = DockStyle.Bottom;
                buttonPanel.Controls.Add(okButton);
                buttonPanel.Controls.Add(cancelButton);

                sheetSelectionForm.Controls.Add(sheetList);
                sheetSelectionForm.Controls.Add(buttonPanel);

                if (sheetSelectionForm.ShowDialog() == DialogResult.OK)
                {
                    for (int i = 0; i < sheetList.Items.Count; i++)
                    {
                        if (sheetList.GetItemChecked(i))
                        {
                            string displayName = (string)sheetList.Items[i];
                            selectedSheets.Add(sheetDict[displayName]);
                        }
                    }
                }
            }

            return selectedSheets;
        }

        private bool IsNoTitleViewport(Viewport viewport, Document doc)
        {
            if (viewport == null) return false;
            ElementId typeId = viewport.GetTypeId();
            Element typeElem = doc.GetElement(typeId);
            string name = typeElem?.Name ?? "";
            // Case-insensitive, substring check ignoring superfluous text
            return name.IndexOf("no title", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void AssignTemporaryNumbers(Document doc, ViewSheet sheet)
        {
            ICollection<ElementId> viewIdSet = sheet.GetAllPlacedViews();
            int tempNumber = 1000;

            var nonLegendViewIds = viewIdSet.Where(viewId =>
            {
                Autodesk.Revit.DB.View view = doc.GetElement(viewId) as Autodesk.Revit.DB.View;
                if (view == null) return false;

                // Exclude legends fully
                if (view.ViewType == ViewType.Legend) return false;
                Category category = view.Category;
                if (category != null && category.Name == "Legends") return false;

                return true;
            }).ToList();

            foreach (ElementId viewId in nonLegendViewIds)
            {
                Viewport viewport = new FilteredElementCollector(doc)
                    .OfClass(typeof(Viewport))
                    .Cast<Viewport>()
                    .FirstOrDefault(vp => vp.ViewId == viewId);

                if (viewport != null)
                {
                    bool isNoTitle = IsNoTitleViewport(viewport, doc);

                    // Assign temporary numbers to all views except those detail number starts with a letter other than 'X'
                    Parameter detailNumber = viewport.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER);
                    if (detailNumber != null && !detailNumber.IsReadOnly)
                    {
                        string currentDetail = detailNumber.AsString() ?? "";

                        if (isNoTitle || !StartsWithExcludedLetter(currentDetail))
                        {
                            detailNumber.Set(tempNumber.ToString());
                            tempNumber++;
                        }
                    }
                }
            }
        }

        private int AssignFinalNumbers(Document doc, ViewSheet sheet)
        {
            int viewsRenumbered = 0;
            try
            {
                ICollection<ElementId> viewIdSet = sheet.GetAllPlacedViews();

                var nonLegendViewIds = viewIdSet.Where(viewId =>
                {
                    Autodesk.Revit.DB.View view = doc.GetElement(viewId) as Autodesk.Revit.DB.View;
                    if (view == null) return false;

                    if (view.ViewType == ViewType.Legend) return false;

                    Category category = view.Category;
                    if (category != null && category.Name == "Legends") return false;

                    return true;
                }).ToList();

                List<Viewport> viewports = new List<Viewport>();

                foreach (ElementId viewId in nonLegendViewIds)
                {
                    Viewport viewport = new FilteredElementCollector(doc)
                        .OfClass(typeof(Viewport))
                        .Cast<Viewport>()
                        .FirstOrDefault(vp => vp.ViewId == viewId);

                    if (viewport != null)
                    {
                        viewports.Add(viewport);
                    }
                }

                if (viewports.Count == 0) return 0;

                // Separate No Title viewports for special processing
                var normalViewports = new List<Viewport>();
                var noTitleViewports = new List<Viewport>();

                foreach (var vp in viewports)
                {
                    if (IsNoTitleViewport(vp, doc))
                    {
                        noTitleViewports.Add(vp);
                    }
                    else
                    {
                        normalViewports.Add(vp);
                    }
                }

                // Process normal viewports with ordering and filtering
                var filteredNormalVPs = normalViewports.Where(vp =>
                {
                    Parameter detailNumber = vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER);
                    if (detailNumber == null) return false;
                    string detailVal = detailNumber.AsString() ?? "";
                    // Exclude if detail number starts with a letter other than 'X'
                    return !StartsWithExcludedLetter(detailVal);
                }).ToList();

                // Sort normal viewports by top-left position
                var viewportsWithPositions = filteredNormalVPs.Select(vp =>
                {
                    Outline boxOutline = vp.GetBoxOutline();
                    XYZ topLeft = new XYZ(boxOutline.MinimumPoint.X, boxOutline.MaximumPoint.Y, 0);
                    return new Tuple<Viewport, XYZ>(vp, topLeft);
                }).ToList();

                viewportsWithPositions.Sort((a, b) =>
                {
                    int yCompare = b.Item2.Y.CompareTo(a.Item2.Y);
                    if (yCompare != 0) return yCompare;
                    return a.Item2.X.CompareTo(b.Item2.X);
                });

                // Group into rows (3 inch tolerance)
                double rowTolerance = 3.0 / 12.0;
                List<List<Tuple<Viewport, XYZ>>> rows = new List<List<Tuple<Viewport, XYZ>>>();
                List<Tuple<Viewport, XYZ>> currentRow = null;

                foreach (var vpTuple in viewportsWithPositions)
                {
                    if (currentRow == null ||
                        Math.Abs(currentRow[0].Item2.Y - vpTuple.Item2.Y) > rowTolerance)
                    {
                        currentRow = new List<Tuple<Viewport, XYZ>>();
                        rows.Add(currentRow);
                    }
                    currentRow.Add(vpTuple);
                }

                int finalNumber = 1;
                // Number normal views top to bottom, left to right
                foreach (var row in rows.OrderByDescending(r => r[0].Item2.Y))
                {
                    foreach (var vpTuple in row.OrderBy(v => v.Item2.X))
                    {
                        Parameter detailNumber = vpTuple.Item1.get_Parameter(
                            BuiltInParameter.VIEWPORT_DETAIL_NUMBER);

                        if (detailNumber != null && !detailNumber.IsReadOnly)
                        {
                            detailNumber.Set(finalNumber.ToString());
                            viewsRenumbered++;
                            finalNumber++;
                        }
                    }
                }

                // Number No Title views with 'X' prefix, order does not matter
                int noTitleNumber = 1;
                foreach (var noTitleVP in noTitleViewports)
                {
                    Parameter detailNumber = noTitleVP.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER);
                    if (detailNumber != null && !detailNumber.IsReadOnly)
                    {
                        detailNumber.Set("X" + noTitleNumber.ToString());
                        viewsRenumbered++;
                        noTitleNumber++;
                    }
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Error processing sheet {sheet.SheetNumber}: {ex.Message}");
            }

            return viewsRenumbered;
        }

        // Helper method to check if detail number starts with a letter other than 'X'
        private bool StartsWithExcludedLetter(string detailNumber)
        {
            if (string.IsNullOrEmpty(detailNumber)) return false;
            char firstChar = detailNumber[0];
            if (char.IsLetter(firstChar) && char.ToUpper(firstChar) != 'X')
            {
                return true;
            }
            return false;
        }
    }
}