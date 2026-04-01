using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace H2M
{
    /// <summary>
    /// Represents a single piece of text collected from a Revit sheet element,
    /// prior to spell checking. One item may produce multiple <see cref="SpellCheckResult"/>
    /// instances if the text contains more than one misspelled word.
    /// </summary>
    public class SheetTextItem
    {
        /// <summary>Gets or sets the text string to be spell-checked.</summary>
        public string Text { get; set; }

        /// <summary>Gets or sets the Revit element ID of the owning element.</summary>
        public ElementId ElementId { get; set; }

        /// <summary>Gets or sets the sheet number the element appears on.</summary>
        public string SheetNumber { get; set; }

        /// <summary>Gets or sets the name of the sheet the element appears on.</summary>
        public string SheetName { get; set; }

        /// <summary>
        /// Gets or sets a human-readable element type label shown in the results grid,
        /// e.g. "Text Note", "Title Block", "Legend", "Viewport Title", "Sheet Name".
        /// </summary>
        public string ElementType { get; set; }

        /// <summary>
        /// Gets or sets whether this text can be edited.
        /// <c>false</c> for linked-model elements or read-only parameters.
        /// </summary>
        public bool IsEditable { get; set; }

        /// <summary>
        /// Gets or sets the Revit parameter name for parameter-sourced text.
        /// <c>null</c> for the body text of a <see cref="TextNote"/>.
        /// </summary>
        public string ParameterName { get; set; }
    }

    /// <summary>
    /// Collects text elements that are visible on placed sheets in the active Revit document.
    /// Only elements belonging to the active document are collected; linked-model elements
    /// are explicitly excluded.
    /// </summary>
    public static class SheetTextCollector
    {
        /// <summary>
        /// Collects all text items from every sheet in the document.
        /// Legend TextNote elements that appear on multiple sheets are deduplicated —
        /// they are reported under the first sheet they are encountered on.
        /// </summary>
        /// <param name="doc">The active Revit document.</param>
        /// <returns>
        /// A list of <see cref="SheetTextItem"/> instances ready for spell checking,
        /// ordered by sheet number.
        /// </returns>
        public static List<SheetTextItem> CollectAllSheets(Document doc)
        {
            var results = new List<SheetTextItem>();
            // Track legend element IDs already reported to prevent duplicates across sheets.
            var seenLegendElementIds = new HashSet<ElementId>();

            var sheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .OrderBy(s => s.SheetNumber)
                .ToList();

            foreach (var sheet in sheets)
                CollectFromSheet(doc, sheet, seenLegendElementIds, results);

            return results;
        }

        /// <summary>
        /// Collects all text items from a single sheet and appends them to
        /// <paramref name="accumulator"/>.
        /// </summary>
        /// <param name="doc">The active Revit document.</param>
        /// <param name="sheet">The sheet to scan.</param>
        /// <param name="seenLegendElementIds">
        /// Shared set used to deduplicate legend TextNote elements that appear on
        /// multiple sheets. Pass an empty <see cref="HashSet{T}"/> when scanning
        /// a single sheet in isolation.
        /// </param>
        /// <param name="accumulator">List to which collected items are appended.</param>
        public static void CollectFromSheet(
            Document doc,
            ViewSheet sheet,
            HashSet<ElementId> seenLegendElementIds,
            List<SheetTextItem> accumulator)
        {
            string sheetNum  = sheet.SheetNumber;
            string sheetName = sheet.Name;

            // ── 1. TextNote elements placed directly on the sheet ─────────────
            var sheetTextNotes = new FilteredElementCollector(doc, sheet.Id)
                .OfClass(typeof(TextNote))
                .Cast<TextNote>()
                .ToList();

            foreach (var tn in sheetTextNotes)
            {
                if (tn.Document != doc) continue;   // exclude linked model elements
                string text = tn.Text;
                if (string.IsNullOrWhiteSpace(text)) continue;

                accumulator.Add(new SheetTextItem
                {
                    Text          = text,
                    ElementId     = tn.Id,
                    SheetNumber   = sheetNum,
                    SheetName     = sheetName,
                    ElementType   = "Text Note",
                    IsEditable    = true,
                    ParameterName = null
                });
            }

            // ── 2. Title block string parameters ─────────────────────────────
            var titleBlocks = new FilteredElementCollector(doc, sheet.Id)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .ToList();

            foreach (var tb in titleBlocks)
            {
                if (tb.Document != doc) continue;

                foreach (Parameter param in tb.Parameters)
                {
                    if (param.StorageType != StorageType.String) continue;
                    if (param.IsReadOnly) continue;

                    string val = param.AsString();
                    if (string.IsNullOrWhiteSpace(val)) continue;

                    accumulator.Add(new SheetTextItem
                    {
                        Text          = val,
                        ElementId     = tb.Id,
                        SheetNumber   = sheetNum,
                        SheetName     = sheetName,
                        ElementType   = "Title Block",
                        IsEditable    = true,
                        ParameterName = param.Definition.Name
                    });
                }
            }

            // ── 3. Sheet Name parameter ───────────────────────────────────────
            if (!string.IsNullOrWhiteSpace(sheetName))
            {
                var sheetNameParam = sheet.get_Parameter(BuiltInParameter.SHEET_NAME);
                bool nameEditable  = sheetNameParam != null && !sheetNameParam.IsReadOnly;

                accumulator.Add(new SheetTextItem
                {
                    Text          = sheetName,
                    ElementId     = sheet.Id,
                    SheetNumber   = sheetNum,
                    SheetName     = sheetName,
                    ElementType   = "Sheet Name",
                    IsEditable    = nameEditable,
                    ParameterName = "Sheet Name"
                });
            }

            // ── 4. Viewport titles and legend TextNotes ───────────────────────
            foreach (var viewId in sheet.GetAllPlacedViews())
            {
                var view = doc.GetElement(viewId) as Autodesk.Revit.DB.View;
                if (view == null) continue;

                bool isLegend = view.ViewType == ViewType.Legend;

                // Viewport title (the view name displayed in the viewport title bar)
                string viewName = view.Name;
                if (!string.IsNullOrWhiteSpace(viewName))
                {
                    var viewNameParam = view.get_Parameter(BuiltInParameter.VIEW_NAME);
                    bool nameEditable = viewNameParam != null && !viewNameParam.IsReadOnly
                                        && view.Document == doc;

                    accumulator.Add(new SheetTextItem
                    {
                        Text          = viewName,
                        ElementId     = view.Id,
                        SheetNumber   = sheetNum,
                        SheetName     = sheetName,
                        ElementType   = "Viewport Title",
                        IsEditable    = nameEditable,
                        ParameterName = "View Name"
                    });
                }

                // Legend TextNote elements (only for legend views on this sheet)
                if (!isLegend) continue;

                var legendTextNotes = new FilteredElementCollector(doc, viewId)
                    .OfClass(typeof(TextNote))
                    .Cast<TextNote>()
                    .ToList();

                foreach (var tn in legendTextNotes)
                {
                    if (tn.Document != doc) continue;
                    if (!seenLegendElementIds.Add(tn.Id)) continue; // skip duplicates

                    string text = tn.Text;
                    if (string.IsNullOrWhiteSpace(text)) continue;

                    accumulator.Add(new SheetTextItem
                    {
                        Text          = text,
                        ElementId     = tn.Id,
                        SheetNumber   = sheetNum,
                        SheetName     = sheetName,
                        ElementType   = "Legend",
                        IsEditable    = true,
                        ParameterName = null
                    });
                }
            }
        }
    }
}
