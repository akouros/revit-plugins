using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Interop;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace H2M
{
    /// <summary>
    /// Revit external command that spell-checks all text elements visible on
    /// every placed sheet in the active document.
    /// <para>
    /// Flow: scan (ProgressDialog) → review (ResultsDialog) → summary (SummaryDialog).
    /// Telemetry is fired on entry and on session completion.
    /// </para>
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SpellCheckCommand : IExternalCommand
    {
        /// <summary>
        /// Accumulates statistics for the current spell-check session.
        /// Passed by reference through ResultsDialog so the dialog can update counts
        /// as the user fixes or skips issues.
        /// </summary>
        public class SpellCheckStats
        {
            /// <summary>Gets or sets the number of sheets scanned.</summary>
            public int SheetsScanned { get; set; }

            /// <summary>Gets or sets the total number of spelling issues found.</summary>
            public int IssuesFound { get; set; }

            /// <summary>Gets or sets the number of issues corrected via the Fix action.</summary>
            public int IssuesFixed { get; set; }

            /// <summary>Gets or sets the number of issues dismissed via the Skip action.</summary>
            public int IssuesSkipped { get; set; }

            /// <summary>Gets or sets the number of words added to the approved-abbreviation list.</summary>
            public int WordsAddedToDictionary { get; set; }

            /// <summary>Gets or sets whether the user cancelled the scan before it completed.</summary>
            public bool Cancelled { get; set; }
        }

        /// <summary>
        /// Entry point called by Revit when the user clicks the Spell Check ribbon button.
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp  = commandData.Application;
            UIDocument    uidoc  = uiApp.ActiveUIDocument;
            Document      doc    = uidoc.Document;

            // ── Telemetry: button click ───────────────────────────────────────
            var telemetry = new TelemetryService();
            telemetry.TrackEvent("SpellCheck", "button_click");

            // ── Guard: at least one sheet must exist ──────────────────────────
            var sheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .OrderBy(s => s.SheetNumber)
                .ToList();

            if (sheets.Count == 0)
            {
                TaskDialog.Show("Spell Check",
                    "No sheets found in the current project.");
                return Result.Succeeded;
            }

            var stats = new SpellCheckStats { SheetsScanned = sheets.Count };

            // ── Spell check engine (load dictionary once, reuse for session) ──
            using (var engine = new SpellCheckEngine())
            {
                // ── Progress dialog: scan runs on background thread ───────────
                var progressDialog = new ProgressDialog(doc, sheets, engine);
                SetRevitOwner(uiApp, progressDialog);
                progressDialog.ShowDialog();

                if (progressDialog.WasCancelled)
                {
                    stats.Cancelled = true;
                    FireCompletedTelemetry(telemetry, stats);
                    return Result.Cancelled;
                }

                var allResults = progressDialog.Results ?? new List<SpellCheckResult>();
                stats.IssuesFound = allResults.Count;

                if (allResults.Count == 0)
                {
                    // No issues — go straight to summary.
                    var noIssuesSummary = new SummaryDialog(stats);
                    SetRevitOwner(uiApp, noIssuesSummary);
                    noIssuesSummary.ShowDialog();
                    FireCompletedTelemetry(telemetry, stats);
                    return Result.Succeeded;
                }

                // ── Results dialog: user reviews and fixes issues ─────────────
                var resultsDialog = new ResultsDialog(doc, uidoc, allResults, engine, stats);
                SetRevitOwner(uiApp, resultsDialog);
                resultsDialog.ShowDialog();
            }

            // ── Summary dialog ────────────────────────────────────────────────
            var summaryDialog = new SummaryDialog(stats);
            SetRevitOwner(uiApp, summaryDialog);
            summaryDialog.ShowDialog();

            FireCompletedTelemetry(telemetry, stats);
            return Result.Succeeded;
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        /// <summary>
        /// Sets the Revit main window as the WPF owner of <paramref name="window"/>
        /// so that the dialog stays on top and is centred correctly.
        /// </summary>
        private static void SetRevitOwner(UIApplication uiApp, Window window)
        {
            try
            {
                var helper = new WindowInteropHelper(window);
                helper.Owner = uiApp.MainWindowHandle;
            }
            catch
            {
                // Non-fatal: dialog will still show, just without a forced owner.
            }
        }

        /// <summary>
        /// Posts the <c>task_completed</c> telemetry event with full session metadata.
        /// </summary>
        private static void FireCompletedTelemetry(TelemetryService telemetry,
            SpellCheckStats stats)
        {
            telemetry.TrackEvent("SpellCheck", "task_completed",
                new Dictionary<string, object>
                {
                    ["sheets_scanned"]           = stats.SheetsScanned,
                    ["issues_found"]             = stats.IssuesFound,
                    ["issues_fixed"]             = stats.IssuesFixed,
                    ["issues_skipped"]           = stats.IssuesSkipped,
                    ["words_added_to_dictionary"]= stats.WordsAddedToDictionary,
                    ["cancelled"]                = stats.Cancelled
                });
        }
    }
}
