using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Autodesk.Revit.DB;

namespace H2M
{
    /// <summary>
    /// Modal WPF progress dialog displayed while sheets are scanned for text elements.
    /// <para>
    /// The scan (Revit API collection + Hunspell checking) runs on a background
    /// <see cref="Task"/>.  Thread-safe UI updates flow through
    /// <see cref="IProgress{T}"/>.  A <see cref="CancellationToken"/> is wired to
    /// the Cancel button.
    /// </para>
    /// <para>
    /// Call <see cref="Window.ShowDialog"/> from the Revit API thread.  When the
    /// dialog closes, inspect <see cref="WasCancelled"/> and <see cref="Results"/>.
    /// </para>
    /// </summary>
    public partial class ProgressDialog : Window
    {
        // ── Fields ────────────────────────────────────────────────────────────

        private readonly Document          _doc;
        private readonly List<ViewSheet>   _sheets;
        private readonly SpellCheckEngine  _engine;
        private readonly CancellationTokenSource _cts = new CancellationTokenSource();

        // ── Public results ────────────────────────────────────────────────────

        /// <summary>
        /// Gets the list of spell-check results collected by the background scan.
        /// <c>null</c> if the scan was cancelled or threw an unhandled exception.
        /// </summary>
        public List<SpellCheckResult> Results { get; private set; }

        /// <summary>Gets whether the user clicked Cancel before the scan completed.</summary>
        public bool WasCancelled { get; private set; }

        // ── Constructor ───────────────────────────────────────────────────────

        /// <summary>
        /// Initializes a new instance of <see cref="ProgressDialog"/>.
        /// </summary>
        /// <param name="doc">The active Revit document to scan.</param>
        /// <param name="sheets">
        /// The ordered list of sheets to scan; used to drive the progress bar.
        /// </param>
        /// <param name="engine">
        /// A pre-initialized <see cref="SpellCheckEngine"/> to reuse for all checks.
        /// </param>
        public ProgressDialog(Document doc, List<ViewSheet> sheets, SpellCheckEngine engine)
        {
            InitializeComponent();
            _doc    = doc;
            _sheets = sheets;
            _engine = engine;

            ScanProgressBar.Maximum = sheets.Count;
            Loaded += OnLoaded;
        }

        // ── Loaded: kick off background scan ─────────────────────────────────

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            // IProgress<T> marshals reports back to the UI dispatcher automatically.
            var progress = new Progress<ScanProgressReport>(OnProgressReport);
            var ct       = _cts.Token;

            Task.Run(() => RunScan(progress, ct), ct)
                .ContinueWith(
                    t => Dispatcher.Invoke(() => OnScanCompleted(t)),
                    TaskScheduler.Default);
        }

        // ── Progress callback (called on UI thread via Progress<T>) ───────────

        private void OnProgressReport(ScanProgressReport report)
        {
            ScanProgressBar.Value = report.CurrentSheet;
            StatusText.Text =
                $"Sheet {report.CurrentSheet} of {report.TotalSheets}:  " +
                $"{report.SheetNumber} — {report.SheetName}";
        }

        // ── Background scan ───────────────────────────────────────────────────

        /// <summary>
        /// Runs on a background thread.  Collects text items from each sheet using
        /// the Revit API (read-only access) and spell-checks each item.
        /// </summary>
        private List<SpellCheckResult> RunScan(
            IProgress<ScanProgressReport> progress,
            CancellationToken ct)
        {
            var results              = new List<SpellCheckResult>();
            var seenLegendElemIds    = new HashSet<ElementId>();

            for (int i = 0; i < _sheets.Count; i++)
            {
                ct.ThrowIfCancellationRequested();

                var sheet = _sheets[i];
                progress.Report(new ScanProgressReport
                {
                    CurrentSheet = i + 1,
                    TotalSheets  = _sheets.Count,
                    SheetNumber  = sheet.SheetNumber,
                    SheetName    = sheet.Name
                });

                // Collect text items from this sheet.
                var items = new List<SheetTextItem>();
                SheetTextCollector.CollectFromSheet(_doc, sheet, seenLegendElemIds, items);

                // Spell-check each item (no Revit API calls inside).
                foreach (var item in items)
                {
                    ct.ThrowIfCancellationRequested();

                    var findings = _engine.CheckText(item.Text);
                    foreach (var (word, suggestions) in findings)
                    {
                        results.Add(new SpellCheckResult
                        {
                            MisspelledWord = word,
                            OriginalText   = item.Text,
                            ElementId      = item.ElementId,
                            SheetNumber    = item.SheetNumber,
                            SheetName      = item.SheetName,
                            ElementType    = item.ElementType,
                            ParameterName  = item.ParameterName,
                            IsEditable     = item.IsEditable,
                            Suggestions    = suggestions
                        });
                    }
                }
            }

            return results;
        }

        // ── Scan completed (back on UI thread) ────────────────────────────────

        private void OnScanCompleted(Task<List<SpellCheckResult>> task)
        {
            if (task.IsCanceled || _cts.IsCancellationRequested)
            {
                WasCancelled = true;
            }
            else if (task.IsFaulted)
            {
                WasCancelled = false;
                Results      = new List<SpellCheckResult>();
                // Surface the exception message without crashing.
                StatusText.Text = "Error during scan: " + task.Exception?.GetBaseException().Message;
            }
            else
            {
                WasCancelled = false;
                Results      = task.Result;
            }

            Close();
        }

        // ── Cancel button ─────────────────────────────────────────────────────

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            _cts.Cancel();
            CancelButton.IsEnabled = false;
            StatusText.Text        = "Cancelling…";
        }

        // ── Nested progress report type ───────────────────────────────────────

        private sealed class ScanProgressReport
        {
            public int    CurrentSheet { get; set; }
            public int    TotalSheets  { get; set; }
            public string SheetNumber  { get; set; }
            public string SheetName    { get; set; }
        }
    }
}
