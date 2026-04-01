using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Data;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace H2M
{
    // ── View-model for a single DataGrid row ─────────────────────────────────

    /// <summary>
    /// View-model that wraps a <see cref="SpellCheckResult"/> for display in the
    /// results <see cref="System.Windows.Controls.DataGrid"/>.
    /// Implements <see cref="INotifyPropertyChanged"/> so the Suggestions ComboBox
    /// binding updates live as the user changes their selection.
    /// </summary>
    public class ResultRowViewModel : INotifyPropertyChanged
    {
        private string _selectedSuggestion;

        /// <summary>Gets the underlying spell-check result.</summary>
        public SpellCheckResult Result { get; }

        /// <summary>Gets the misspelled word.</summary>
        public string MisspelledWord => Result.MisspelledWord;

        /// <summary>Gets the sheet number.</summary>
        public string SheetNumber => Result.SheetNumber;

        /// <summary>Gets the sheet name.</summary>
        public string SheetName => Result.SheetName;

        /// <summary>Gets the human-readable element type.</summary>
        public string ElementType => Result.ElementType;

        /// <summary>Gets the list of spelling suggestions for the ComboBox.</summary>
        public List<string> Suggestions => Result.Suggestions;

        /// <summary>Gets whether the element can be edited.</summary>
        public bool IsEditable => Result.IsEditable;

        /// <summary>
        /// Gets the tooltip text for the Fix button.
        /// Read-only elements show an explanatory message; editable elements show
        /// a generic action hint.
        /// </summary>
        public string FixButtonTooltip =>
            Result.IsEditable
                ? "Apply the selected suggestion to this element"
                : "Cannot edit — element is read-only";

        /// <summary>
        /// Gets or sets the currently selected suggestion in the Suggestions ComboBox.
        /// Defaults to the first (highest-ranked) suggestion.
        /// </summary>
        public string SelectedSuggestion
        {
            get => _selectedSuggestion;
            set
            {
                if (_selectedSuggestion == value) return;
                _selectedSuggestion = value;
                OnPropertyChanged(nameof(SelectedSuggestion));
            }
        }

        /// <inheritdoc/>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>Initializes a new <see cref="ResultRowViewModel"/>.</summary>
        public ResultRowViewModel(SpellCheckResult result)
        {
            Result             = result;
            _selectedSuggestion = result.Suggestions.FirstOrDefault();
        }

        private void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    // ── Dialog code-behind ────────────────────────────────────────────────────

    /// <summary>
    /// WPF dialog that presents spell-check results in a DataGrid and allows the
    /// user to fix, skip, navigate, or add words to the firm dictionary.
    /// </summary>
    public partial class ResultsDialog : Window
    {
        private readonly Document                    _doc;
        private readonly UIDocument                  _uidoc;
        private readonly SpellCheckEngine            _engine;
        private readonly SpellCheckCommand.SpellCheckStats _stats;
        private readonly ObservableCollection<ResultRowViewModel> _rows;

        /// <summary>Initializes the results dialog.</summary>
        /// <param name="doc">Active Revit document — required for Transactions.</param>
        /// <param name="uidoc">Active UIDocument — required for ShowElements navigation.</param>
        /// <param name="results">Results produced by the background scan.</param>
        /// <param name="engine">Shared engine instance for adding abbreviations.</param>
        /// <param name="stats">Session statistics updated as the user acts on results.</param>
        public ResultsDialog(
            Document                          doc,
            UIDocument                        uidoc,
            List<SpellCheckResult>            results,
            SpellCheckEngine                  engine,
            SpellCheckCommand.SpellCheckStats stats)
        {
            InitializeComponent();
            _doc    = doc;
            _uidoc  = uidoc;
            _engine = engine;
            _stats  = stats;

            _rows = new ObservableCollection<ResultRowViewModel>(
                results.Select(r => new ResultRowViewModel(r)));

            ResultsGrid.ItemsSource = _rows;
            UpdateIssueCount();
        }

        // ── Row-level actions ─────────────────────────────────────────────────

        /// <summary>Fixes the spelling of the row whose Fix button was clicked.</summary>
        private void FixRow_Click(object sender, RoutedEventArgs e)
        {
            var row = RowFromButtonTag(sender);
            if (row == null || !row.IsEditable) return;

            if (string.IsNullOrEmpty(row.SelectedSuggestion)) return;

            if (ApplyFix(row.Result, row.SelectedSuggestion))
            {
                _stats.IssuesFixed++;
                _rows.Remove(row);
                UpdateIssueCount();
            }
        }

        /// <summary>Skips (dismisses) the row whose Skip button was clicked.</summary>
        private void SkipRow_Click(object sender, RoutedEventArgs e)
        {
            var row = RowFromButtonTag(sender);
            if (row == null) return;
            _stats.IssuesSkipped++;
            _rows.Remove(row);
            UpdateIssueCount();
        }

        /// <summary>Navigates Revit to the element in the row whose Go To button was clicked.</summary>
        private void Navigate_Click(object sender, RoutedEventArgs e)
        {
            var row = RowFromButtonTag(sender);
            if (row == null) return;

            try
            {
                _uidoc.ShowElements(new List<ElementId> { row.Result.ElementId });

                // Briefly assert Topmost so the dialog snaps back above Revit.
                bool was = Topmost;
                Topmost = true;
                Topmost = was;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not navigate to element:\n{ex.Message}",
                    "Navigate", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // ── Toolbar actions ───────────────────────────────────────────────────

        /// <summary>
        /// Fixes every row that shares the same misspelled word as the selected row.
        /// Read-only rows that cannot be fixed are skipped instead.
        /// </summary>
        private void FixAll_Click(object sender, RoutedEventArgs e)
        {
            var selected = ResultsGrid.SelectedItem as ResultRowViewModel;
            if (selected == null)
            {
                MessageBox.Show("Select a row first to identify the word.",
                    "Fix All", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string word       = selected.MisspelledWord;
            string suggestion = selected.SelectedSuggestion;
            if (string.IsNullOrEmpty(suggestion)) return;

            // Apply fix to all editable rows with this word.
            var editableRows = _rows.Where(r =>
                r.MisspelledWord == word && r.IsEditable).ToList();

            foreach (var row in editableRows)
            {
                if (ApplyFix(row.Result, suggestion))
                    _stats.IssuesFixed++;
            }

            // Remove all rows for this word (fixed and read-only alike).
            var allRows = _rows.Where(r => r.MisspelledWord == word).ToList();
            foreach (var row in allRows)
            {
                if (!editableRows.Contains(row))
                    _stats.IssuesSkipped++;
                _rows.Remove(row);
            }

            UpdateIssueCount();
        }

        /// <summary>
        /// Removes every row that shares the same misspelled word as the selected row.
        /// </summary>
        private void SkipAll_Click(object sender, RoutedEventArgs e)
        {
            var selected = ResultsGrid.SelectedItem as ResultRowViewModel;
            if (selected == null)
            {
                MessageBox.Show("Select a row first to identify the word.",
                    "Skip All", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string word    = selected.MisspelledWord;
            var toRemove   = _rows.Where(r => r.MisspelledWord == word).ToList();
            foreach (var row in toRemove)
            {
                _stats.IssuesSkipped++;
                _rows.Remove(row);
            }
            UpdateIssueCount();
        }

        /// <summary>
        /// Adds the selected word to <c>ApprovedAbbreviations.json</c> and removes
        /// every matching row from the grid.
        /// </summary>
        private void AddToDictionary_Click(object sender, RoutedEventArgs e)
        {
            var selected = ResultsGrid.SelectedItem as ResultRowViewModel;
            if (selected == null)
            {
                MessageBox.Show("Select a row first.",
                    "Add to Dictionary", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string word = selected.MisspelledWord.ToUpperInvariant();
            _engine.AddApprovedAbbreviation(word);
            _stats.WordsAddedToDictionary++;

            var toRemove = _rows
                .Where(r => r.MisspelledWord.ToUpperInvariant() == word).ToList();
            foreach (var row in toRemove)
                _rows.Remove(row);

            UpdateIssueCount();

            MessageBox.Show($"\"{word}\" has been added to the approved abbreviations list.",
                "Added to Dictionary", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        /// <summary>Groups the DataGrid rows by sheet number.</summary>
        private void GroupBySheet_Checked(object sender, RoutedEventArgs e)
        {
            var view = CollectionViewSource.GetDefaultView(ResultsGrid.ItemsSource);
            if (view == null) return;
            view.GroupDescriptions.Clear();
            view.GroupDescriptions.Add(
                new PropertyGroupDescription(nameof(ResultRowViewModel.SheetNumber)));
        }

        /// <summary>Removes sheet grouping from the DataGrid.</summary>
        private void GroupBySheet_Unchecked(object sender, RoutedEventArgs e)
        {
            CollectionViewSource.GetDefaultView(ResultsGrid.ItemsSource)
                ?.GroupDescriptions.Clear();
        }

        /// <summary>Closes the dialog; any remaining rows are counted as skipped.</summary>
        private void Done_Click(object sender, RoutedEventArgs e)
        {
            _stats.IssuesSkipped += _rows.Count;
            Close();
        }

        // ── Fix implementation ────────────────────────────────────────────────

        /// <summary>
        /// Applies a spelling fix to a Revit element, wrapped in a named Transaction.
        /// </summary>
        /// <param name="result">The result describing the element and misspelled word.</param>
        /// <param name="replacement">The corrected word to substitute.</param>
        /// <returns><c>true</c> if the element was updated successfully.</returns>
        private bool ApplyFix(SpellCheckResult result, string replacement)
        {
            try
            {
                var element = _doc.GetElement(result.ElementId);
                if (element == null) return false;

                using (var tx = new Transaction(_doc, "Fix Spelling"))
                {
                    tx.Start();

                    if (result.ParameterName == null)
                    {
                        // TextNote body text.
                        var textNote = element as TextNote;
                        if (textNote == null) { tx.RollBack(); return false; }

                        textNote.Text = ReplaceWholeWord(
                            textNote.Text, result.MisspelledWord, replacement);
                    }
                    else
                    {
                        // Parameter-based text.
                        Parameter param = GetParameterByName(element, result.ParameterName);
                        if (param == null || param.IsReadOnly)
                        { tx.RollBack(); return false; }

                        string current = param.AsString() ?? string.Empty;
                        param.Set(ReplaceWholeWord(current, result.MisspelledWord, replacement));
                    }

                    tx.Commit();
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Replaces the first whole-word occurrence of <paramref name="word"/> in
        /// <paramref name="text"/> with <paramref name="replacement"/>.
        /// Uses a regex word-boundary so "cast" does not accidentally match "precast".
        /// </summary>
        private static string ReplaceWholeWord(string text, string word, string replacement)
        {
            string pattern = @"\b" + Regex.Escape(word) + @"\b";
            return Regex.Replace(text, pattern, replacement, RegexOptions.None);
        }

        /// <summary>
        /// Finds a parameter on <paramref name="element"/> by name
        /// (case-insensitive match).
        /// </summary>
        private static Parameter GetParameterByName(Element element, string paramName)
        {
            foreach (Parameter p in element.Parameters)
            {
                if (string.Equals(p.Definition.Name, paramName,
                        StringComparison.OrdinalIgnoreCase))
                    return p;
            }
            return null;
        }

        // ── UI helpers ────────────────────────────────────────────────────────

        private static ResultRowViewModel RowFromButtonTag(object sender)
        {
            return (sender as System.Windows.Controls.Button)?.Tag
                   as ResultRowViewModel;
        }

        private void UpdateIssueCount()
        {
            int n = _rows.Count;
            IssueCountText.Text = n == 0
                ? "All issues resolved."
                : $"{n} issue{(n == 1 ? "" : "s")} remaining";
        }
    }
}
