using System.Windows;

namespace H2M
{
    /// <summary>
    /// Simple WPF summary dialog displayed at the end of a spell-check session.
    /// Shows the final counts of sheets scanned, issues found, fixed, skipped,
    /// and words added to the firm dictionary.
    /// </summary>
    public partial class SummaryDialog : Window
    {
        /// <summary>
        /// Initializes the summary dialog and populates all statistic labels.
        /// </summary>
        /// <param name="stats">
        /// The session statistics accumulated during the spell-check run.
        /// </param>
        public SummaryDialog(SpellCheckCommand.SpellCheckStats stats)
        {
            InitializeComponent();

            SheetsScannedText.Text = stats.SheetsScanned.ToString();
            IssuesFoundText.Text   = stats.IssuesFound.ToString();
            IssuesFixedText.Text   = stats.IssuesFixed.ToString();
            IssuesSkippedText.Text = stats.IssuesSkipped.ToString();
            WordsAddedText.Text    = stats.WordsAddedToDictionary.ToString();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();
    }
}
