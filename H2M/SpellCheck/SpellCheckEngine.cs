using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using WeCantSpell.Hunspell;

namespace H2M
{
    /// <summary>
    /// Provides spell checking using the WeCantSpell.Hunspell library with
    /// construction-drawing–aware text preprocessing.
    /// <para>
    /// The Hunspell dictionary (<c>en_US.dic</c> / <c>en_US.aff</c>) and the
    /// approved-abbreviations list are loaded once at construction and reused
    /// for all subsequent checks.  Dispose the engine when the session ends.
    /// </para>
    /// </summary>
    public sealed class SpellCheckEngine : IDisposable
    {
        private WordList _wordList;
        private HashSet<string> _approvedAbbreviations;
        private readonly string _abbrevFilePath;
        private bool _disposed;

        // Matches common construction-drawing dimension tokens so they are stripped
        // before word-level checking.  Examples:
        //   6'-8"   3'-0"   12"   1'-6 1/2"   W14x48   #4@12"   (E)   (N)
        private static readonly Regex _dimPattern = new Regex(
            @"\d+['′\u2019][-\s]?\d*(?:[/\d]+)?[""″]?"  // foot-inch: 6'-8", 3'-0"
          + @"|\d+[""″]"                                  // bare inches: 12"
          + @"|W\d+[Xx]\d+"                               // steel section: W14x48
          + @"|#\d+@\d+[""″]?"                            // rebar: #4@12"
          + @"|\([ENen]\)",                               // existing/new: (E) (N)
            RegexOptions.Compiled);

        // Matches a token that is entirely numeric (possibly with decimals/fractions).
        private static readonly Regex _numberPattern = new Regex(
            @"^[\d/.,]+$", RegexOptions.Compiled);

        /// <summary>
        /// Initializes the spell check engine.  Loads the Hunspell dictionary from
        /// the same directory as the executing assembly and reads
        /// <c>ApprovedAbbreviations.json</c> from that directory.
        /// </summary>
        /// <exception cref="FileNotFoundException">
        /// Thrown if <c>en_US.dic</c> or <c>en_US.aff</c> is missing from the
        /// output directory.
        /// </exception>
        public SpellCheckEngine()
        {
            string dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                         ?? AppDomain.CurrentDomain.BaseDirectory;

            string affPath = Path.Combine(dir, "en_US.aff");
            string dicPath = Path.Combine(dir, "en_US.dic");

            _wordList = WordList.CreateFromFiles(dicPath, affPath);

            _abbrevFilePath = Path.Combine(dir, "ApprovedAbbreviations.json");
            LoadApprovedAbbreviations();
        }

        // ── Public API ────────────────────────────────────────────────────────

        /// <summary>
        /// Checks <paramref name="text"/> for misspelled words after applying
        /// construction-drawing preprocessing rules.
        /// </summary>
        /// <param name="text">The text string to check.  Empty or whitespace strings are
        /// silently skipped and return an empty list.</param>
        /// <returns>
        /// A list of <c>(word, suggestions)</c> tuples — one entry per distinct
        /// misspelled word occurrence.  Suggestions contain up to five candidates
        /// ordered by likelihood; the list may be empty if Hunspell has no suggestions.
        /// </returns>
        public List<(string Word, List<string> Suggestions)> CheckText(string text)
        {
            var results = new List<(string, List<string>)>();
            if (string.IsNullOrWhiteSpace(text)) return results;

            // Strip dimension tokens.
            string cleaned = _dimPattern.Replace(text, " ");

            // Tokenize on whitespace and common punctuation characters.
            // We keep dots inside tokens so "B.O.C." can be handled as an abbreviation.
            string[] tokens = Regex.Split(cleaned, @"[\s,;:()\[\]{}<>!?""]+");

            foreach (string token in tokens)
            {
                if (string.IsNullOrWhiteSpace(token)) continue;

                // Strip dot-notation abbreviations: "B.O.C." → "BOC"
                string noDots = token.Replace(".", string.Empty);
                if (noDots.Length < 2) continue;

                // Split on hyphens so "pre-cast" checks "pre" and "cast" individually.
                string[] parts = noDots.Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (string part in parts)
                {
                    var finding = CheckWord(part);
                    if (finding.HasValue)
                        results.Add(finding.Value);
                }
            }

            return results;
        }

        /// <summary>
        /// Appends <paramref name="word"/> (normalised to uppercase) to the approved
        /// abbreviations list both in memory and in <c>ApprovedAbbreviations.json</c>
        /// on disk.  No-ops silently if the word is already approved.
        /// </summary>
        /// <param name="word">The word to approve.  Stored as uppercase.</param>
        public void AddApprovedAbbreviation(string word)
        {
            if (string.IsNullOrWhiteSpace(word)) return;
            string upper = word.ToUpperInvariant();
            if (!_approvedAbbreviations.Add(upper)) return;   // already present

            PersistAbbreviations();
        }

        // ── Private helpers ───────────────────────────────────────────────────

        /// <summary>
        /// Checks a single word token against the approved list and the Hunspell
        /// dictionary.  Returns <c>null</c> when the word is acceptable.
        /// </summary>
        private (string Word, List<string> Suggestions)? CheckWord(string word)
        {
            if (string.IsNullOrWhiteSpace(word)) return null;

            // Strip possessives and stray apostrophes/dashes from edges.
            string w = word.Trim('-', '\'', '\u2018', '\u2019');
            // Strip trailing possessive 's  ("building's" → "building")
            if (w.EndsWith("'s", StringComparison.OrdinalIgnoreCase)
                || w.EndsWith("\u2019s", StringComparison.OrdinalIgnoreCase))
                w = w.Substring(0, w.Length - 2);

            w = w.Trim('-', '\'', '\u2018', '\u2019');

            if (w.Length < 2) return null;

            // Skip pure numeric tokens (standalone numbers, fractions like "1/2").
            if (_numberPattern.IsMatch(w)) return null;

            // ── Approved abbreviation lookup (normalised to UPPERCASE) ────────
            // This covers "Typ.", "TYP", "typ", "T.Y.P." — all normalise to "TYP".
            string upper = w.ToUpperInvariant();
            if (_approvedAbbreviations.Contains(upper)) return null;

            // ── Hunspell dictionary check ─────────────────────────────────────
            if (_wordList.Check(w)) return null;

            // For all-caps words that failed the abbreviation check, also try
            // lowercase so "CONCRETE" passes even if not in the abbreviation list.
            if (w == upper && _wordList.Check(w.ToLowerInvariant())) return null;

            // Misspelled — collect suggestions (cap at 5).
            var suggestions = _wordList.Suggest(w).Take(5).ToList();
            return (w, suggestions);
        }

        /// <summary>Loads or re-loads the approved abbreviations from disk into memory.</summary>
        private void LoadApprovedAbbreviations()
        {
            _approvedAbbreviations = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (!File.Exists(_abbrevFilePath)) return;

            try
            {
                string json = File.ReadAllText(_abbrevFilePath, Encoding.UTF8);
                using (JsonDocument doc = JsonDocument.Parse(json))
                {
                    if (doc.RootElement.TryGetProperty("ApprovedAbbreviations", out JsonElement arr))
                    {
                        foreach (JsonElement el in arr.EnumerateArray())
                        {
                            string val = el.GetString();
                            if (!string.IsNullOrWhiteSpace(val))
                                _approvedAbbreviations.Add(val.ToUpperInvariant());
                        }
                    }
                }
            }
            catch
            {
                // Silently tolerate a missing or malformed JSON file.
                // The engine will run without approved abbreviations until the file is corrected.
            }
        }

        /// <summary>Serialises the current in-memory abbreviation set back to disk.</summary>
        private void PersistAbbreviations()
        {
            try
            {
                var sorted  = _approvedAbbreviations.OrderBy(x => x).ToList();
                var payload = new { ApprovedAbbreviations = sorted };
                string json = JsonSerializer.Serialize(payload,
                    new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_abbrevFilePath, json, Encoding.UTF8);
            }
            catch
            {
                // Silently swallow file-write errors (e.g. read-only network share).
            }
        }

        // ── IDisposable ───────────────────────────────────────────────────────

        /// <summary>
        /// Releases the Hunspell word list and any native resources it holds.
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;
            _wordList?.Dispose();
            _wordList  = null;
            _disposed  = true;
        }
    }
}
