using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace H2M
{
    /// <summary>
    /// Represents a single misspelled word instance found within a Revit document element.
    /// One result is produced per misspelled word per element — an element with two
    /// misspelled words produces two <see cref="SpellCheckResult"/> instances.
    /// </summary>
    public class SpellCheckResult
    {
        /// <summary>Gets or sets the misspelled word as extracted from the source text.</summary>
        public string MisspelledWord { get; set; }

        /// <summary>
        /// Gets or sets the complete original text string from which the word was extracted.
        /// Used when performing the Fix operation to replace the word in context.
        /// </summary>
        public string OriginalText { get; set; }

        /// <summary>Gets or sets the Revit element ID of the element that owns this text.</summary>
        public ElementId ElementId { get; set; }

        /// <summary>Gets or sets the sheet number on which the element is placed.</summary>
        public string SheetNumber { get; set; }

        /// <summary>Gets or sets the name of the sheet on which the element is placed.</summary>
        public string SheetName { get; set; }

        /// <summary>
        /// Gets or sets a human-readable description of the source element type,
        /// e.g. "Text Note", "Title Block", "Legend", "Viewport Title", "Sheet Name".
        /// </summary>
        public string ElementType { get; set; }

        /// <summary>
        /// Gets or sets the Revit parameter name when the text originated from a parameter.
        /// <c>null</c> when the source is the body text of a <see cref="TextNote"/>.
        /// </summary>
        public string ParameterName { get; set; }

        /// <summary>
        /// Gets or sets whether the element's text can be edited.
        /// <c>false</c> when the element belongs to a linked model or the parameter is read-only.
        /// </summary>
        public bool IsEditable { get; set; }

        /// <summary>
        /// Gets or sets the spelling suggestions returned by the engine (up to five),
        /// ordered by likelihood. The first entry is pre-selected in the UI.
        /// </summary>
        public List<string> Suggestions { get; set; } = new List<string>();
    }
}
