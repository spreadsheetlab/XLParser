using System;
using System.Reflection;
using System.Text;
using Irony.Parsing;

namespace XLParser.Web.XLParserVersions.v170
{
    // TODO: This class is a good example of why an AST is a good idea, for the prefixes the parse trees are too complicated to work with. See #23

    /// <summary>
    /// Simple data class that holds information about a Prefix.
    /// </summary>
    /// <seealso cref="ExcelFormulaParser.GetPrefixInfo"/>
    public class PrefixInfo : IEquatable<PrefixInfo>
    {
        public string FilePath { get; }
        public bool HasFilePath => FilePath != null;

        private readonly int? _fileNumber;
        public int FileNumber => _fileNumber.GetValueOrDefault();
        public bool HasFileNumber => _fileNumber.HasValue;

        public string FileName { get; }
        public bool HasFileName => FileName != null;

        public bool HasFile => HasFileName || HasFileNumber;

        public string Sheet { get; }
        public bool HasSheet => Sheet != null;

        public string MultipleSheets { get; }
        public bool HasMultipleSheets => MultipleSheets != null;

        public bool IsQuoted { get; }

        public PrefixInfo(string sheet = null, int? fileNumber = null, string fileName = null, string filePath = null, string multipleSheets = null, bool isQuoted = false)
        {
            Sheet = sheet;
            _fileNumber = fileNumber;
            FileName = fileName;
            FilePath = filePath;
            MultipleSheets = multipleSheets;
            IsQuoted = isQuoted;
        }

        /// <summary>
        /// Create a PrefixInfo class from a parse tree node
        /// </summary>
        internal static PrefixInfo From(ParseTreeNode prefix)
        {
            if (prefix.Type() != GrammarNames.Prefix)
            {
                throw new ArgumentException("Not a prefix", nameof(prefix));
            }

            string filePath = null;
            int? fileNumber = null;
            string fileName = null;
            string sheetName = null;
            string multipleSheets = null;

            // Token number we're processing
            var cur = 0;

            // Check for quotes
            var isQuoted = prefix.ChildNodes[cur].Is("'");
            if (isQuoted)
            {
                cur++;
            }

            // Check and process file
            if (prefix.ChildNodes[cur].Is(GrammarNames.File))
            {
                ParseTreeNode file = prefix.ChildNodes[cur];

                if (file.ChildNodes[0].Is(GrammarNames.TokenFileNameNumeric))
                {
                    // Numeric filename
                    fileNumber = int.TryParse(Substr(file.ChildNodes[0].Print(), 1, 1), out var n) ? n : default(int?);
                }
                else
                {
                    // String filename
                    var iCur = 0;
                    // Check if it includes a path
                    if (file.ChildNodes[iCur].Is(GrammarNames.TokenFilePath))
                    {
                        filePath = file.ChildNodes[iCur].Print();
                        iCur++;
                    }

                    if (file.ChildNodes[iCur].Is(GrammarNames.TokenFileNameEnclosedInBrackets))
                    {
                        fileName = Substr(file.ChildNodes[iCur].Print(), 1, 1);
                    }
                    else
                    {
                        fileName = file.ChildNodes[iCur].Print();
                    }
                }

                cur++;
            }

            // Check for a non-quoted sheet
            if (prefix.ChildNodes[cur].Is(GrammarNames.TokenSheet))
            {
                sheetName = Substr(prefix.ChildNodes[cur].Print(), 1);
            }
            // Check for a quoted sheet
            else if (prefix.ChildNodes[cur].Is(GrammarNames.TokenSheetQuoted))
            {
                // remove quote and !
                sheetName = Substr(prefix.ChildNodes[cur].Print(), 2);

                if (sheetName == "")
                {
                    // The sheet name consists solely of whitespace (see https://github.com/spreadsheetlab/XLParser/issues/37)
                    // We can not identify the sheet name in the case, and return all whitespace-only sheet names as if they were a single-space sheet name.
                    sheetName = " ";
                }
            }
            // Check if multiple sheets
            else if (prefix.ChildNodes[cur].Is(GrammarNames.TokenMultipleSheets))
            {
                multipleSheets = Substr(prefix.ChildNodes[cur].Print(), 1);
            }

            return new PrefixInfo(sheetName, fileNumber, fileName, filePath, multipleSheets, isQuoted);
        }

        internal static void FixQuotedSheetNodeForWhitespace(ParseTreeNode quotedSheetNode, string sourceText)
        {
            var newPosition = GetSheetNamePositionFromSourceText(quotedSheetNode, sourceText);
            SourceLocation currentLocation = quotedSheetNode.Span.Location;
            if (newPosition == currentLocation.Position)
            {
                return;
            }

            var newLocation = new SourceLocation(newPosition, currentLocation.Line, currentLocation.Column + currentLocation.Position - newPosition);
            quotedSheetNode.Span = new SourceSpan(newLocation, quotedSheetNode.Span.EndPosition - newPosition);

            // Cannot directly assign to quotedSheetNode.Token.Text; it is read-only. Falling back on reflection.
            typeof(Token).GetField("Text", BindingFlags.Instance | BindingFlags.Public)
                ?.SetValue(quotedSheetNode.Token, sourceText.Substring(newPosition, quotedSheetNode.Span.Length));
        }

        private static int GetSheetNamePositionFromSourceText(ParseTreeNode nodeSheetQuoted, string sourceText)
        {
            var startIndex = nodeSheetQuoted.Span.Location.Position;

            while (startIndex > 0)
            {
                if (!char.IsWhiteSpace(sourceText[startIndex - 1]))
                {
                    break;
                }

                startIndex--;
            }

            return startIndex;
        }

        private static string Substr(string s, int removeLast = 0, int removeFirst = 0)
        {
            return s.Substring(removeFirst, s.Length - removeLast - removeFirst);
        }

        public override bool Equals(object other) => Equals(other as PrefixInfo);
        public bool Equals(PrefixInfo other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return _fileNumber == other._fileNumber && string.Equals(FilePath, other.FilePath, StringComparison.OrdinalIgnoreCase) && string.Equals(FileName, other.FileName, StringComparison.OrdinalIgnoreCase) && string.Equals(Sheet, other.Sheet, StringComparison.OrdinalIgnoreCase) && string.Equals(MultipleSheets, other.MultipleSheets, StringComparison.OrdinalIgnoreCase);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = StringComparer.OrdinalIgnoreCase.GetHashCode(Sheet ?? "");
                hashCode = (hashCode*397) ^ (FilePath != null ? StringComparer.OrdinalIgnoreCase.GetHashCode(FilePath) : 0);
                hashCode = (hashCode*397) ^ (FileName != null ? StringComparer.OrdinalIgnoreCase.GetHashCode(FileName) : 0);
                hashCode = (hashCode*397) ^ (_fileNumber?.GetHashCode() ?? 0);
                hashCode = (hashCode*397) ^ (MultipleSheets != null ? StringComparer.OrdinalIgnoreCase.GetHashCode(MultipleSheets) : 0);
                return hashCode;
            }
        }

        public static bool operator ==(PrefixInfo left, PrefixInfo right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(PrefixInfo left, PrefixInfo right)
        {
            return !Equals(left, right);
        }

        public override string ToString()
        {
            var res = new StringBuilder();
            if (IsQuoted) res.Append("'");
            if (HasFilePath) res.Append(FilePath);
            if (HasFileNumber) res.Append($"[{FileNumber}]");
            if (HasFileName) res.Append($"[{FileName}]");
            if (HasSheet) res.Append(Sheet);
            if (HasMultipleSheets) res.Append(MultipleSheets);
            if (IsQuoted) res.Append("'");
            res.Append("!");
            return res.ToString();
        }

    }
}
