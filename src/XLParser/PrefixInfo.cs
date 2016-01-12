using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Irony.Parsing;

namespace XLParser
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

        private readonly int? fileNumber;
        public int FileNumber => fileNumber.Value;
        public bool HasFileNumber => fileNumber.HasValue;

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
            this.fileNumber = fileNumber;
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
            if (prefix.Type() != GrammarNames.Prefix) throw new ArgumentException("Not a prefix", nameof(prefix));

            string filePath = null;
            int? fileNumber = null;
            string fileName = null;
            string sheetName = null;
            string multipleSheets = null;

            // Token number we're processing
            int cur = 0;

            // Check for quotes
            bool quoted = prefix.ChildNodes[cur].Is("'");
            if (quoted) cur++;

            // Check and process file
            if (prefix.ChildNodes[cur].Is(GrammarNames.File))
            {
                var file = prefix.ChildNodes[cur];

                if (file.ChildNodes[0].Is(GrammarNames.TokenFileNameNumeric))
                {
                    // Numeric filename
                    int n;
                    int.TryParse(Substr(file.ChildNodes[0].Print(), 1, 1), out n);
                    fileNumber = n;
                    if (fileNumber == 0) fileNumber = null;
                }
                else
                {
                    // String filename
                    var icur = 0;
                    // Check if it includes a path
                    if (file.ChildNodes[icur].Is(GrammarNames.TokenFilePathWindows))
                    {
                        filePath = file.ChildNodes[icur].Print();
                        icur++;
                    }
                    fileName = Substr(file.ChildNodes[icur].Print(), 1, 1);
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
                    // The sheetname consists solely of whitespace (see https://github.com/spreadsheetlab/XLParser/issues/37)
                    // We can not identify the sheetname in the case, and return all whitespace-only sheetnames as if they were a single-space sheetname.
                    sheetName = " ";
                }
            }
            // Check if multiple sheets
            else if (prefix.ChildNodes[cur].Is(GrammarNames.TokenMultipleSheets))
            {
                multipleSheets = Substr(prefix.ChildNodes[cur].Print(), 1);
            }

            // Put it all into the convencience class
            return new PrefixInfo(
                sheetName,
                fileNumber,
                fileName,
                filePath,
                multipleSheets,
                quoted
                );
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
            return fileNumber == other.fileNumber && string.Equals(FilePath, other.FilePath, StringComparison.OrdinalIgnoreCase) && string.Equals(FileName, other.FileName, StringComparison.OrdinalIgnoreCase) && string.Equals(Sheet, other.Sheet, StringComparison.OrdinalIgnoreCase) && string.Equals(MultipleSheets, other.MultipleSheets, StringComparison.OrdinalIgnoreCase);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = StringComparer.OrdinalIgnoreCase.GetHashCode(Sheet ?? "");
                hashCode = (hashCode*397) ^ (FilePath != null ? StringComparer.OrdinalIgnoreCase.GetHashCode(FilePath) : 0);
                hashCode = (hashCode*397) ^ (FileName != null ? StringComparer.OrdinalIgnoreCase.GetHashCode(FileName) : 0);
                hashCode = (hashCode*397) ^ (fileNumber?.GetHashCode() ?? 0);
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
            string res = "";
            if (IsQuoted) res += "'";
            if (HasFilePath) res += FilePath;
            if (HasFileNumber) res += $"[{FileNumber}]";
            if (HasFileName) res += $"[{FileName}]";
            if (HasSheet) res += Sheet;
            if (HasMultipleSheets) res += MultipleSheets;
            if (IsQuoted) res += "'";
            res += "!";
            return res;
        }

    }
}
