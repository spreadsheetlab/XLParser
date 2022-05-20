using System.Linq;
using Irony.Parsing;

namespace XLParser.Web.XLParserVersions.v152
{
    public enum ReferenceType
    {
        Cell,
        CellRange,
        UserDefinedName,
        HorizontalRange,
        VerticalRange,
        RefError,
        Table
    }

    public class ParserReference
    {
        public const int MaxRangeHeight = 1048576;
        public const int MaxRangeWidth = 16384;

        public ReferenceType ReferenceType { get; set; }
        public string LocationString { get; set; }
        public string Worksheet { get; set; }
        public string LastWorksheet { get; set; }
        public string FilePath { get; set; }
        public string FileName { get; set; }
        public string Name { get; private set; }
        public string MinLocation { get; set; } //Location as appearing in the formula, eg $A$1
        public string MaxLocation { get; set; }

        public ParserReference(ReferenceType referenceType, string locationString = null, string worksheet = null, string lastWorksheet = null,
            string filePath = null, string fileName = null, string name = null, string minLocation = null, string maxLocation = null)
        {
            ReferenceType = referenceType;
            LocationString = locationString;
            Worksheet = worksheet;
            LastWorksheet = lastWorksheet;
            FilePath = filePath;
            FileName = fileName;
            Name = name;
            MinLocation = minLocation;
            MaxLocation = maxLocation != null ? maxLocation : minLocation;
        }

        public ParserReference(ParseTreeNode node)
        {
            InitializeReference(node);
        }

        /// <summary>
        ///     Initializes the current object based on the input ParseTreeNode
        /// </summary>
        /// <remarks>
        ///     For Reference nodes (Prefix ReferenceItem), it initialize the values derived from the Prefix node and
        ///     is re-invoked for the ReferenceItem node.
        /// </remarks>
        public void InitializeReference(ParseTreeNode node)
        {
            switch (node.Type())
            {
                case GrammarNames.Reference:
                    PrefixInfo prefix = node.ChildNodes[0].GetPrefixInfo();
                    Worksheet = prefix.HasSheet ? prefix.Sheet.Replace("''", "'") : "(Undefined sheet)";

                    if (prefix.HasMultipleSheets)
                    {
                        string[] sheets = prefix.MultipleSheets.Split(':');
                        Worksheet = sheets[0];
                        LastWorksheet = sheets[1];
                    }

                    if (prefix.HasFilePath)
                    {
                        FilePath = prefix.FilePath;
                    }

                    if (prefix.HasFileNumber)
                    {
                        FileName = prefix.FileNumber.ToString();
                    }
                    else if (prefix.HasFileName)
                    {
                        FileName = prefix.FileName;
                    }

                    InitializeReference(node.ChildNodes[1]);
                    break;
                case GrammarNames.Cell:
                    ReferenceType = ReferenceType.Cell;
                    MinLocation = node.ChildNodes[0].Token.ValueString;
                    MaxLocation = MinLocation;
                    break;
                case GrammarNames.NamedRange:
                    ReferenceType = ReferenceType.UserDefinedName;
                    Name = node.ChildNodes[0].Token.ValueString;
                    break;
                case GrammarNames.StructuredReference:
                    ReferenceType = ReferenceType.Table;
                    Name = node.ChildNodes.FirstOrDefault(x => x.Type() == GrammarNames.StructuredReferenceTable)?.ChildNodes[0].Token.ValueString;
                    break;
                case GrammarNames.HorizontalRange:
                    string[] horizontalLimits = node.ChildNodes[0].Token.ValueString.Split(':');
                    ReferenceType = ReferenceType.HorizontalRange;
                    MinLocation = horizontalLimits[0];
                    MaxLocation = horizontalLimits[1];
                    break;
                case GrammarNames.VerticalRange:
                    string[] verticalLimits = node.ChildNodes[0].Token.ValueString.Split(':');
                    ReferenceType = ReferenceType.VerticalRange;
                    MinLocation = verticalLimits[0];
                    MaxLocation = verticalLimits[1];
                    break;
                case GrammarNames.RefError:
                    ReferenceType = ReferenceType.RefError;
                    break;
            }

            LocationString = node.Print();
        }

        /// <summary>
        ///     Converts the column number to an Excel column string representation.
        /// </summary>
        /// <param name="columnNumber">The zero-based column number.</param>
        private  string ConvertColumnToStr(int columnNumber)
        {
            var sb = new System.Text.StringBuilder();
            while (columnNumber >= 0)
            {
                sb.Insert(0, (char)(65 + columnNumber % 26));
                columnNumber = columnNumber / 26 - 1;
            }
            return sb.ToString();
        }

        public override string ToString()
        {
            return ReferenceType == ReferenceType.Cell ? MinLocation.ToString() : string.Format("{0}:{1}", MinLocation, MaxLocation);
        }
    }
}