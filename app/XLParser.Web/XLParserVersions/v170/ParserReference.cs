using System.Linq;
using Irony.Parsing;

namespace XLParser.Web.XLParserVersions.v170
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
        public ReferenceType ReferenceType { get; set; }
        public ParseTreeNode ReferenceNode { get; set; }
        public string LocationString { get; set; }
        public string Worksheet { get; set; }
        public string LastWorksheet { get; set; }
        public string FilePath { get; set; }
        public string FileName { get; set; }
        public string Name { get; set; }
        public string MinLocation { get; set; }
        public string MaxLocation { get; set; }
        public string[] TableSpecifiers { get; set; }
        public string[] TableColumns { get; set; }
        
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
                    Name = node.ChildNodes.FirstOrDefault(x => x.Type() == GrammarNames.StructuredReferenceQualifier)?.ChildNodes[0].Token.ValueString;
                    TableSpecifiers = node.AllNodes().Where(x => x.Is(GrammarNames.TokenSRSpecifier) || x.Is("@")).Select(x => UnEscape(x.Token.ValueString, "'")).ToArray();
                    TableColumns = node.AllNodes().Where(x => x.Is(GrammarNames.TokenSRColumn)).Select(x => UnEscape(x.Token.ValueString, "'")).ToArray();
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

            ReferenceNode = node;
            LocationString = node.Print();
        }

        private string UnEscape(string value, string escapeCharacter)
        {
            return System.Text.RegularExpressions.Regex.Replace(value, $"{escapeCharacter}(?!{escapeCharacter})", "");
        }

        public override string ToString()
        {
            return LocationString;
        }
    }
}