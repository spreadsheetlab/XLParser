using System;
using Irony.Parsing;

namespace XLParser.Web.XLParserVersions.v120
{
    /// <summary>
    /// Contains the XLParser grammar
    /// </summary>
    [Language("Excel Formulas", "1.2.0", "Grammar for Excel Formulas")]
    public class ExcelFormulaGrammar : Grammar
    {
        #region 1-Terminals

        #region Symbols and operators

        public Terminal at => ToTerm("@");
        public Terminal comma => ToTerm(",");
        public Terminal colon => ToTerm(":");
        public Terminal semicolon => ToTerm(";");
        public Terminal OpenParen => ToTerm("(");
        public Terminal CloseParen => ToTerm(")");
        public Terminal CloseSquareParen => ToTerm("]");
        public Terminal OpenSquareParen => ToTerm("[");
        public Terminal exclamationMark => ToTerm("!");
        public Terminal CloseCurlyParen => ToTerm("}");
        public Terminal OpenCurlyParen => ToTerm("{");
        public Terminal QuoteS => ToTerm("'");

        public Terminal mulop => ToTerm("*");
        public Terminal plusop => ToTerm("+");
        public Terminal divop => ToTerm("/");
        public Terminal minop => ToTerm("-");
        public Terminal concatop => ToTerm("&");
        public Terminal expop => ToTerm("^");

        // Intersect op is a single space, which cannot be parsed normally so we need an ImpliedSymbolTerminal
        // Attention: ImpliedSymbolTerminal seems to break if you assign it a priority, and it's default priority is low
        public Terminal intersectop { get; } = new ImpliedSymbolTerminal(GrammarNames.TokenIntersect);

        public Terminal percentop => ToTerm("%");

        public Terminal gtop => ToTerm(">");
        public Terminal eqop => ToTerm("=");
        public Terminal ltop => ToTerm("<");
        public Terminal neqop => ToTerm("<>");
        public Terminal gteop => ToTerm(">=");
        public Terminal lteop => ToTerm("<=");

        #endregion

        #region Literals

        public Terminal BoolToken { get; } = new RegexBasedTerminal(GrammarNames.TokenBool, "TRUE|FALSE")
        {
            Priority = TerminalPriority.Bool
        };

        public Terminal NumberToken { get; } = new NumberLiteral(GrammarNames.TokenNumber, NumberOptions.None)
        {
            DefaultIntTypes = new[] {TypeCode.Int32, TypeCode.Int64, NumberLiteral.TypeCodeBigInt}
        };

        public Terminal TextToken { get; } = new StringLiteral(GrammarNames.TokenText, "\"",
            StringOptions.AllowsDoubledQuote | StringOptions.AllowsLineBreak);

        public Terminal SingleQuotedStringToken { get; } = new StringLiteral(GrammarNames.TokenSingleQuotedString, "'",
            StringOptions.AllowsDoubledQuote | StringOptions.AllowsLineBreak)
        { Priority = TerminalPriority.SingleQuotedString };

        public Terminal ErrorToken { get; } = new RegexBasedTerminal(GrammarNames.TokenError, "#NULL!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#N/A");
        public Terminal RefErrorToken => ToTerm("#REF!", GrammarNames.TokenRefError);

        #endregion

        #region Functions

        public Terminal UDFToken { get; } = new RegexBasedTerminal(GrammarNames.TokenUDF, @"(_xll\.)?[\w\\.]+\(")
        { Priority = TerminalPriority.UDF };

        public Terminal ExcelRefFunctionToken { get; } = new RegexBasedTerminal(GrammarNames.TokenExcelRefFunction, "(INDEX|OFFSET|INDIRECT)\\(")
        { Priority = TerminalPriority.ExcelRefFunction };

        public Terminal ExcelConditionalRefFunctionToken { get; } = new RegexBasedTerminal(GrammarNames.TokenExcelConditionalRefFunction, "(IF|CHOOSE)\\(")
        { Priority = TerminalPriority.ExcelRefFunction };

        public Terminal ExcelFunction { get; } = new RegexBasedTerminal(GrammarNames.ExcelFunction,  "(" + string.Join("|", excelFunctionList) + ")\\(")
        { Priority = TerminalPriority.ExcelFunction };

        // Using this instead of Empty allows a more accurate trees
        public Terminal EmptyArgumentToken { get; } = new ImpliedSymbolTerminal(GrammarNames.TokenEmptyArgument);

        #endregion

        #region References and names

        public Terminal VRangeToken { get; } = new RegexBasedTerminal(GrammarNames.TokenVRange, "[$]?[A-Z]{1,4}:[$]?[A-Z]{1,4}");
        public Terminal HRangeToken { get; } = new RegexBasedTerminal(GrammarNames.TokenHRange, "[$]?[1-9][0-9]*:[$]?[1-9][0-9]*");

        private const string CellTokenRegex = "[$]?[A-Z]{1,4}[$]?[1-9][0-9]*";
        public Terminal CellToken { get; } = new RegexBasedTerminal(GrammarNames.TokenCell, CellTokenRegex)
        { Priority = TerminalPriority.CellToken };

        private const string NamedRangeRegex = @"[A-Za-z\\_][\w\.]*";
        public Terminal NameToken { get; } = new RegexBasedTerminal(GrammarNames.TokenName, NamedRangeRegex)
        { Priority = TerminalPriority.Name };

        // To prevent e.g. "A1A1" being parsed as 2 celltokens
        public Terminal NamedRangeCombinationToken { get; } = new RegexBasedTerminal(GrammarNames.TokenNamedRangeCombination, "(TRUE|FALSE|" + CellTokenRegex + ")" + NamedRangeRegex)
        { Priority = TerminalPriority.NamedRangeCombination };

        public Terminal ReservedNameToken = new RegexBasedTerminal(GrammarNames.TokenReservedName, @"_xlnm\.[a-zA-Z_]+")
        { Priority = TerminalPriority.ReservedName };

        #region Structured References

        //public Terminal SRTableNameToken = new RegexBasedTerminal(GrammarNames.TokenSRTableName, @"[\w\\.]+\[")
        //{Priority = 0};

        public Terminal SRColumnToken = new RegexBasedTerminal(GrammarNames.TokenSRColumn, @"[\w\\.]+")
        { Priority = TerminalPriority.SRColumn };

        //public Terminal SREnclosedColumnToken = new RegexBasedTerminal(GrammarNames.TokenSREnclosedColumn, @"\[( )*[\w+\\.,:#'""{}$^&*+=-></]+( )*\]")
        //{Priority = 0};

        #endregion

        #region Prefixes
        private const string mustBeQuotedInSheetName = @"\(\);{}#""=<>&+\-*/\^%, ";
        private const string notSheetNameChars = @"'*\[\]\\:/?";
        //const string singleQuotedContent = @"\w !@#$%^&*()\-\+={}|:;<>,\./\?" + "\\\"";
        //const string sheetRegEx = @"(([\w\.]+)|('([" + singleQuotedContent + @"]|'')+'))!";
        private static readonly string normalSheetName = $"[^{notSheetNameChars}{mustBeQuotedInSheetName}]+";
        private static readonly string quotedSheetName = $"([^{notSheetNameChars}]|'')+";
        //private static readonly string sheetRegEx = $"(({normalSheetName})|('{quotedSheetName}'))!";

        public Terminal SheetToken = new RegexBasedTerminal(GrammarNames.TokenSheet, $"{normalSheetName}!")
        { Priority = TerminalPriority.SheetToken };

        public Terminal SheetQuotedToken = new RegexBasedTerminal(GrammarNames.TokenSheetQuoted, $"{quotedSheetName}'!")
        { Priority = TerminalPriority.SheetQuotedToken };

        private static readonly string multiSheetRegex = $"(({normalSheetName}:{normalSheetName})|('{quotedSheetName}:{quotedSheetName}'))!";
        public Terminal MultipleSheetsToken = new RegexBasedTerminal(GrammarNames.TokenMultipleSheets, multiSheetRegex)
        { Priority = TerminalPriority.MultipleSheetsToken };

        private const string fileNameNumericRegex = @"\[[0-9]+\]";
        public Terminal FileToken = new RegexBasedTerminal(GrammarNames.TokenFileNameNumeric, fileNameNumericRegex)
        { Priority = TerminalPriority.FileNameNumericToken };

        private const string fileNameRegex = @"\[[^<>:""/\|?*\[\]]+\]";
        public Terminal EnclosedInBracketsToken { get; } = new RegexBasedTerminal(GrammarNames.TokenEnclosedInBrackets, fileNameRegex)
        { Priority = TerminalPriority.FileName };

        // Source: http://stackoverflow.com/a/6416209/572635
        private const string filePathRegex = @"(?:[a-zA-Z]\:|\\\\[\w\.]+\\[\w.$]+)\\(?:[\w]+\\)*";
        public Terminal FilePathWindowsToken { get; } = new RegexBasedTerminal(GrammarNames.TokenFilePathWindows, filePathRegex);
        #endregion

        #endregion

        #endregion

        #region 2-NonTerminals
        // Most nonterminals are first defined here, so they can be used anywhere in the rules
        // Otherwise you can only use nonterminals that have been defined previously

        public NonTerminal Argument{ get; } = new NonTerminal(GrammarNames.Argument);
        public NonTerminal Arguments{ get; } = new NonTerminal(GrammarNames.Arguments);
        public NonTerminal ArrayColumns{ get; } = new NonTerminal(GrammarNames.ArrayColumns);
        public NonTerminal ArrayConstant{ get; } = new NonTerminal(GrammarNames.ArrayConstant);
        public NonTerminal ArrayFormula{ get; } = new NonTerminal(GrammarNames.ArrayFormula);
        public NonTerminal ArrayRows{ get; } = new NonTerminal(GrammarNames.ArrayRows);
        public NonTerminal Bool{ get; } = new NonTerminal(GrammarNames.Bool);
        public NonTerminal Cell{ get; } = new NonTerminal(GrammarNames.Cell);
        public NonTerminal Constant{ get; } = new NonTerminal(GrammarNames.Constant);
        public NonTerminal ConstantArray{ get; } = new NonTerminal(GrammarNames.ConstantArray);
        public NonTerminal DynamicDataExchange{ get; } = new NonTerminal(GrammarNames.DynamicDataExchange);
        public NonTerminal EmptyArgument{ get; } = new NonTerminal(GrammarNames.EmptyArgument);
        public NonTerminal Error{ get; } = new NonTerminal(GrammarNames.Error);
        public NonTerminal File { get; } = new NonTerminal(GrammarNames.File);
        public NonTerminal Formula{ get; } = new NonTerminal(GrammarNames.Formula);
        public NonTerminal FormulaWithEq{ get; } = new NonTerminal(GrammarNames.FormulaWithEq);
        public NonTerminal FunctionCall{ get; } = new NonTerminal(GrammarNames.FunctionCall);
        public NonTerminal FunctionName{ get; } = new NonTerminal(GrammarNames.FunctionName);
        public NonTerminal HRange{ get; } = new NonTerminal(GrammarNames.HorizontalRange);
        public NonTerminal InfixOp{ get; } = new NonTerminal(GrammarNames.TransientInfixOp);
        public NonTerminal NamedRange{ get; } = new NonTerminal(GrammarNames.NamedRange);
        public NonTerminal Number{ get; } = new NonTerminal(GrammarNames.Number);
        public NonTerminal PostfixOp{ get; } = new NonTerminal(GrammarNames.TransientPostfixOp);
        public NonTerminal Prefix{ get; } = new NonTerminal(GrammarNames.Prefix);
        public NonTerminal PrefixOp{ get; } = new NonTerminal(GrammarNames.TransientPrefixOp);
        public NonTerminal QuotedFileSheet{ get; } = new NonTerminal(GrammarNames.QuotedFileSheet);
        public NonTerminal Reference{ get; } = new NonTerminal(GrammarNames.Reference);
        //public NonTerminal ReferenceFunction{ get; } = new NonTerminal(GrammarNames.ReferenceFunction);
        public NonTerminal ReferenceItem{ get; } = new NonTerminal(GrammarNames.TransientReferenceItem);
        public NonTerminal ReferenceFunctionCall{ get; } = new NonTerminal(GrammarNames.ReferenceFunctionCall);
        public NonTerminal RefError{ get; } = new NonTerminal(GrammarNames.RefError);
        public NonTerminal RefFunctionName{ get; } = new NonTerminal(GrammarNames.RefFunctionName);
        public NonTerminal ReservedName{ get; } = new NonTerminal(GrammarNames.ReservedName);
        public NonTerminal Sheet{ get; } = new NonTerminal(GrammarNames.Sheet);
        public NonTerminal Start{ get; } = new NonTerminal(GrammarNames.TransientStart);
        public NonTerminal StructureReference { get; } = new NonTerminal(GrammarNames.StructureReference);
        public NonTerminal StructureReferenceColumnOrKeyword { get; } = new NonTerminal(GrammarNames.StructureReferenceColumnOrKeyword);
        public NonTerminal StructureReferenceExpression { get; } = new NonTerminal(GrammarNames.StructureReferenceExpression);
        //public NonTerminal StructureReferenceKeyword { get; } = new NonTerminal(GrammarNames.StructureReferenceKeyword);
        public NonTerminal StructureReferenceTable { get; } = new NonTerminal(GrammarNames.StructureReferenceTable);
        public NonTerminal Text{ get; } = new NonTerminal(GrammarNames.Text);
        public NonTerminal UDFName{ get; } = new NonTerminal(GrammarNames.UDFName);
        public NonTerminal UDFunctionCall{ get; } = new NonTerminal(GrammarNames.UDFunctionCall);
        public NonTerminal Union{ get; } = new NonTerminal(GrammarNames.Union);
        public NonTerminal VRange{ get; } = new NonTerminal(GrammarNames.VerticalRange);
        #endregion

        public ExcelFormulaGrammar() : base(false)
        {
            
            #region Punctuation
            MarkPunctuation(OpenParen, CloseParen);
            MarkPunctuation(OpenSquareParen, CloseSquareParen);
            MarkPunctuation(OpenCurlyParen, CloseCurlyParen);
            //exclamationMark.SetFlag(TermFlags.IsDelimiter);
            #endregion
            
            #region Rules

            #region Base rules
            Root = Start;

            Start.Rule = FormulaWithEq
                         | Formula
                         | ArrayFormula
                         ;
            MarkTransient(Start);

            ArrayFormula.Rule = OpenCurlyParen + eqop + Formula + CloseCurlyParen;

            FormulaWithEq.Rule = eqop + Formula;

            Formula.Rule =
                Reference
                | Constant
                | FunctionCall
                | ConstantArray
                | OpenParen + Formula + CloseParen
                | ReservedName
                ;
            //MarkTransient(Formula);

            ReservedName.Rule = ReservedNameToken;

            Constant.Rule = Number
                            | Text
                            | Bool
                            | Error
                            ;

            Text.Rule = TextToken;
            Number.Rule = NumberToken;
            Bool.Rule = BoolToken;
            Error.Rule = ErrorToken;
            RefError.Rule = RefErrorToken;
            #endregion

            #region Functions

            FunctionCall.Rule =
                  FunctionName + Arguments + CloseParen
                | PrefixOp + Formula
                | Formula + PostfixOp
                | Formula + InfixOp + Formula
                ;
                
            FunctionName.Rule = ExcelFunction;

            Arguments.Rule = MakeStarRule(Arguments, comma, Argument);
            //Arguments.Rule = Argument | Argument + comma + Arguments;

            EmptyArgument.Rule = EmptyArgumentToken;
            Argument.Rule = Formula | EmptyArgument;
            //MarkTransient(Argument);

            PrefixOp.Rule =
                ImplyPrecedenceHere(Precedence.UnaryPreFix) + plusop
                | ImplyPrecedenceHere(Precedence.UnaryPreFix) + minop;
            MarkTransient(PrefixOp);

            InfixOp.Rule =
                  expop
                | mulop
                | divop
                | plusop
                | minop
                | concatop
                | gtop
                | eqop
                | ltop
                | neqop
                | gteop
                | lteop;
            MarkTransient(InfixOp);

            //PostfixOp.Rule = ImplyPrecedenceHere(Precedence.UnaryPostFix) + percentop;
            // ImplyPrecedenceHere doesn't seem to work for this rule, but postfix has such a high priority shift will nearly always be the correct action
            PostfixOp.Rule = PreferShiftHere() + percentop;
            MarkTransient(PostfixOp);
            #endregion

            #region References

            Reference.Rule = ReferenceItem
                | ReferenceFunctionCall
                | OpenParen + Reference + PreferShiftHere() + CloseParen
                | Prefix + ReferenceItem
                | DynamicDataExchange
                ;

            ReferenceFunctionCall.Rule =
                  Reference + colon + Reference
                | Reference + intersectop + Reference
                | OpenParen + Union + CloseParen
                | RefFunctionName + Arguments + CloseParen
                //| ConditionalRefFunctionName + Arguments + CloseParen
                ;

            RefFunctionName.Rule = ExcelRefFunctionToken | ExcelConditionalRefFunctionToken;

            Union.Rule = MakePlusRule(Union, comma, Reference);

            ReferenceItem.Rule =
                Cell
                | NamedRange
                | VRange
                | HRange
                | RefError
                | UDFunctionCall
                | StructureReference
                ;
            MarkTransient(ReferenceItem);

            UDFunctionCall.Rule = UDFName + Arguments + CloseParen;
            UDFName.Rule = UDFToken;

            VRange.Rule = VRangeToken;
            HRange.Rule = HRangeToken;

            Cell.Rule = CellToken;

            File.Rule = FileToken
				| EnclosedInBracketsToken
				| FilePathWindowsToken + EnclosedInBracketsToken
                ;

            DynamicDataExchange.Rule = File + exclamationMark + SingleQuotedStringToken;

            NamedRange.Rule = NameToken | NamedRangeCombinationToken;
            Prefix.Rule =
                SheetToken
                | QuoteS + SheetQuotedToken
                | File + SheetToken
                | QuoteS + File + SheetQuotedToken
                | File + exclamationMark
                | MultipleSheetsToken
                | File + MultipleSheetsToken
                ;

            StructureReferenceColumnOrKeyword.Rule =
                  OpenSquareParen + SRColumnToken + CloseSquareParen
                | OpenSquareParen + NameToken + CloseSquareParen
                | EnclosedInBracketsToken;

            //StructureReferenceKeyword.Rule = EnclosedInBracketsToken;

            StructureReferenceTable.Rule = NameToken;

            StructureReferenceExpression.Rule =
                  StructureReferenceColumnOrKeyword
                | StructureReferenceColumnOrKeyword + colon + StructureReferenceColumnOrKeyword
                | StructureReferenceColumnOrKeyword + comma + StructureReferenceColumnOrKeyword
                | StructureReferenceColumnOrKeyword + comma + StructureReferenceColumnOrKeyword + colon + StructureReferenceColumnOrKeyword
                | StructureReferenceColumnOrKeyword + comma + StructureReferenceColumnOrKeyword + comma + StructureReferenceColumnOrKeyword
                | StructureReferenceColumnOrKeyword + comma + StructureReferenceColumnOrKeyword + comma + StructureReferenceColumnOrKeyword + colon + StructureReferenceColumnOrKeyword
                ;

            StructureReference.Rule =
                  StructureReferenceColumnOrKeyword
                | OpenSquareParen + StructureReferenceExpression + CloseSquareParen
                | StructureReferenceTable + StructureReferenceColumnOrKeyword
                | StructureReferenceTable + OpenSquareParen + StructureReferenceExpression + CloseSquareParen
                ;
            #endregion

            #region Arrays
            ConstantArray.Rule = OpenCurlyParen + ArrayColumns + CloseCurlyParen;

            ArrayColumns.Rule = MakePlusRule(ArrayColumns, semicolon, ArrayRows);
            ArrayRows.Rule = MakePlusRule(ArrayRows, comma, ArrayConstant);

            ArrayConstant.Rule = Constant | PrefixOp + Number | RefError;
            #endregion

            #endregion

            #region 5-Operator Precedence            
            // Some of these operators are neutral associative instead of left associative,
            // but this ensures a consistent parse tree. As a lot of code is "hardcoded" onto the specific
            // structure of the parse tree, we like consistency.
            RegisterOperators(Precedence.Comparison, Associativity.Left, eqop, ltop, gtop, lteop, gteop, neqop);
            RegisterOperators(Precedence.Concatenation, Associativity.Left, concatop);
            RegisterOperators(Precedence.Addition, Associativity.Left, plusop, minop);
            RegisterOperators(Precedence.Multiplication, Associativity.Left, mulop, divop);
            RegisterOperators(Precedence.Exponentiation, Associativity.Left, expop);
            RegisterOperators(Precedence.UnaryPostFix, Associativity.Left, percentop);
            RegisterOperators(Precedence.Union, Associativity.Left, comma);
            RegisterOperators(Precedence.Intersection, Associativity.Left, intersectop);
            RegisterOperators(Precedence.Range, Associativity.Left, colon);

            //RegisterOperators(Precedence.ParameterSeparator, comma);

            #endregion
        }

        

        #region Precedence and Priority constants
        // Source: https://support.office.com/en-us/article/Calculation-operators-and-precedence-48be406d-4975-4d31-b2b8-7af9e0e2878a
        // Could also be an enum, but this way you don't need int casts
        private static class Precedence
        {
            // Don't use priority 0, Irony seems to view it as no priority set
            public const int Comparison = 1;
            public const int Concatenation = 2;
            public const int Addition = 3;
            public const int Multiplication = 4;
            public const int Exponentiation = 5;
            public const int UnaryPostFix = 6;
            public const int UnaryPreFix = 7;
            //public const int Reference = 8;
            public const int Union = 9;
            public const int Intersection = 10;
            public const int Range = 11;
        }

        // Terminal priorities, indicates to lexer which token it should pick when multiple tokens can match
        // E.g. "A1" is both a CellToken and NamedRange, pick celltoken because it has a higher priority
        // E.g. "A1Blah" Is Both a CellToken + NamedRange, NamedRange and NamedRangeCombination, pick NamedRangeCombination
        private static class TerminalPriority
        {
            // Irony Low value
            //public const int Low = -1000;

            public const int SRColumn = -900;

            public const int Name = -800;
            public const int ReservedName = -700;

            public const int FileName = -500;

            public const int SingleQuotedString = -100;

            // Irony Normal value, default value
            //public const int Normal = 0;
            public const int Bool = 0;

            public const int MultipleSheetsToken = 100;

            // Irony High value
            //public const int High = 1000;

            public const int CellToken = 1000;

            public const int NamedRangeCombination = 1100;

            public const int UDF = 1150;

            public const int ExcelFunction = 1200;
            public const int ExcelRefFunction = 1200;
            public const int FileNameNumericToken = 1200;
            public const int SheetToken = 1200;
            public const int SheetQuotedToken = 1200;
        }
        #endregion

        private static string[] excelFunctionList => Properties.Resources.ExcelBuiltinFunctionList.Split(new [] {'\n', '\r'}, StringSplitOptions.RemoveEmptyEntries);
    }

    #region Names
    /// <summary>
    /// Collection of names used for terminals and non-terminals in the Excel Formula Grammar.
    /// </summary>
    /// <remarks>
    /// Using these is strongly recommended, as these will change when breaking changes occur.
    /// It also allows you to see which code works on what grammar constructs.
    /// </remarks>
    // Keep these constants instead of methods/properties, since that allows them to be used in switch statements.
    public static class GrammarNames
    {
        #region Non-Terminals
        public const string Argument = "Argument";
        public const string Arguments = "Arguments";
        public const string ArrayColumns = "ArrayColumns";
        public const string ArrayConstant = "ArrayConstant";
        public const string ArrayFormula = "ArrayFormula";
        public const string ArrayRows = "ArrayRows";
        public const string Bool = "Bool";
        public const string Cell = "Cell";
        public const string Constant = "Constant";
        public const string ConstantArray = "ConstantArray";
        public const string DynamicDataExchange = "DynamicDataExchange";
        public const string EmptyArgument = "EmptyArgument";
        public const string Error = "Error";
        public const string ExcelFunction = "ExcelFunction";
        public const string File = "File";
        public const string Formula = "Formula";
        public const string FormulaWithEq = "FormulaWithEq";
        public const string FunctionCall = "FunctionCall";
        public const string FunctionName = "FunctionName";
        public const string HorizontalRange = "HRange";
        public const string NamedRange = "NamedRange";
        public const string Number = "Number";
        public const string Prefix = "Prefix";
        public const string QuotedFileSheet = "QuotedFileSheet";
        public const string Range = "Range";
        public const string Reference = "Reference";
        public const string ReferenceFunctionCall = "ReferenceFunctionCall";
        public const string RefError = "RefError";
        public const string RefFunctionName = "RefFunctionName";
        public const string ReservedName = "ReservedName";
        public const string Sheet = "Sheet";
        public const string StructureReference = "StructureReference";
        public const string StructureReferenceColumnOrKeyword = "StructureReferenceColumnOrKeyword";
        public const string StructureReferenceExpression = "StructureReferenceExpression";
        public const string StructureReferenceTable = "StructureReferenceTable";
        public const string Text = "Text";
        public const string UDFName = "UDFName";
        public const string UDFunctionCall = "UDFunctionCall";
        public const string Union = "Union";
        public const string VerticalRange = "VRange";
        #endregion

        #region Transient Non-Terminals
        public const string TransientStart = "Start";
        public const string TransientInfixOp = "InfixOp";
        public const string TransientPostfixOp = "PostfixOp";
        public const string TransientPrefixOp = "PrefixOp";
        public const string TransientReferenceItem = "ReferenceItem";
        #endregion

        #region Terminals
        public const string TokenBool = "BoolToken";
        public const string TokenCell = "CellToken";
        public const string TokenEmptyArgument = "EmptyArgumentToken";
        public const string TokenError = "ErrorToken";
        public const string TokenExcelRefFunction = "ExcelRefFunctionToken";
        public const string TokenExcelConditionalRefFunction = "ExcelConditionalRefFunctionToken";
        public const string TokenFilePathWindows = "FilePathWindowsToken";
        public const string TokenEnclosedInBrackets = "EnclosedInBracketsToken";
        public const string TokenFileNameNumeric = "FileNameNumericToken";
        public const string TokenHRange = "HRangeToken";
        public const string TokenIntersect = "INTERSECT";
        public const string TokenMultipleSheets = "MultipleSheetsToken";
        public const string TokenName = "NameToken";
        public const string TokenNamedRangeCombination = "NamedRangeCombinationToken";
        public const string TokenNumber = "NumberToken";
        public const string TokenRefError = "RefErrorToken";
        public const string TokenReservedName = "ReservedNameToken";
        public const string TokenSingleQuotedString = "SingleQuotedString";
        public const string TokenSheet = "SheetNameToken";
        public const string TokenSheetQuoted = "SheetNameQuotedToken";
        public const string TokenSRTableName = "SRTableName";
        public const string TokenSRKeyword = "SRKeyword";
        public const string TokenSRColumn = "SRColumn";
        public const string TokenSREnclosedColumn = "SREnclosedColumn";
        public const string TokenText = "TextToken";
        public const string TokenUDF = "UDFToken";
        public const string TokenUnionOperator = ",";
        public const string TokenVRange = "VRangeToken";

        #endregion

    }
    #endregion
}
