using Irony.Parsing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;

namespace XLParser.Web.XLParserVersions.v170
{
    /// <summary>
    /// Contains the XLParser grammar
    /// </summary>
    [Language("Excel Formulas", "1.7.0", "Grammar for Excel Formulas")]
    public class ExcelFormulaGrammar : Grammar
    {
        #region 1-Terminals

        #region Symbols and operators

        public Terminal at => ToTerm("@");
        public Terminal comma => ToTerm(",");
        public Terminal colon => ToTerm(":");
        public Terminal hash => ToTerm("#");
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
        // Attention: ImpliedSymbolTerminal seems to break if you assign it a priority, and its default priority is low
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

        public Terminal BoolToken { get; } = new RegexBasedTerminal(GrammarNames.TokenBool, "TRUE|FALSE", "T", "F")
        {
            Priority = TerminalPriority.Bool
        };

        public Terminal NumberToken { get; } = new NumberLiteral(GrammarNames.TokenNumber, NumberOptions.None)
        {
            DefaultIntTypes = new[] { TypeCode.Int32, TypeCode.Int64, NumberLiteral.TypeCodeBigInt }
        };

        public Terminal TextToken { get; } = new StringLiteral(GrammarNames.TokenText, "\"",
            StringOptions.AllowsDoubledQuote | StringOptions.AllowsLineBreak | StringOptions.NoEscapes);

        public Terminal SingleQuotedStringToken { get; } = new StringLiteral(GrammarNames.TokenSingleQuotedString, "'",
            StringOptions.AllowsDoubledQuote | StringOptions.AllowsLineBreak | StringOptions.NoEscapes)
        { Priority = TerminalPriority.SingleQuotedString };

        public Terminal ErrorToken { get; } = new RegexBasedTerminal(GrammarNames.TokenError, "#NULL!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#N/A|#GETTING_DATA|#SPILL!", "#");
        public Terminal RefErrorToken => ToTerm("#REF!", GrammarNames.TokenRefError);

        #endregion

        #region Functions
        private const string SpecialUdfChars = "¡¢£¤¥¦§¨©«¬­®¯°±²³´¶·¸¹»¼½¾¿×÷"; // Non-word characters from ISO 8859-1 that are allowed in VBA identifiers
        private const string AllUdfChars = SpecialUdfChars + @"\\.\w";
        private const string UdfPrefixRegex = @"('[^<>""/\|?*]+\.xla'!|_xll\.)";

        // The following regex uses the rather exotic feature Character Class Subtraction
        // https://docs.microsoft.com/en-us/dotnet/standard/base-types/character-classes-in-regular-expressions#CharacterClassSubtraction
        private static readonly string UdfTokenRegex = $@"([{AllUdfChars}-[CcRr]]|{UdfPrefixRegex}[{AllUdfChars}]|{UdfPrefixRegex}?[{AllUdfChars}]{{2,1023}})\(";

        public Terminal UDFToken { get; } = new RegexBasedTerminal(GrammarNames.TokenUDF, UdfTokenRegex) { Priority = TerminalPriority.UDF };

        public Terminal ExcelRefFunctionToken { get; } = new RegexBasedTerminal(GrammarNames.TokenExcelRefFunction, "(INDEX|OFFSET|INDIRECT)\\(", "I", "O")
        { Priority = TerminalPriority.ExcelRefFunction };

        public Terminal ExcelConditionalRefFunctionToken { get; } = new RegexBasedTerminal(GrammarNames.TokenExcelConditionalRefFunction, "(IF|CHOOSE)\\(", "I", "C")
        { Priority = TerminalPriority.ExcelRefFunction };

        public Terminal ExcelFunction { get; } = new WordsTerminal(GrammarNames.ExcelFunction,  excelFunctionList.Select(f => f + '('))
        { Priority = TerminalPriority.ExcelFunction };

        // Using this instead of Empty allows a more accurate tree
        public Terminal EmptyArgumentToken { get; } = new ImpliedSymbolTerminal(GrammarNames.TokenEmptyArgument);

        #endregion

        #region References and names

        private const string ColumnPattern = @"(?:[A-W][A-Z]{1,2}|X[A-E][A-Z]|XF[A-D]|[A-Z]{1,2})";
        private const string RowPattern = @"(?:104857[0-6]|10485[0-6][0-9]|1048[0-4][0-9]{2}|104[0-7][0-9]{3}|10[0-3][0-9]{4}|[1-9][0-9]{1,5}|[1-9])";

        private static readonly string[] ColumnPrefix = Enumerable.Range('A', 'Z' - 'A' + 1).Select(c => char.ToString((char)c)).Concat(new[] { "$" }).ToArray();
        private static readonly string[] RowPrefix = Enumerable.Range('1', '9' - '1' + 1).Select(c => char.ToString((char)c)).Concat(new[] { "$" }).ToArray();

        public Terminal VRangeToken { get; } = new RegexBasedTerminal(GrammarNames.TokenVRange, "[$]?" + ColumnPattern + ":[$]?" + ColumnPattern, ColumnPrefix);
        public Terminal HRangeToken { get; } = new RegexBasedTerminal(GrammarNames.TokenHRange, "[$]?" + RowPattern + ":[$]?" + RowPattern, RowPrefix);

        private const string CellTokenRegex = "[$]?" + ColumnPattern + "[$]?" + RowPattern;
        public Terminal CellToken { get; } = new RegexBasedTerminal(GrammarNames.TokenCell, CellTokenRegex, ColumnPrefix)
        { Priority = TerminalPriority.CellToken };

        private static readonly HashSet<UnicodeCategory> UnicodeLetterCategories = new HashSet<UnicodeCategory>
        {
            UnicodeCategory.UppercaseLetter,
            UnicodeCategory.LowercaseLetter,
            UnicodeCategory.TitlecaseLetter,
            UnicodeCategory.ModifierLetter,
            UnicodeCategory.OtherLetter
        };

        // 48718 letters, but it allows parser to from tokens starting with digits, parentheses, operators...
        private static readonly string[] UnicodeLetters = Enumerable.Range(0, ushort.MaxValue).Where(codePoints => UnicodeLetterCategories.Contains(CharUnicodeInfo.GetUnicodeCategory((char)codePoints))).Select(codePoint => char.ToString((char)codePoint)).ToArray();
        private static readonly string[] NameStartCharPrefix = UnicodeLetters.Concat(new[] { @"\", "_" }).ToArray();

        // Start with a letter or underscore, continue with word character (letters, numbers and underscore), dot or question mark 
        private const string NameStartCharRegex = @"[\p{L}\\_]";
        private const string NameValidCharacterRegex = @"[\w\\_\.\?€]";

        public Terminal NameToken { get; } = new RegexBasedTerminal(GrammarNames.TokenName, NameStartCharRegex + NameValidCharacterRegex + "*", NameStartCharPrefix)
        { Priority = TerminalPriority.Name };

        // Words that are valid names, but are disallowed by Excel. E.g. "A1" is a valid name, but it is not because it is also a cell reference.
        // If we ever parse R1C1 references, make sure to include them here
        // TODO: Add all function names here

        private const string NamedRangeCombinationRegex =
              "((TRUE|FALSE)" + NameValidCharacterRegex + "+)"
            // \w is equivalent to [\p{Ll}\p{Lu}\p{Lt}\p{Lo}\p{Lm}\p{Nd}\p{Pc}], we want the decimal left out here because otherwise "A11" would be a combination token
            + "|(" + CellTokenRegex + @"[\p{Ll}\p{Lu}\p{Lt}\p{Lo}\p{Lm}\p{Pc}\\_\.\?]" + NameValidCharacterRegex + "*)"
            // allow large cell references (e.g. A1048577) as named range
            + "|(" + ColumnPattern + @"(104857[7-9]|10485[89][0-9]|1048[6-9][0-9]{2}|1049[0-9]{3}|10[5-9][0-9]{4}|1[1-9][0-9]{5}|[2-9][0-9]{6}|d{8,})" + NameValidCharacterRegex + "*)"
            ;

        // To prevent e.g. "A1A1" being parsed as 2 cell tokens
        public Terminal NamedRangeCombinationToken { get; } = new RegexBasedTerminal(GrammarNames.TokenNamedRangeCombination, NamedRangeCombinationRegex,
                ColumnPrefix.Concat(new[] { "T", "F" }).ToArray())
        { Priority = TerminalPriority.NamedRangeCombination };

        public Terminal ReservedNameToken = new RegexBasedTerminal(GrammarNames.TokenReservedName, @"_xlnm\.[a-zA-Z_]+", "_")
        { Priority = TerminalPriority.ReservedName };

        #region Structured References
        private const string SRSpecifierRegex = @"#(All|Data|Headers|Totals|This Row)";
        public Terminal SRSpecifierToken = new RegexBasedTerminal(GrammarNames.TokenSRSpecifier, SRSpecifierRegex, "#")
        { Priority = TerminalPriority.StructuredReference };

        private const string SRColumnRegex = @"(?:[^\[\]'#@]|(?:'['\[\]#@]))+";
        public Terminal SRColumnToken = new RegexBasedTerminal(GrammarNames.TokenSRColumn, SRColumnRegex)
        { Priority = TerminalPriority.StructuredReference };
        #endregion

        #region Prefixes
        private const string mustBeQuotedInSheetName = @"\(\);{}#""=<>&+\-*/\^%, ";
        private const string notSheetNameChars = @"'*\[\]\\:/?";
        //const string singleQuotedContent = @"\w !@#$%^&*()\-\+={}|:;<>,\./\?" + "\\\"";
        //const string sheetRegEx = @"(([\w\.]+)|('([" + singleQuotedContent + @"]|'')+'))!";
        private static readonly string normalSheetName = $"[^{notSheetNameChars}{mustBeQuotedInSheetName}]+";
        private static readonly string quotedSheetName = $"([^{notSheetNameChars}]|'')*";
        //private static readonly string sheetRegEx = $"(({normalSheetName})|('{quotedSheetName}'))!";

        public Terminal SheetToken = new RegexBasedTerminal(GrammarNames.TokenSheet, $"{normalSheetName}!")
        { Priority = TerminalPriority.SheetToken };

        public Terminal SheetQuotedToken = new RegexBasedTerminal(GrammarNames.TokenSheetQuoted, $"{quotedSheetName}'!")
        { Priority = TerminalPriority.SheetQuotedToken };

        private static readonly string multiSheetRegex = $"{normalSheetName}:{normalSheetName}!";
        private static readonly string multiSheetQuotedRegex = $"{quotedSheetName}:{quotedSheetName}'!";
        public Terminal MultipleSheetsToken = new RegexBasedTerminal(GrammarNames.TokenMultipleSheets, multiSheetRegex)
        { Priority = TerminalPriority.MultipleSheetsToken };
        public Terminal MultipleSheetsQuotedToken = new RegexBasedTerminal(GrammarNames.TokenMultipleSheetsQuoted, multiSheetQuotedRegex)
        { Priority = TerminalPriority.MultipleSheetsToken };

        private const string fileNameNumericRegex = @"\[[0-9]+\](?!,)(?=.*!)";
        public Terminal FileNameNumericToken = new RegexBasedTerminal(GrammarNames.TokenFileNameNumeric, fileNameNumericRegex, "[")
        { Priority = TerminalPriority.FileNameNumericToken };

        private const string fileNameInBracketsRegex = @"\[[^\[\]]+\](?!,)(?=.*!)";
        public Terminal FileNameEnclosedInBracketsToken { get; } = new RegexBasedTerminal(GrammarNames.TokenFileNameEnclosedInBrackets, fileNameInBracketsRegex, "[")
        { Priority = TerminalPriority.FileName };

        // Source: https://stackoverflow.com/a/14632579
        private const string fileNameRegex = @"[^\.\\\[\]]+\..{1,4}";
        public Terminal FileName { get; } = new RegexBasedTerminal(GrammarNames.TokenFileName, fileNameRegex)
        { Priority = TerminalPriority.FileName };

        // Source: http://stackoverflow.com/a/6416209/572635
        private const string windowsFilePathRegex = @"(?:[a-zA-Z]:|\\?\\?[\w\-.$ @]+)\\(([^<>\"" /\|?*\\']|( |''))*\\)*";
        private const string urlPathRegex = @"http(s?)\:\/\/[0-9a-zA-Z]([-.\w]*[0-9a-zA-Z])*(:(0-9)*)*[/]([a-zA-Z0-9\-\.\?\,\'+&%\$#_ ()]*[/])*";
        private const string filePathRegex = @"(" + windowsFilePathRegex + @"|" + urlPathRegex + @")";
        public Terminal FilePathToken { get; } = new RegexBasedTerminal(GrammarNames.TokenFilePath, filePathRegex)
        { Priority = TerminalPriority.FileNamePath };

        #endregion

        #endregion

        #endregion

        #region 2-NonTerminals
        // Most non-terminals are first defined here, so they can be used anywhere in the rules
        // Otherwise you can only use non-terminals that have been defined previously

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
        public NonTerminal MultiRangeFormula{ get; } = new NonTerminal(GrammarNames.MultiRangeFormula);
        public NonTerminal NamedRange{ get; } = new NonTerminal(GrammarNames.NamedRange);
        public NonTerminal Number{ get; } = new NonTerminal(GrammarNames.Number);
        public NonTerminal PostfixOp{ get; } = new NonTerminal(GrammarNames.TransientPostfixOp);
        public NonTerminal Prefix{ get; } = new NonTerminal(GrammarNames.Prefix);
        public NonTerminal PrefixOp{ get; } = new NonTerminal(GrammarNames.TransientPrefixOp);
        public NonTerminal QuotedFileSheet{ get; } = new NonTerminal(GrammarNames.QuotedFileSheet);
        public NonTerminal Reference{ get; } = new NonTerminal(GrammarNames.Reference);
        public NonTerminal ReferenceItem{ get; } = new NonTerminal(GrammarNames.TransientReferenceItem);
        public NonTerminal ReferenceFunctionCall{ get; } = new NonTerminal(GrammarNames.ReferenceFunctionCall);
        public NonTerminal RefError{ get; } = new NonTerminal(GrammarNames.RefError);
        public NonTerminal RefFunctionName{ get; } = new NonTerminal(GrammarNames.RefFunctionName);
        public NonTerminal ReservedName{ get; } = new NonTerminal(GrammarNames.ReservedName);
        public NonTerminal Sheet{ get; } = new NonTerminal(GrammarNames.Sheet);
        public NonTerminal Start{ get; } = new NonTerminal(GrammarNames.TransientStart);
        public NonTerminal StructuredReference { get; } = new NonTerminal(GrammarNames.StructuredReference);
        public NonTerminal StructuredReferenceColumn { get; } = new NonTerminal(GrammarNames.StructuredReferenceColumn);
        public NonTerminal StructuredReferenceExpression { get; } = new NonTerminal(GrammarNames.StructuredReferenceExpression);
        public NonTerminal StructuredReferenceSpecifier { get; } = new NonTerminal(GrammarNames.StructuredReferenceSpecifier);
        public NonTerminal StructuredReferenceQualifier { get; } = new NonTerminal(GrammarNames.StructuredReferenceQualifier);
        public NonTerminal Text{ get; } = new NonTerminal(GrammarNames.Text);
        public NonTerminal UDFName{ get; } = new NonTerminal(GrammarNames.UDFName);
        public NonTerminal UDFunctionCall{ get; } = new NonTerminal(GrammarNames.UDFunctionCall);
        public NonTerminal Union{ get; } = new NonTerminal(GrammarNames.Union);
        public NonTerminal VRange{ get; } = new NonTerminal(GrammarNames.VerticalRange);
        #endregion

        public ExcelFormulaGrammar()
        {

            #region Punctuation
            MarkPunctuation(OpenParen, CloseParen);
            MarkPunctuation(OpenCurlyParen, CloseCurlyParen);
            #endregion

            #region Rules

            #region Base rules
            Root = Start;

            Start.Rule =
                  FormulaWithEq
                | Formula
                | ArrayFormula
                | MultiRangeFormula
                ;
            MarkTransient(Start);

            ArrayFormula.Rule = OpenCurlyParen + eqop + Formula + CloseCurlyParen;

            MultiRangeFormula.Rule = eqop + Union;

            FormulaWithEq.Rule = eqop + Formula;

            Formula.Rule =
                  Reference + ReduceHere()
                | Constant
                | FunctionCall
                | ConstantArray
                | OpenParen + Formula + CloseParen
                | ReservedName
                ;

            ReservedName.Rule = ReservedNameToken;

            Constant.Rule =
                  Number
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

            EmptyArgument.Rule = EmptyArgumentToken;
            Argument.Rule = Formula | EmptyArgument;

            PrefixOp.Rule =
                  ImplyPrecedenceHere(Precedence.UnaryPreFix) + plusop
                | ImplyPrecedenceHere(Precedence.UnaryPreFix) + minop
                | ImplyPrecedenceHere(Precedence.UnaryPreFix) + at;
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

            // ImplyPrecedenceHere doesn't seem to work for this rule, but postfix has such a high priority shift will nearly always be the correct action
            PostfixOp.Rule = PreferShiftHere() + percentop;
            MarkTransient(PostfixOp);
            #endregion

            #region References

            Reference.Rule =
                  ReferenceItem
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
                | Reference + hash
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
                | StructuredReference
                ;
            MarkTransient(ReferenceItem);

            UDFunctionCall.Rule = UDFName + Arguments + CloseParen;
            UDFName.Rule = UDFToken;

            VRange.Rule = VRangeToken;
            HRange.Rule = HRangeToken;

            Cell.Rule = CellToken;

            File.Rule =
                  FileNameNumericToken
                | FileNameEnclosedInBracketsToken
                | FilePathToken + FileNameEnclosedInBracketsToken
                | FilePathToken + FileName
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
                | QuoteS + MultipleSheetsQuotedToken
                | File + MultipleSheetsToken
                | QuoteS + File + MultipleSheetsQuotedToken
                | RefErrorToken
                ;

            StructuredReferenceQualifier.Rule = NameToken;

            StructuredReferenceSpecifier.Rule =
                  SRSpecifierToken
                | at
                | OpenSquareParen + SRSpecifierToken + CloseSquareParen;

            StructuredReferenceColumn.Rule =
                  SRColumnToken
                | OpenSquareParen + SRColumnToken + CloseSquareParen;

            StructuredReferenceExpression.Rule =
                  StructuredReferenceColumn
                | StructuredReferenceColumn + colon + StructuredReferenceColumn
                | at + StructuredReferenceColumn
                | at + StructuredReferenceColumn + colon + StructuredReferenceColumn
                | StructuredReferenceSpecifier
                | StructuredReferenceSpecifier + comma + StructuredReferenceColumn
                | StructuredReferenceSpecifier + comma + StructuredReferenceColumn + colon + StructuredReferenceColumn
                | StructuredReferenceSpecifier + comma + StructuredReferenceSpecifier
                | StructuredReferenceSpecifier + comma + StructuredReferenceSpecifier + comma + StructuredReferenceColumn
                | StructuredReferenceSpecifier + comma + StructuredReferenceSpecifier + comma + StructuredReferenceColumn + colon + StructuredReferenceColumn
                ;

            StructuredReference.Rule =
                  OpenSquareParen + StructuredReferenceExpression + CloseSquareParen
                | StructuredReferenceQualifier + OpenSquareParen + CloseSquareParen
                | StructuredReferenceQualifier + OpenSquareParen + StructuredReferenceExpression + CloseSquareParen
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
            RegisterOperators(Precedence.UnaryPostFix, Associativity.Left, percentop, hash);
            RegisterOperators(Precedence.UnaryPreFix, Associativity.Left, at);
            RegisterOperators(Precedence.Union, Associativity.Left, comma);
            RegisterOperators(Precedence.Intersection, Associativity.Left, intersectop);
            RegisterOperators(Precedence.Range, Associativity.Left, colon);
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
        // E.g. "A1" is both a CellToken and NamedRange, pick cell token because it has a higher priority
        // E.g. "A1Blah" Is Both a CellToken + NamedRange, NamedRange and NamedRangeCombination, pick NamedRangeCombination
        private static class TerminalPriority
        {
            // Irony Low value
            //public const int Low = -1000;

            public const int Name = -800;
            public const int ReservedName = -700;

            public const int StructuredReference = -500;

            public const int FileName = -500;
            public const int FileNamePath = -800;

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

        private static string[] excelFunctionList => GetExcelFunctionList();
        private static string[] GetExcelFunctionList()
        {
            var resource = Properties.Resources.ExcelBuiltinFunctionList_v170;
            using (var sr = new StringReader(resource))
                return sr.ReadToEnd().Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
        }

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
        public const string MultiRangeFormula = "MultiRangeFormula";
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
        public const string StructuredReference = "StructuredReference";
        public const string StructuredReferenceColumn = "StructuredReferenceColumn";
        public const string StructuredReferenceExpression = "StructuredReferenceExpression";
        public const string StructuredReferenceSpecifier = "StructuredReferenceSpecifier";
        public const string StructuredReferenceQualifier = "StructuredReferenceQualifier";
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
        public const string TokenFilePath = "FilePathToken";
        public const string TokenFileName = "FileNameToken";
        public const string TokenFileNameEnclosedInBrackets = "FileNameEnclosedInBracketsToken";
        public const string TokenFileNameNumeric = "FileNameNumericToken";
        public const string TokenHRange = "HRangeToken";
        public const string TokenIntersect = "INTERSECT";
        public const string TokenMultipleSheets = "MultipleSheetsToken";
        public const string TokenMultipleSheetsQuoted = "MultipleSheetsQuotedToken";
        public const string TokenName = "NameToken";
        public const string TokenNamedRangeCombination = "NamedRangeCombinationToken";
        public const string TokenNumber = "NumberToken";
        public const string TokenRefError = "RefErrorToken";
        public const string TokenReservedName = "ReservedNameToken";
        public const string TokenSingleQuotedString = "SingleQuotedString";
        public const string TokenSheet = "SheetNameToken";
        public const string TokenSheetQuoted = "SheetNameQuotedToken";
        public const string TokenSRColumn = "SRColumnToken";
        public const string TokenSRSpecifier = "SRSpecifierToken";
        public const string TokenText = "TextToken";
        public const string TokenUDF = "UDFToken";
        public const string TokenUnionOperator = ",";
        public const string TokenVRange = "VRangeToken";
        #endregion

    }
    #endregion
}
