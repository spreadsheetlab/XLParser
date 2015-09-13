using Irony.Parsing;
using System;
using System.Collections.Generic;

namespace XLParser.Web.XLParserVersions.v114
{
    /// <summary>
    /// Contains the XLParser grammar
    /// </summary>
    [Language("Excel Formulas", "1.1.3", "Grammar for Excel Formulas")]
    public class ExcelFormulaGrammar : Grammar
    {
        public ExcelFormulaGrammar() : base(false)
        {
            #region 1-Terminals

            #region Symbols and operators
            var comma = ToTerm(",");
            var colon = ToTerm(":");
            var semicolon = ToTerm(";");
            var OpenParen = ToTerm("(");
            var CloseParen = ToTerm(")");
            var CloseSquareParen = ToTerm("]");
            var OpenSquareParen = ToTerm("[");
            var exclamationMark = ToTerm("!");
            var CloseCurlyParen = ToTerm("}");
            var OpenCurlyParen = ToTerm("{");

            var mulop = ToTerm("*");
            var plusop = ToTerm("+");
            var divop = ToTerm("/");
            var minop = ToTerm("-");
            var concatop = ToTerm("&");
            var expop = ToTerm("^");
            // Intersect op is a single space, which cannot be parsed normally so we need an ImpliedSymbolTerminal
            // Attention: ImpliedSymbolTerminal seems to break if you assign it a priority, and it's default priority is low
            var intersectop = new ImpliedSymbolTerminal(GrammarNames.TokenIntersect);

            var percentop = ToTerm("%");

            var gtop = ToTerm(">");
            var eqop = ToTerm("=");
            var ltop = ToTerm("<");
            var neqop = ToTerm("<>");
            var gteop = ToTerm(">=");
            var lteop = ToTerm("<=");
            #endregion

            #region Literals
            var BoolToken = new RegexBasedTerminal(GrammarNames.TokenBool, "TRUE|FALSE");
            BoolToken.Priority = TerminalPriority.Bool;

            var NumberToken = new NumberLiteral(GrammarNames.TokenNumber, NumberOptions.None);
            NumberToken.DefaultIntTypes = new TypeCode[] { TypeCode.Int32, TypeCode.Int64, NumberLiteral.TypeCodeBigInt };

            var TextToken = new StringLiteral(GrammarNames.TokenText, "\"", StringOptions.AllowsDoubledQuote | StringOptions.AllowsLineBreak);

            var ErrorToken = new RegexBasedTerminal(GrammarNames.TokenError, "#NULL!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#N/A");
            var RefErrorToken = ToTerm("#REF!", GrammarNames.TokenRefError);
            #endregion

            #region Functions

            var UDFToken = new RegexBasedTerminal(GrammarNames.TokenUDF, @"(_xll\.)?[\w\\.]+\(");
            UDFToken.Priority = TerminalPriority.UDF;

            var ExcelRefFunctionToken = new RegexBasedTerminal(GrammarNames.TokenExcelRefFunction, "(INDEX|OFFSET|INDIRECT)\\(");
            ExcelRefFunctionToken.Priority = TerminalPriority.ExcelRefFunction;
            
            var ExcelConditionalRefFunctionToken = new RegexBasedTerminal(GrammarNames.TokenExcelConditionalRefFunction, "(IF|CHOOSE)\\(");
            ExcelConditionalRefFunctionToken.Priority = TerminalPriority.ExcelRefFunction;

            var ExcelFunction = new RegexBasedTerminal(GrammarNames.ExcelFunction, "(" + String.Join("|", excelFunctionList)  +")\\(");
            ExcelFunction.Priority = TerminalPriority.ExcelFunction;

            // Using this instead of Empty allows a more accurate trees
            var EmptyArgumentToken = new ImpliedSymbolTerminal(GrammarNames.TokenEmptyArgument);

            #endregion

            #region References and names

            var VRangeToken = new RegexBasedTerminal(GrammarNames.TokenVRange, "[$]?[A-Z]{1,4}:[$]?[A-Z]{1,4}");
            var HRangeToken = new RegexBasedTerminal(GrammarNames.TokenHRange, "[$]?[1-9][0-9]*:[$]?[1-9][0-9]*");
            
            const string CellTokenRegex = "[$]?[A-Z]{1,4}[$]?[1-9][0-9]*";
            var CellToken = new RegexBasedTerminal(GrammarNames.TokenCell, CellTokenRegex);
            CellToken.Priority = TerminalPriority.CellToken;

            const string NamedRangeRegex = @"[A-Za-z\\_][\w\.]*";
            var NamedRangeToken = new RegexBasedTerminal(GrammarNames.TokenNamedRange, NamedRangeRegex);
            NamedRangeToken.Priority = TerminalPriority.NamedRange;

            // To prevent e.g. "A1A1" being parsed as 2 celltokens
            var NamedRangeCombinationToken = new RegexBasedTerminal(GrammarNames.TokenNamedRangeCombination, "(TRUE|FALSE|" + CellTokenRegex + ")" + NamedRangeRegex);
            NamedRangeCombinationToken.Priority = TerminalPriority.NamedRangeCombination;

            const string mustBeQuotedInSheetName = @"\(\);{}#""=<>&+\-*/\^%, ";
            const string notSheetNameChars = @"'*\[\]\\:/?";
            //const string singleQuotedContent = @"\w !@#$%^&*()\-\+={}|:;<>,\./\?" + "\\\"";
            //const string sheetRegEx = @"(([\w\.]+)|('([" + singleQuotedContent + @"]|'')+'))!";
            const string normalSheetName = "[^" + notSheetNameChars + mustBeQuotedInSheetName + "]+";
            const string quotedSheetName = "([^" + notSheetNameChars +  "]|'')+";
            const string sheetRegEx = "((" + normalSheetName + ")|('" + quotedSheetName + "'))!";

            var SheetToken = new RegexBasedTerminal(GrammarNames.TokenSheet, sheetRegEx);
            SheetToken.Priority = TerminalPriority.SheetToken;

            var multiSheetRegex = String.Format("(({0}:{0})|('{1}:{1}'))!", normalSheetName, quotedSheetName);
            var MultipleSheetsToken = new RegexBasedTerminal(GrammarNames.TokenMultipleSheets, multiSheetRegex);
            MultipleSheetsToken.Priority = TerminalPriority.MultipleSheetsToken;

            var FileToken = new RegexBasedTerminal(GrammarNames.TokenFileNameNumeric, "[0-9]+");
            FileToken.Priority = TerminalPriority.FileToken;;

            const string quotedFileSheetRegex = @"'\[\d+\]" + quotedSheetName + "'!";
            
            var QuotedFileSheetToken = new RegexBasedTerminal(GrammarNames.TokenFileSheetQuoted, quotedFileSheetRegex);
            QuotedFileSheetToken.Priority = TerminalPriority.QuotedFileToken;

            var ReservedNameToken = new RegexBasedTerminal(GrammarNames.TokenReservedName, @"_xlnm\.[a-zA-Z_]+");
            ReservedNameToken.Priority = TerminalPriority.ReservedName;

            var DDEToken = new RegexBasedTerminal(GrammarNames.TokenDDE, @"'([^']|'')+'");

            #endregion

            #region Punctuation
            MarkPunctuation(exclamationMark);
            MarkPunctuation(OpenParen, CloseParen);
            MarkPunctuation(OpenSquareParen, CloseSquareParen);
            MarkPunctuation(OpenCurlyParen, CloseCurlyParen);
            #endregion
            #endregion

            #region 2-NonTerminals
            // Most nonterminals are first defined here, so they can be used anywhere in the rules
            // Otherwise you can only use nonterminals that have been defined previously

            var Argument = new NonTerminal(GrammarNames.Argument);
            var Arguments = new NonTerminal(GrammarNames.Arguments);
            var ArrayColumns = new NonTerminal(GrammarNames.ArrayColumns);
            var ArrayConstant = new NonTerminal(GrammarNames.ArrayConstant);
            var ArrayFormula = new NonTerminal(GrammarNames.ArrayFormula);
            var ArrayRows = new NonTerminal(GrammarNames.ArrayRows);
            var Bool = new NonTerminal(GrammarNames.Bool);
            var Cell = new NonTerminal(GrammarNames.Cell);
            var Constant = new NonTerminal(GrammarNames.Constant);
            var ConstantArray = new NonTerminal(GrammarNames.ConstantArray);
            var DynamicDataExchange = new NonTerminal(GrammarNames.DynamicDataExchange);
            var EmptyArgument = new NonTerminal(GrammarNames.EmptyArgument);
            var Error = new NonTerminal(GrammarNames.Error);
            var File = new NonTerminal(GrammarNames.File);
            var Formula = new NonTerminal(GrammarNames.Formula);
            var FormulaWithEq = new NonTerminal(GrammarNames.FormulaWithEq);
            var FunctionCall = new NonTerminal(GrammarNames.FunctionCall);
            var FunctionName = new NonTerminal(GrammarNames.FunctionName);
            var HRange = new NonTerminal(GrammarNames.HorizontalRange);
            var InfixOp = new NonTerminal(GrammarNames.TransientInfixOp);
            var MultipleSheets = new NonTerminal(GrammarNames.MultipleSheets);
            var NamedRange = new NonTerminal(GrammarNames.NamedRange);
            var Number = new NonTerminal(GrammarNames.Number);
            var PostfixOp = new NonTerminal(GrammarNames.TransientPostfixOp);
            var Prefix = new NonTerminal(GrammarNames.Prefix);
            var PrefixOp = new NonTerminal(GrammarNames.TransientPrefixOp);
            var QuotedFileSheet = new NonTerminal(GrammarNames.QuotedFileSheet);
            var Reference = new NonTerminal(GrammarNames.Reference);
            //var ReferenceFunction = new NonTerminal(GrammarNames.ReferenceFunction);
            var ReferenceItem = new NonTerminal(GrammarNames.TransientReferenceItem);
            var ReferenceFunctionCall = new NonTerminal(GrammarNames.ReferenceFunctionCall);
            var RefError = new NonTerminal(GrammarNames.RefError);
            var RefFunctionName = new NonTerminal(GrammarNames.RefFunctionName);
            var ReservedName = new NonTerminal(GrammarNames.ReservedName);
            var Sheet = new NonTerminal(GrammarNames.Sheet);
            var Start = new NonTerminal(GrammarNames.TransientStart);
            var Text = new NonTerminal(GrammarNames.Text);
            var UDFName = new NonTerminal(GrammarNames.UDFName);
            var UDFunctionCall = new NonTerminal(GrammarNames.UDFunctionCall);
            var Union = new NonTerminal(GrammarNames.Union);
            var VRange = new NonTerminal(GrammarNames.VerticalRange);
            #endregion


            #region 3-Rules

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
                ;
            MarkTransient(ReferenceItem);

            UDFunctionCall.Rule = UDFName + Arguments + CloseParen;
            UDFName.Rule = UDFToken;

            VRange.Rule = VRangeToken;
            HRange.Rule = HRangeToken;
            
            //ConditionalRefFunctionName.Rule = ExcelConditionalRefFunctionToken;

            QuotedFileSheet.Rule = QuotedFileSheetToken;
            Sheet.Rule = SheetToken;
            MultipleSheets.Rule = MultipleSheetsToken;

            Cell.Rule = CellToken;

            File.Rule = OpenSquareParen + FileToken + CloseSquareParen;

            DynamicDataExchange.Rule = File + exclamationMark + DDEToken;

            NamedRange.Rule = NamedRangeToken | NamedRangeCombinationToken;

            Prefix.Rule =
                Sheet
                | File + Sheet
                | File + exclamationMark
                | QuotedFileSheet
                | MultipleSheets
                | File + MultipleSheets;

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
            
            public const int NamedRange = -800;
            public const int ReservedName = -700;

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
            public const int FileToken = 1200;
            public const int SheetToken = 1200;
            public const int QuotedFileToken = 1200;
        }
        #endregion

        #region Excel function list
        private static readonly IList<string> excelFunctionList = new List<String>
        {
            "ABS",
            "ACCRINT",
            "ACCRINTM",
            "ACOS",
            "ACOSH",
            "ADDRESS",
            "AMORDEGRC",
            "AMORLINC",
            "AND",
            "AREAS",
            "ASC",
            "ASIN",
            "ASINH",
            "ATAN",
            "ATAN2",
            "ATANH",
            "AVEDEV",
            "AVERAGE",
            "AVERAGEA",
            "AVERAGEIF",
            "AVERAGEIFS",
            "BAHTTEXT",
            "BESSELI",
            "BESSELJ",
            "BESSELK",
            "BESSELY",
            "BETADIST",
            "BETAINV",
            "BIN2DEC",
            "BIN2HEX",
            "BIN2OCT",
            "BINOMDIST",
            "CALL",
            "CEILING",
            "CELL",
            "CHAR",
            "CHIDIST",
            "CHIINV",
            "CHITEST",
            //"CHOOSE",
            "CLEAN",
            "CODE",
            "COLUMN",
            "COLUMNS",
            "COMBIN",
            "COMPLEX",
            "CONCATENATE",
            "CONFIDENCE",
            "CONVERT",
            "CORREL",
            "COS",
            "COSH",
            "COUNT",
            "COUNTA",
            "COUNTBLANK",
            "COUNTIF",
            "COUNTIFS",
            "COUPDAYBS",
            "COUPDAYS",
            "COUPDAYSNC",
            "COUPNCD",
            "COUPNUM",
            "COUPPCD",
            "COVAR",
            "CRITBINOM",
            "CUBEKPIMEMBER",
            "CUBEMEMBER",
            "CUBEMEMBERPROPERTY",
            "CUBERANKEDMEMBER",
            "CUBESET",
            "CUBESETCOUNT",
            "CUBEVALUE",
            "CUMIPMT",
            "CUMPRINC",
            "DATE",
            "DATEVALUE",
            "DAVERAGE",
            "DAY",
            "DAYS360",
            "DB",
            "DCOUNT",
            "DCOUNTA",
            "DDB",
            "DEC2BIN",
            "DEC2HEX",
            "DEC2OCT",
            "DEGREES",
            "DELTA",
            "DEVSQ",
            "DGET",
            "DISC",
            "DMAX",
            "DMIN",
            "DOLLAR",
            "DOLLARDE",
            "DOLLARFR",
            "DPRODUCT",
            "DSTDEV",
            "DSTDEVP",
            "DSUM",
            "DURATION",
            "DVAR",
            "DVARP",
            "EDATEEFFECT",
            "EOMONTH",
            "ERF",
            "ERFC",
            "ERROR.TYPE",
            "EUROCONVERT",
            "EVEN",
            "EXACT",
            "EXP",
            "EXPONDIST",
            "FACT",
            "FACTDOUBLE",
            "FALSE",
            "FDIST",
            "FIND",
            "FINV",
            "FISHER",
            "FISHERINV",
            "FIXED",
            "FLOOR",
            "FORECAST",
            "FREQUENCY",
            "FTEST",
            "FV",
            "FVSCHEDULE",
            "GAMMADIST",
            "GAMMAINV",
            "GAMMALN",
            "GCD",
            "GEOMEAN",
            "GESTEP",
            "GETPIVOTDATA",
            "GROWTH",
            "HARMEAN",
            "HEX2BIN",
            "HEX2DEC",
            "HEX2OCT",
            "HLOOKUP",
            "HOUR",
            "HYPERLINK",
            "HYPGEOMDIST",
            //"IF",
            "ISBLANK",
            "IFERROR",
            "IMABS",
            "IMAGINARY",
            "IMARGUMENT",
            "IMCONJUGATE",
            "IMCOS",
            "IMDIV",
            "IMEXP",
            "IMLN",
            "IMLOG10",
            "IMLOG2",
            "IMPOWER",
            "IMPRODUCT",
            "IMREAL",
            "IMSIN",
            "IMSQRT",
            "IMSUB",
            "IMSUM",
            "INFO",
            "INT",
            "INTERCEPT",
            "INTRATE",
            "IPMT",
            "IRR",
            "IS",
            "ISB",
            "ISERROR",
            "ISNA",
            "ISNUMBER",
            "ISPMT",
            "JIS",
            "KURT",
            "LARGE",
            "LCM",
            "LEFT",
            "LEFTB",
            "LEN",
            "LENB",
            "LINEST",
            "LN",
            "LOG",
            "LOG10",
            "LOGEST",
            "LOGINV",
            "LOGNORMDIST",
            "LOOKUP",
            "LOWER",
            "MATCH",
            "MAX",
            "MAXA",
            "MDETERM",
            "MDURATION",
            "MEDIAN",
            "MID",
            "MIDB",
            "MIN",
            "MINA",
            "MINUTE",
            "MINVERSE",
            "MIRR",
            "MMULT",
            "MOD",
            "MODE",
            "MONTH",
            "MROUND",
            "MULTINOMIAL",
            "N",
            "NA",
            "NEGBINOMDIST",
            "NETWORKDAYS",
            "NOMINAL",
            "NORMDIST",
            "NORMINV",
            "NORMSDIST",
            "NORMSINV",
            "NOT",
            "NOW",
            "NPER",
            "NPV",
            "OCT2BIN",
            "OCT2DEC",
            "OCT2HEX",
            "ODD",
            "ODDFPRICE",
            "ODDFYIELD",
            "ODDLPRICE",
            "ODDLYIELD",
            "OR",
            "PEARSON",
            "PERCENTILE",
            "PERCENTRANK",
            "PERMUT",
            "PHONETIC",
            "PI",
            "PMT",
            "POISSON",
            "POWER",
            "PPMT",
            "PRICE",
            "PRICEDISC",
            "PRICEMAT",
            "PROB",
            "PRODUCT",
            "PROPER",
            "PV",
            "QUARTILE",
            "QUOTIENT",
            "RADIANS",
            "RAND",
            "RANDBETWEEN",
            "RANK",
            "RATE",
            "RECEIVED",
            "REGISTER.ID",
            "REPLACE",
            "REPLACEB",
            "REPT",
            "RIGHT",
            "RIGHTB",
            "ROMAN",
            "ROUND",
            "ROUNDDOWN",
            "ROUNDUP",
            "ROW",
            "ROWS",
            "RSQ",
            "RTD",
            "SEARCH",
            "SEARCHB",
            "SECOND",
            "SERIESSUM",
            "SIGN",
            "SIN",
            "SINH",
            "SKEW",
            "SLN",
            "SLOPE",
            "SMALL",
            "SQL.REQUEST",
            "SQRT",
            "SQRTPI",
            "STANDARDIZE",
            "STDEV",
            "STDEVA",
            "STDEVP",
            "STDEVPA",
            "STEYX",
            "SUBSTITUTE",
            "SUBTOTAL",
            "SUM",
            "SUMIF",
            "SUMIFS",
            "SUMPRODUCT",
            "SUMSQ",
            "SUMX2MY2",
            "SUMX2PY2",
            "SUMXMY2",
            "SYD",
            "T",
            "TAN",
            "TANH",
            "TBILLEQ",
            "TBILLPRICE",
            "TBILLYIELD",
            "TDIST",
            "TEXT",
            "TIME",
            "TIMEVALUE",
            "TINV",
            "TODAY",
            "TRANSPOSE",
            "TREND",
            "TRIM",
            "TRIMMEAN",
            "TRUE",
            "TRUNC",
            "TTEST",
            "TYPE",
            "UPPER",
            "VALUE",
            "VAR",
            "VARA",
            "VARP",
            "VARPA",
            "VDB",
            "VLOOKUP",
            "WEEKDAY",
            "WEEKNUM",
            "WEIBULL",
            "WORKDAY",
            "XIRR",
            "XNPV",
            "YEAR",
            "YEARFRAC",
            "YIELD",
            "YIELDDISC",
            "YIELDMAT",
            "ZTEST"
        };
        #endregion
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
        public const string MultipleSheets = "MultipleSheets";
        public const string NamedRange = "NamedRange";
        public const string Number = "Number";
        public const string Prefix = "Prefix";
        public const string QuotedFileSheet = "QuotedFileSheet";
        public const string Range = "Range";
        public const string Reference = "Reference";
        //public const string ReferenceFunction = "ReferenceFunction";
        public const string ReferenceFunctionCall = "ReferenceFunctionCall";
        public const string RefError = "RefError";
        public const string RefFunctionName = "RefFunctionName";
        public const string ReservedName = "ReservedName";
        public const string Sheet = "Sheet";
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
        public const string TokenDDE = "DDEToken";
        public const string TokenEmptyArgument = "EmptyArgumentToken";
        public const string TokenError = "ErrorToken";
        public const string TokenExcelRefFunction = "ExcelRefFunctionToken";
        public const string TokenExcelConditionalRefFunction = "ExcelConditionalRefFunctionToken";
        public const string TokenFileNameNumeric = "FileNameNumericToken";
        public const string TokenFileSheetQuoted = "FileSheetQuotedToken";
        public const string TokenHRange = "HRangeToken";
        public const string TokenIntersect = "INTERSECT";
        public const string TokenMultipleSheets = "MultipleSheetsToken";
        public const string TokenNamedRange = "NamedRangeToken";
        public const string TokenNamedRangeCombination = "NamedRangeCombinationToken";
        public const string TokenNumber = "NumberToken";
        public const string TokenRefError = "RefErrorToken";
        public const string TokenReservedName = "ReservedNameToken";
        public const string TokenSheet = "SheetNameToken";
        public const string TokenText = "TextToken";
        public const string TokenUDF = "UDFToken";
        public const string TokenUnionOperator = ",";
        public const string TokenVRange = "VRangeToken";

        #endregion

    }
    #endregion
}
