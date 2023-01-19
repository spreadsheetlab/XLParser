using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Text.RegularExpressions;
using Irony.Parsing;

namespace XLParser.Tests
{
    [TestClass]
    public class ParserTests
    {
        [TestMethod]
        // Ensure that the grammar has no conflicts
        public void NoGrammarConflicts()
        {
            /* A conflict indicates an ambiguity in grammar or an error in the rules. In the
             * case of an ambiguity, Irony will choose an action, which may not be the right
             * one. Rewrite the grammar rules until there are no more conflicts, and the
             * parsing is correct. As a last resort, you can manually specify a shift with a
             * grammar hint such as PreferShiftHere() or ReduceHere(). However, this should
             * only be done if such a hint is applicable in all cases; otherwise, the grammar
             * must be written differently.
             */

            /* An example:
             * (A1) can be parsed both as an union and a bracketed reference. This is an
             * ambiguity, and thus must be solved by precedence or a grammar hint.
             */

            /* An example:
             * Functioncall.Rule =
             *   Prefix + Formula           // Prefix unop
             * | Formula + infix + Formula  // Binop
             * | Formula + Formula          // Intersection
             *
             * With this 1+1 can be parsed as both 1-1 and 1 (-1)
             * This is obviously erroneous as there is only one correct interpretation, so the
             * rules had to be rewritten.
             */

            var parser = new Parser(new ExcelFormulaGrammar());
            Assert.IsTrue(parser.Language.Errors.Count == 0, "Grammar has {0} error(s) or conflict(s)", parser.Language.Errors.Count);
        }

        private static ParseTreeNode Parse(string input)
        {
            return ExcelFormulaParser.Parse(input);
        }

        internal static void Test(string input, Predicate<ParseTreeNode> condition = null)
        {
            ParseTreeNode p = Parse(input);
            if (condition != null)
            {
                Assert.IsTrue(condition.Invoke(p.SkipToRelevant(true)), "condition failed for input '{0}'", input);
            }
            // Also do a print test for every parse
            PrintTests.test(formula: input, parsed: p);
        }

        private static void Test(IEnumerable<string> inputs, Predicate<ParseTreeNode> condition = null)
        {
            foreach (var input in inputs)
            {
                Test(input, condition);
            }
        }

        private static void Test(params string[] inputs)
        {
            foreach (var input in inputs)
            {
                Test(input);
            }
        }

        [TestMethod]
        public void BinaryOperators()
        {
            Test(new [] {"A1*5", "1+1", "1-1", "1/1", "1^1", "1&1"},
                n=> n.IsBinaryOperation() && n.IsOperation() && n.IsBinaryNonReferenceOperation());
            Test(new[] { "A1:A1", "A1 A1" },
                n => n.IsBinaryOperation() && n.IsOperation() && n.IsBinaryReferenceOperation());
        }

        [TestMethod]
        public void UnaryOperators()
        {
            Test(new [] {"+A5", "-1", "1%"},
            n=> n.IsUnaryOperation() && n.IsOperation());
            Test("-1", node => node.GetFunction() == "-");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void NotAFunction()
        {
            Test("1", node=>node.GetFunction()=="No function here");
        }

        [TestMethod]
        public void ExcelFunction()
        {
            Test("DAYS360(1)", node => node.IsFunction() && node.GetFunction() == "DAYS360" && node.IsBuiltinFunction());
            Test("SUM(1)", node => node.IsFunction() && node.GetFunction() == "SUM" && node.IsBuiltinFunction());
            Test("INDEX(1)", node => node.IsFunction() && node.GetFunction() == "INDEX" && node.IsBuiltinFunction());
            Test("IF(1)", node => node.IsFunction() && node.GetFunction() == "IF" && node.MatchFunction("IF") && node.IsBuiltinFunction());
            Test("MYUSERDEFINEDFUNCTION()", node => node.IsFunction() && node.GetFunction() == "MYUSERDEFINEDFUNCTION");
        }

        [TestMethod]
        public void UserDefinedFunction()
        {
            Test("AABS(1)", node => node.IsFunction() && node.GetFunction() == "AABS");
        }

        [TestMethod]
        public void ExternalUserDefinedFunction()
        {
            Test("[1]!myFunction()", node => node.IsFunction() && node.GetFunction() == "[1]!MYFUNCTION");
        }

        [TestMethod]
        public void Bool()
        {
            Test("True");
            Test("false");
        }



        [TestMethod]
        public void Range()
        {
            Test("SUM(A1:B5)");
        }

       [TestMethod]
        public void RangeWithIntersections()
        {
            Test("SUM((Total_Cost Jan):(Total_Cost Apr.))");
        }

        [TestMethod]
        public void RangeWithParentheses()
        {
            Test("SUM(_YR1:(YP))");
        }

        [TestMethod]
        public void NamedRange()
        {
            Test("SUM(testrange)");
        }

        [TestMethod]
        public void NamedRangeWithDigits()
        {
            Test("SUM(PC3eur)");
        }

        [TestMethod]
        public void NamedRangeSingleChar()
        {
            Test("MMULT(A,invA)");
        }

        [TestMethod]
        public void NamedRangeWithUnderscore()
        {
            Test("IF(SUM($E15:E15)<$C15*1000,MIN($C15*1000-SUM($C15:C15),$C15*1000*VLOOKUP(F$7,[0]!_4_15_YR_MACRS,3)),0)");
        }

        [TestMethod]
        public void NamedRangeOthers()
        {
            Test(new []
            {
                // See https://github.com/spreadsheetlab/XLParser/issues/40
                "XDO_?Amount?",
                "_ABC",
                @"\ABC",
                "äBC",
                "ABC.ABC",
                "ABC?ABC",
                "Abcäbc",
                @"ABC\ABC",
                // See https://github.com/spreadsheetlab/XLParser/issues/106
                "ct_per_€",
                "ctM3_naar_€MWh€"
            }, node => node.SkipToRelevant(true).Type() == GrammarNames.NamedRange);
        }

        [TestMethod]
        public void VRange()
        {
            Test("SUM(A:B)");
        }

        [TestMethod]
        public void NamedCell()
        {
            Test("+openstellingBSO48.");
        }

        [TestMethod]
        public void CellReference()
        {
            var cells = new[] {"A1", "Z9", "AB10", "ABC999", "BSO48"};
            foreach (var cell in cells)
            {
                Test(cell, tree => tree.AllNodes(GrammarNames.Cell).Select(ExcelFormulaParser.Print).FirstOrDefault() == cell);
            }
        }

        [TestMethod]
        public void MaxRowAddress()
        {
            Test("A1048576", node => node.SkipToRelevant(true).Type() == GrammarNames.Cell);
        }

        [TestMethod]
        public void InvalidRowAddress()
        {
            Test("A1048577", node => node.SkipToRelevant(true).Type() == GrammarNames.NamedRange);
        }

        [TestMethod]
        public void MaxColumnAddress()
        {
            Test("XFD1", node => node.SkipToRelevant(true).Type() == GrammarNames.Cell);
        }

        [TestMethod]
        public void InvalidColumnAddress()
        {
            Test("XFE1", node => node.SkipToRelevant(true).Type() == GrammarNames.NamedRange);
        }

        [TestMethod]
        public void TestErrorCodeNull()
        {
            Test("#NULL!");
        }

        [TestMethod]
        public void TestErrorCodeNullFormula()
        {
            Test("SUM(A1 A2)");
        }

        [TestMethod]
        public void TestErrorCodeRef()
        {
            Test("'hulp voor grafiek'!#REF!");
        }

        [TestMethod]
        public void TestErrorCodeRange()
        {
            Test("SUM(#REF!:#REF!)");
        }


        [TestMethod]
        public void TestErrorCodeName()
        {
            Test("#NAME?");
        }


        [TestMethod]
        public void EmptyArgumentAllowed()
        {
            Test("IF(AZ109=0,,AZ109*J109/12)");
            Test("EXP(,)");
            Test("EXP(,2,)");
        }

        [TestMethod]
        public void NoArgumentsAllowed()
        {
            Test("YEAR(TODAY())");
        }

        [TestMethod]
        public void TopLevelParentheses()
        {
            Test("('TOR PP FE-TR'!U87/('TOR PP FE-TR'!S87-'TOR PP FE-TR'!S68))");
        }


        [TestMethod]
        public void TwoEmptyArgumentsAllowed()
        {
            Test("IF(AZ109=0,,,AZ109*J109/12)");
        }

        [TestMethod]
        public void Union()
        {
            Test("LARGE((F38,C38),1)", node => node.ChildNodes[1].ChildNodes[0].SkipToRelevant(true).GetFunction() == ",");
            Test("LARGE((2:2,C38,$A$1:A6),1)", node => node.ChildNodes[1].ChildNodes[0].SkipToRelevant(true).GetFunction() == ",");
        }

        [TestMethod]
        public void BoolParses()
        {
            Test("IF(TRUE,A1,B6)");
        }

        [TestMethod]
        public void LongSheetRefence()
        {
            Test("Sheet1!A6");
        }

        [TestMethod]
        public void ShortSheetRefence()
        {
            Test("S1!A6");
        }

        [TestMethod]
        public void MultipleSheetRefence()
        {
            Test("Sheet1:Sheet2!A6");
            Test("'Sheet1:Sheet2'!A6");
            Test("SUM('[74]Miami P&L:Venezuela P&L'!G10)");
            //test(@"=SUM('D:\TV_LATAM_NET_ACC\Financial Statement Comments\Ad Sales\FY''15\Period 3[3-Ad Sales Financial Statements FY15 June-v2.xlsx]Miami P&L:Venezuela P&L'!G10)");
        }

        [TestMethod]
        public void MultipleSheetVSRange()
        {
            Test("AAA:XXX!A:B");
        }

        [TestMethod]
        public void DoublePrefixedRange()
        {
            Test("Sheet1!A1:'Sheet1'!A2");
        }

        [TestMethod]
        public void LongCellReference()
        {
            Test("Sheet2!A123456");
        }

        [TestMethod]
        public void Dollar()
        {
            Test("$B$6+ B$7+ $V9");
        }

        [TestMethod]
        public void DollarRange()
        {
            Test("SUM($B$6:F9) + SUM(B$7:$V9)");
        }

        [TestMethod]
        public void Comparison()
        {
            Test("IF(A1=3,A6,B9)");
        }

        [TestMethod]
        public void AdditionalBrackets()
        {
            Test("(SUM(B6:B8))");
        }

        [TestMethod]
        public void VLOOKUP()
        {
            Test("VLOOKUP(A1,A6:B25,6,TRUE)");
        }

        [TestMethod]
        public void Nested_Formula()
        {
            Test("ROUND(DAYS360(C7,(E132+30),TRUE)/360,2)");
        }

        [TestMethod]
        public void Space_in_Sheetname()
        {
            Test("D19+'Required Funds'!B16");
        }

        [TestMethod]
        public void NonAlphaSheetname()
        {
            Test("'Welcome!+-&'!B16");
        }

        [TestMethod]
        public void Text_with_spaces()
        {
            Test("IF(C36>0,\"Hello World\",\"Or not\")");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException), "If string contains quote")]
        public void Text_with_Quotes()
        {
            // Incorrect " within string, should be rejected.
            Test("IF(C36>0,\"Hello\" World\",\"Or not\")");
        }

        [TestMethod]
        public void Text_with_escaped_quotes()
        {
            Test("\"a\"\"ap\"");
        }

        [TestMethod]
        public void TextWithEscapesAndDoubleQuotes()
        {
            Test(@"IF(RIGHT(G8,1)=""\"","""",""PLEASE END THE PATH STRING WITH A '\' SYMBOL."")",
                // Check if any of the strings is "\"
                tree => tree.AllNodes(GrammarNames.Text).Any(text => text.Print() == @"""\"""));
        }

        [TestMethod]
        public void Text_with_linebreak()
        {
            Test("\"line1" + Environment.NewLine + "line two\"");
        }

        [TestMethod]
        public void Text_with_dot()
        {
            Test("IF(C36>0,\". says Hello Nurse\")");
        }

        [TestMethod]
        public void Text_with_comma()
        {
            Test("IF(C36>0,\"Hello Nurse, says .\")");
        }

        [TestMethod]
        public void Testcase_Arie()
        {
            Test("IF(C5<>1,noshow, basis + 9 * (AC5 + bonuspunten) / ($AC$2 - ignored))");
        }

        [TestMethod]
        public void Testcase_120()
        {
            Test("VLOOKUP($B6,Invoeren!$B:$I,2,FALSE)");
        }

        [TestMethod]
        public void Testcase_130_Inequality()
        {
            Test("IF(D2>=voldoende, vresult, oresult)");
        }

        [TestMethod]
        public void Testcase_130_Compare_in_String()
        {
            Test("COUNTIF(AD5:AD181, \">= 5.75\")");
        }

        [TestMethod]
        public void Testcase_439_VRange()
        {
            Test("+SUMIF('# cl gespec zorg'!A:A,A:A,'# cl gespec zorg'!B:B)");
        }

        [TestMethod]
        public void Testcase_130_Empty_String()
        {
            Test("\"\"");
        }

        [TestMethod]
        public void Testcase_437_NamedRange()
        {
            Test("ROUND((SUM(PC3eur))/1000,)");
        }

        [TestMethod]
        public void Testcase_79_EmptyArgument()
        {
            Test("IF($K$20<K21,SUM(L21:N21,),0)");
        }

        [TestMethod]
        public void Testcase_524()
        {
            Test("RIGHT([1]!SheetName(),(LEN([1]!SheetName())-20))");
        }

        [TestMethod]
        public void BracesInFormula()
        {
            Test("OR(MONTH(GQ$2) = {3,9})");
        }

        [TestMethod]
        public void BracesSemicolonInFormula()
        {
            Test("VLOOKUP($B192,'Sheet1'!$B:$CD,{35;36},FALSE)");
        }

        [TestMethod]
        public void xlmnInFormula()
        {
            Test("SUMIF('150000'!A:A,_xlnm.Print_Titles,'150000'!C:C)");
        }

        [TestMethod]
        public void CellReferenceFunctionInFormula()
        {
            Test("XIRR($Y$7:INDEX($D$3:$AS$7,5,MATCH(Assumptions!$C$18,Returns!$D$3:$AS$3)),$Y$3:INDEX($D$3:$AS$3,1,MATCH(Assumptions!$C$18,Returns!$D$3:$AS$3)))");
            Test("IF(O3=1,1,LINEST(OFFSET(O$9,ABS($F$2-$F$1)+1,0):OFFSET(O$9,$F$1,0),OFFSET($N$9,ABS($F$2-$F$1)+1,0):OFFSET($N9,$F$1,0)))");
            Test("AVERAGE(INDIRECT(\"D\"&$A$2):INDIRECT(\"D\"&$A$7))");
        }

        [TestMethod]
        public void xllInFormula()
        {
            Test("_xll.HEAT($B9,$C9)");
        }

        [TestMethod]
        public void BigNumberInFormula()
        {
            Test("30426000000/E7/1000");
        }

        [TestMethod]
        public void UDFWithNumericCharacters()
        {
            Test("ASTRIP2_m(E9,E10)");
        }

        [TestMethod]
        public void UDFLikeNamedRangeCombination()
        {
            Test("Prob1OptimalRiskyWeight(C7,C6,E7,G5)");
        }

        [TestMethod]
        public void UDFWithDot()
        {
            Test("Functions.BScall(C3,C4,C5,C6,C7,C8)");
        }

        [TestMethod]
        public void NamedRangeReference()
        {
            Test("SUM(TestRange)");
            Test("SUM(Sheet1!TestRange)");
        }

        [TestMethod]
        public void SheetWithUnderscore()
        {
            Test("aap_noot!B12");
        }

        [TestMethod]
        public void SheetWithPeriod()
        {
            Test("vrr2011_Omz.!M84");
        }

        [TestMethod]
        public void SheetAsString()
        {
            Test("'[20]Algemene info + \"Overview\"'!T95");
        }

        [TestMethod]
        public void SheetWithQuote()
        {
            Test("'Owner''s Engineer'!$A$2");
        }

        [TestMethod]
        public void ExternalSheetWithQuote()
        {
            Test("'[1]Stacey''s Reconciliation'!C3");
        }

        [TestMethod]
        public void DirectSheetReference()
        {
            Test("Sheet1!F7");
        }

        [TestMethod]
        public void SheetReference()
        {
            Test("SUM(Sheet1!X1)");
        }

        [TestMethod]
        public void FileReference()
        {
            Test("[2]Sheet1!X1");
        }

        [TestMethod]
        public void FileReferenceInRange()
        {
            Test("[2]Sheet1!X1:X10");
        }

        [TestMethod]
        public void QuotedFileReference()
        {
            Test("'[2]Sheet1'!X1");
        }


        [TestMethod]
        public void SheetReferenceRange()
        {
            Test("SUM(Sheet1!X1:X6)");
        }

        [TestMethod]
        public void MultipleSheetsReferenceCell()
        {
            Test("SUM(Sheet1:Sheet2!A1)");
        }

        [TestMethod]
        public void MultipleSheetsReferenceRange()
        {
            Test("SUM(Sheet1:Sheet2!A1:A3)");
        }

        [TestMethod]
        public void MultipleSheetsInFileReferenceCell()
        {
            Test("SUM([1]Sheet1:Sheet2!B15)");
        }

        [TestMethod]
        public void RangeWithPrefixedRightLimitReference ()
        {
            Test("SUM(Deals!F9:Deals!F16)");
        }

        [TestMethod]
        public void ReferencesInSingleColumn()
        {
            Test("Sheet1!A:A");
        }

        [TestMethod]
        public void ReferencesInLargeRange()
        {
            Test("'B-Com'!A1:Z1048576");
        }

        [TestMethod]
        public void ReferenceFunctionAsArgument()
        {
            Test("ROUND(INDEX(A:A,1,1:1),1)");
        }

        [TestMethod]
        public void RangeWithReferenceFunction()
        {
            Test("SUM(A1:INDEX(A:A,1,1:1))");
        }

        [TestMethod]
        public void UnionArgument()
        {
            Test("LARGE((F38,C38:C48),1)");
        }

        [TestMethod]
        public void DDE()
        {
            Test("[1]!'INDU Index,[PX_close_5d]'");
            // Test("=REUTER|IDN!'NGH2,PRIM ACT 1,1'"); TODO
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException), "If formula can't be parsed")]
        public void ParsingFails()
        {
            Test("]");
        }

        [TestMethod]
        public void EqualsIsFunction()
        {
            Test("A1=A2", node => node.IsFunction() && node.GetFunction() == "=");
        }

        [TestMethod]
        public void PercentIsFunction()
        {
            Test("1%", node => node.IsFunction() && node.GetFunction() == "%");
        }

        [TestMethod]
        public void FunctionsAsRefExpressions()
        {
            Test("IF(TRUE,A1,A2):B5", "INDEX():B5", "MyUDFunction:B5", "Sheet!MyUDFunction:B5");
        }

        [TestMethod]
        public void Bug()
        {
            Test("SUM(B5,2)");
        }

        [TestMethod]
        public void TestIsParentheses()
        {
            // Can't use test() for this one, since test() invokes skipFormula()
            Assert.IsTrue(ExcelFormulaParser.Parse("(1)").IsParentheses());
            Assert.IsTrue(ExcelFormulaParser.Parse("(A1)").ChildNodes[0].IsParentheses());
            // Make sure unions aren't recognized as parentheses
            var union = ExcelFormulaParser.Parse("(A1,A2)");
            Assert.IsFalse(union.IsParentheses());
            Assert.IsFalse(union.ChildNodes[0].IsParentheses());
        }

        [TestMethod]
        public void TestQuotedFileSheetWithPath()
        {
            Test(@"='C:\mypath\[myfile.xlsm]Sheet'!A1");
        }

        [TestMethod]
        public void ExternalWorkbook()
        {
            Test(@"='C:\Users\Test\Desktop\[Book1.xlsx]Sheet1'!$A$1");
        }

        [TestMethod]
        public void ExternalWorkbookNetworkPath()
        {
            // See [#107](https://github.com/spreadsheetlab/XLParser/issues/107)
            Test(@"='\\TEST-01\Folder\[Book1.xlsx]Sheet1'!$A$1");
        }

        [TestMethod]
        public void ExternalWorkbookUrlPathHttp()
        {
            // See [#108](https://github.com/spreadsheetlab/XLParser/issues/108)
            Test(@"='http://example.com/test/[Book1.xlsx]Sheet1'!$A$1");
        }

        [TestMethod]
        public void ExternalWorkbookUrlPathHttps()
        {
            // See [#114](https://github.com/spreadsheetlab/XLParser/issues/114)
            Test(@"='https://d.docs.live.net/3fade139bf25879f/Documents/[Tracer.xlsx]Sheet2'!$C$5+Sheet10!J44");
        }

        [TestMethod]
        public void ExternalWorkbookRelativePath()
        {
            // See [#109](https://github.com/spreadsheetlab/XLParser/issues/109)
            Test(@"='Test\Folder\[Book1.xlsx]Sheet1'!$A$1");
        }

        [TestMethod]
        public void ExternalWorkbookSingleCell()
        {
            Test(@"='C:\Users\Test\Desktop\[Book1.xlsx]Sheet1'!$A$1");
        }

        [TestMethod]
        public void ExternalWorkbookCellRange()
        {
            Test(@"='C:\Users\Test\Desktop\[Book1.xlsx]Sheet1'!$A$1:$A$10");
        }

        [TestMethod]
        public void ExternalWorkbookDefinedNameLocalScope()
        {
            Test(@"='C:\Users\Test\Desktop\[Book1.xlsx]Sheet1'!FirstItem");
        }

        [TestMethod]
        public void ExternalWorkbookDefinedNameGlobalScope()
        {
            // See [#101](https://github.com/spreadsheetlab/XLParser/issues/101)
            Test(@"='C:\Users\Test\Desktop\Book1.xlsx'!Items");
        }

        [TestMethod]
        public void MultipleExternalWorkbookSingleCell()
        {
            Test(@"=SUM('C:\Users\Test\Desktop\[Book1.xlsx]Sheet1'!$A$1,'C:\Users\Test\Desktop\[Book1.xlsx]Sheet1'!$A$2)");
        }

        [TestMethod]
        public void MultipleExternalWorkbookCellRange()
        {
            Test(@"=SUM('C:\Users\Test\Desktop\[Book1.xlsx]Sheet1'!$A$1:$A$10,'C:\Users\Test\Desktop\[Book1.xlsx]Sheet1'!$A$11:$A$20)");
        }

        [TestMethod]
        public void MultipleExternalWorkbookDefinedNameLocalScope()
        {
            Test(@"=SUM('C:\Users\Test\Desktop\[Book1.xlsx]Sheet1'!FirstItem,'C:\Users\Test\Desktop\[Book1.xlsx]Sheet1'!SecondItem)");
        }

        [TestMethod]
        public void MultipleExternalWorkbookDefinedNameGlobalScope()
        {
            Test(@"=SUM('C:\Users\Test\Desktop\Book1.xlsx'!Items,'C:\Users\Test\Desktop\Book1.xlsx'!Items2)");
        }

        [TestMethod]
        public void ExternalWorkbookWithoutPath()
        {
            Test("=[Book1.xlsx]Sheet!A1");
            Test("=[Book1.xlsx]!Salary");
        }

        [TestMethod]
        public void ExternalWorkbookWithoutPathAndExtension()
        {
            Test("=[Book1]Sheet!A1");
            Test("=[Book1]!Salary");
        }

        [TestMethod]
        public void TestFilePathWithSpace()
        {
            Test(@"='C:\EOL\Management Report\DATAMART\REGIONAL ANALYSIS\REPORTS\052301\[DEAL BREAKDOWN ANALYSIS 05-23-01.xls]PHYSICAL+FINANCIAL PIVOT '!D8",
                @"='C:\EOL\Management Report\DATAMART\REGIONAL ANALYSIS\REPORTS\052301\[DEAL BREAKDOWN ANALYSIS 05-23-01.xls]PHYSICAL+FINANCIAL PIVOT '!D8",
                @"='C:\EOL\Management Report\DATAMART\REGIONAL ANALYSIS\REPORTS\052301\[DEAL BREAKDOWN ANALYSIS 05-23-01.xls]PHYSICAL+FINANCIAL PIVOT '!D6",
                @"='C:\EOL\Management Report\DATAMART\REGIONAL ANALYSIS\REPORTS\052301\[DEAL BREAKDOWN ANALYSIS 05-23-01.xls]PHYSICAL+FINANCIAL PIVOT '!D8",
                @"='C:\EOL\Management Report\DATAMART\REGIONAL ANALYSIS\REPORTS\052301\[DEAL BREAKDOWN ANALYSIS 05-23-01.xls]PHYSICAL+FINANCIAL PIVOT '!D7",
                @"='C:\EOL\Management Report\DATAMART\REGIONAL ANALYSIS\REPORTS\052301\[DEAL BREAKDOWN ANALYSIS 05-23-01.xls]FINANCIAL PIVOT'!D8",
                @"='C:\EOL\Management Report\DATAMART\REGIONAL ANALYSIS\REPORTS\052301\[DEAL BREAKDOWN ANALYSIS 05-23-01.xls]FINANCIAL PIVOT'!D8",
                @"='C:\EOL\Management Report\DATAMART\REGIONAL ANALYSIS\REPORTS\052301\[DEAL BREAKDOWN ANALYSIS 05-23-01.xls]FINANCIAL PIVOT'!D7",
                @"='C:\EOL\Management Report\DATAMART\REGIONAL ANALYSIS\REPORTS\052301\[DEAL BREAKDOWN ANALYSIS 05-23-01.xls]PHYSICAL PIVOT'!D6",
                @"='C:\EOL\Management Report\DATAMART\REGIONAL ANALYSIS\REPORTS\052301\[DEAL BREAKDOWN ANALYSIS 05-23-01.xls]PHYSICAL PIVOT'!D8",
                @"='C:\EOL\Management Report\DATAMART\REGIONAL ANALYSIS\REPORTS\052301\[DEAL BREAKDOWN ANALYSIS 05-23-01.xls]PHYSICAL PIVOT'!D7"
                );
        }

        [TestMethod]
        public void TestSpaceAsSheetName()
        {
            // See [Issue 37](https://github.com/spreadsheetlab/XLParser/issues/37)
            var inputs = new[] {"VLOOKUP(' '!G3,' '!B4:F99,5)", "VLOOKUP('\t'!G3,'\t'!B4:F99,5)", "VLOOKUP('   '!G3,'   '!B4:F99,5)"};
            Test(inputs, tree => tree.AllNodes(GrammarNames.Prefix).All(node => Regex.IsMatch(node.GetPrefixInfo().Sheet, @"^\s+$")));
        }

        [TestMethod]
        public void SheetNamesWithSpacesCanBeExtractedCorrectly()
        {
            var strangeSheetNames = new[] {"\t", " ","   ", " A", " ''A", " A ", " ''A1'' "};
            foreach (var sheetName in strangeSheetNames)
            {
                // Simple reference to another sheet
                var sourceText = $"'{sheetName}'!A1";
                ParseTreeNode node = ExcelFormulaParser.Parse(sourceText);

                var actual = node.AllNodes(GrammarNames.Prefix).First().GetPrefixInfo().Sheet;

                Assert.AreEqual(sheetName, actual);
            }
        }

        [TestMethod]
        public void SheetNameIsReferenceError()
        {
            var formulas = new[] {"#REF!A1", "B1+#REF!A1"};
            Test(formulas);

            // See [#76](https://github.com/spreadsheetlab/XLParser/issues/76)
            foreach (var formula in formulas)
            {
                Assert.AreEqual(0, ExcelFormulaParser.ParseToTree(formula).Tokens.Count(TokenIsIntersect));
            }

            const string formulaWithOneIntersect = "#REF!A:A #REF!1:1";
            Test(formulaWithOneIntersect);
            Assert.AreEqual(1, ExcelFormulaParser.ParseToTree(formulaWithOneIntersect).Tokens.Count(TokenIsIntersect));

            bool TokenIsIntersect(Token token)
            {
                return string.Equals(token.Terminal.Name, "INTERSECT", StringComparison.OrdinalIgnoreCase);
            }
        }

        [TestMethod]
        public void TestNamedRangeCombination()
        {
            // See [Issue 46](https://github.com/spreadsheetlab/XLParser/issues/46)
            // This concerns names beginning with non-name words.
            var names = new[] {"A1ABC", "A1A1", "A2.PART_NUM", "A2?PART_NUM", "TRUEFOO", "FALSEFOO", "TRUEMODEL", "W1."};
            foreach (var name in names)
            {
                Test(name, tree => tree.AllNodes(GrammarNames.NamedRange).Select(ExcelFormulaParser.Print).Contains(name));
            }
        }

        [TestMethod]
        public void BackslashInEnclosedInBracketsToken()
        {
            // See [#84](https://github.com/spreadsheetlab/XLParser/issues/84)
            Test(@"'C:\MyAddins\MyAddin.xla'!MyFunc(""sampleTextArg"", 1)");
        }

        [TestMethod]
        public void ParseUdfNamesWithSpecialCharacters()
        {
            // See [#55](https://github.com/spreadsheetlab/XLParser/issues/55)
            Test("·()", "¡¢£¤¥¦§¨©«¬­®¯°±²³´¶·¸¹»¼½¾¿×÷()");

            Assert.ThrowsException<ArgumentException>(() => ExcelFormulaParser.ParseToTree("A↑()"));
        }

        [TestMethod]
        public void DoNotParseUdfNamesConsistingOnlyOfROrC()
        {
            // See [#56](https://github.com/spreadsheetlab/XLParser/issues/56)
            // UDF function names consisting of a single character "R" or "C" must be prefixed by the module name
            foreach (var disallowed in "RrCc")
            {
                Assert.ThrowsException<ArgumentException>(() => ExcelFormulaParser.ParseToTree($"{disallowed}()"));
            }

            Test("Module1.R()", "N()", "T()", "Ñ()");
        }

        [TestMethod]
        public void ImplicitIntersection()
        {
            Test("=@B10:B20");
            Test("=@namedRange");
            Test("=SQRT(A1:A4)+ SQRT(@A1:A4)");
            Test("=SQRT(A1:A4)+ @SQRT(A1:A4)");
            Test("=SQRT(A1:A4)+ SQRT(@namedRange)");
            Test("=IF(J131=1,@Index($F125:$F129,J132)/$F130,0)");            
        }

        [TestMethod]
        public void SpillError()
        {
            Test("#SPILL!");
        }

        [TestMethod]
        public void SpillRangeReference()
        {
            Test("=B10#");
            Test("=(B10)#");
            Test("=OFFSET(\"A1\",10,1)#");
        }

        [TestMethod]
        public void UnionOperator()
        {
            // See [#98](https://github.com/spreadsheetlab/XLParser/issues/98)
            // See [#124](https://github.com/spreadsheetlab/XLParser/issues/124)
            Test("=(A1:A3,C1:C3)");
            Test("=A1:A3,C1:C3");
            Test("=A1:A5,C1:C5,E1:E5");
            Test("=Sheet1!$A$1,Sheet1!$B$2");
        }


        [TestMethod]
        public void SmbPaths()
        {
            // See [#136](https://github.com/spreadsheetlab/XLParser/issues/136)
            Test("='\\\\TEST-01\\Folder\\[Book1.xlsx]Sheet1'!$A$1", tree =>
                tree.AllNodes().Count(x => x.Is(GrammarNames.Reference)) == 1);
            Test("='C:\\TEST-01\\Folder\\[Book1.xlsx]Sheet1'!$A$1+'C:\\TEST-01\\Folder\\[Book1.xlsx]Sheet1'!$A$2",
                tree =>
                    tree.AllNodes().Count(x => x.Is(GrammarNames.Reference)) == 2);
            Test("='\\\\TEST-01\\Folder\\[Book1.xlsx]Sheet1'!$A$1+'\\\\TEST-01\\Folder\\[Book1.xlsx]Sheet1'!$A$2",
                tree =>
                    tree.AllNodes().Count(x => x.Is(GrammarNames.Reference)) == 2);
            Test("='\\\\TEST-01\\Folder\\[Book1.xlsx]Sheet1'!$A$1+'\\\\TEST-01\\[Folder]\\[Book1.xlsx]Sheet1'!$A$2",
                tree =>
                    tree.AllNodes().Count(x => x.Is(GrammarNames.Reference)) == 2);
        }
    }
}
