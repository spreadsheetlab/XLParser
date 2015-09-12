using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections;
using System.IO;
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
            /* Conflict either indicate an ambiguity in the grammar or an error in the rules
             * In case of an ambiguity Irony will choose an action itself, which is not always the correct one
             * Rewrite the grammar rules untill there are no conflicts, and the parses are corrent 
             * As a very last resort, you can manually specify shift with a grammar hint like PreferShiftHere() or ReduceHere()
             * However, this should only be done if one of the two is always correct, otherwise the grammar will need to be written differently
             */

            /* An example:
             * (A1) can be parsed both as an union and a bracketed reference
             * This is an ambiguity, and thus must be solved by precedence or a grammar hint
             */
            
            /* An example:
             * Functioncall.Rule =
             *   Prefix + Formula           // Prefix unop
             * | Formula + infix + Formula  // Binop
             * | Formula + Formula          // Intersection
             * 
             * With this 1+1 can be parsed as both 1-1 and 1 (-1)
             * This is obviously erroneous as there is only 1 correct interpertation, so the rules had to be rewritten.
             */

            var parser = new Parser(new ExcelFormulaGrammar());
            Assert.IsTrue(parser.Language.Errors.Count == 0, "Grammar has {0} error(s) or conflict(s)", parser.Language.Errors.Count);
        }

        private static ParseTreeNode parse(string input)
        {
            return ExcelFormulaParser.Parse(input);
        }

        internal static void test(string input, Predicate<ParseTreeNode> condition = null)
        {
            var p = parse(input);
            if (condition != null)
            {
                Assert.IsTrue(condition.Invoke(p.SkipToRelevant()), "condition failed for input '{0}'", input);
            }
            // Also do a print test for every parse
            PrintTests.test(formula: input, parsed: p);
        }

        private void test(IEnumerable<string> inputs, Predicate<ParseTreeNode> condition = null)
        {
            foreach (string input in inputs)
               test(input, condition);
        }

        private void test(params string[] inputs) {
            foreach (string input in inputs) test(input);
        }

        
        [TestMethod]
        public void BinaryOperators()
        {
            test(new [] {"A1*5", "1+1", "1-1", "1/1", "1^1", "1&1"},
                n=> n.IsBinaryOperation() && n.IsOperation() && n.IsBinaryNonReferenceOperation());
            test(new[] { "A1:A1", "A1 A1" },
                n => n.IsBinaryOperation() && n.IsOperation() && n.IsBinaryReferenceOperation());
        }

        [TestMethod] 
        public void UnaryOperators()
        {
            test(new [] {"+A5", "-1", "1%"},
            n=> n.IsUnaryOperation() && n.IsOperation());
            test("-1", node => node.GetFunction() == "-");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void NotAFunction()
        {
            test("1", node=>node.GetFunction()=="No function here");
        }

        [TestMethod]
        public void ExcelFunction()
        {
            test("DAYS360(1)", node => node.IsFunction() && node.GetFunction() == "DAYS360" && node.IsBuiltinFunction());
            test("SUM(1)", node => node.IsFunction() && node.GetFunction() == "SUM" && node.IsBuiltinFunction());
            test("INDEX(1)", node => node.IsFunction() && node.GetFunction() == "INDEX" && node.IsBuiltinFunction());
            test("IF(1)", node => node.IsFunction() && node.GetFunction() == "IF" && node.MatchFunction("IF") && node.IsBuiltinFunction());
            test("MYUSERDEFINEDFUNCTION()", node => node.IsFunction() && node.GetFunction() == "MYUSERDEFINEDFUNCTION");
        }

        [TestMethod]
        public void UserDefinedFunction()
        {
            test("AABS(1)", node => node.IsFunction() && node.GetFunction() == "AABS");
        }

        [TestMethod]
        public void ExternalUserDefinedFunction()
        {
            test("[1]!myFunction()", node => node.IsFunction() && node.GetFunction() == "[1]!MYFUNCTION");
        }

        [TestMethod]
        public void Bool()
        {
            test("True");
            test("false");
        }



        [TestMethod]
        public void Range()
        {
            test("SUM(A1:B5)");
        }

       [TestMethod]
        public void RangeWithIntersections()
        {
            test("SUM((Total_Cost Jan):(Total_Cost Apr.))");
        }

        [TestMethod]
        public void RangeWithParentheses()
        {
            test("SUM(_YR1:(YP))");
        }

        [TestMethod]
        public void NamedRange()
        {
            test("SUM(testrange)");
        }

        [TestMethod]
        public void NamedRangeWithDigits()
        {
            test("SUM(PC3eur)");
        }

        [TestMethod]
        public void NamedRangeSingleChar()
        {
            test("MMULT(A,invA)");
        }

        [TestMethod]
        public void NamedRangeWithUnderscore()
        {
            test("IF(SUM($E15:E15)<$C15*1000,MIN($C15*1000-SUM($C15:C15),$C15*1000*VLOOKUP(F$7,[0]!_4_15_YR_MACRS,3)),0)");
        }

        [TestMethod]
        public void VRange()
        {
            test("SUM(A:B)");   
        }

        [TestMethod]
        public void NamedCell()
        {
            test("+openstellingBSO48.");
        }

        [TestMethod]
        public void CellReference()
        {
            test("+BSO48");
        }



        [TestMethod]
        public void TestErrorCodeNull()
        {
            test("#NULL!");   
        }

        [TestMethod]
        public void TestErrorCodeNullFormula()
        {
            test("SUM(A1 A2)");   
        }

        [TestMethod]
        public void TestErrorCodeRef()
        {
            test("'hulp voor grafiek'!#REF!");   
        }

        [TestMethod]
        public void TestErrorCodeRange()
        {
            test("SUM(#REF!:#REF!)");   
        }


        [TestMethod]
        public void TestErrorCodeName()
        {
            test("#NAME?");   
        }


        [TestMethod]
        public void EmptyArgumentAllowed()
        {
            test("IF(AZ109=0,,AZ109*J109/12)");
            test("EXP(,)"); 
            test("EXP(,2,)");  
        }

        [TestMethod]
        public void NoArgumentsAllowed()
        {
            
            test("YEAR(TODAY())");
            
        }

        [TestMethod]
        public void TopLevelParentheses()
        {
            
            test("('TOR PP FE-TR'!U87/('TOR PP FE-TR'!S87-'TOR PP FE-TR'!S68))");
            
        }      


        [TestMethod]
        public void TwoEmptyArgumentsAllowed()
        {
            test("IF(AZ109=0,,,AZ109*J109/12)");
        }

        [TestMethod]
        public void Union()
        {
            test("LARGE((F38,C38),1)", node => node.ChildNodes[1].ChildNodes[0].SkipToRelevant().GetFunction() == ",");
            test("LARGE((2:2,C38,$A$1:A6),1)", node => node.ChildNodes[1].ChildNodes[0].SkipToRelevant().GetFunction() == ",");   
        }

        [TestMethod]
        public void BoolParses()
        {
            test("IF(TRUE,A1,B6)");
        }

        [TestMethod]
        public void LongSheetRefence()
        {
            test("Sheet1!A6");   
        }

        [TestMethod]
        public void ShortSheetRefence()
        {
            test("S1!A6");   
        }

        [TestMethod]
        public void MultipleSheetRefence()
        {
            test("Sheet1:Sheet2!A6");
        }

        [TestMethod]
        public void MultipleSheetVSRange()
        {
            test("AAA:XXX!A:B");   
        }

        [TestMethod]
        public void DoublePrefixedRange()
        {
            test("Sheet1!A1:'Sheet1'!A2");   
        }

        [TestMethod]
        public void LongCellReference()
        {
            test("Sheet2!A1234567");   
        }

        [TestMethod]
        public void Dollar()
        {   
            test("$B$6+ B$7+ $V9");   
        }

        [TestMethod]
        public void DollarRange()
        {
            test("SUM($B$6:F9) + SUM(B$7:$V9)");   
        }

        [TestMethod]
        public void Comparison()
        {
            test("IF(A1=3,A6,B9)");
        }

        [TestMethod]
        public void AdditionalBrackets()
        {
            test("(SUM(B6:B8))");   
        }

        [TestMethod]
        public void VLOOKUP()
        {
            test("VLOOKUP(A1,A6:B25,6,TRUE)");   
        }

        [TestMethod]
        public void Nested_Formula()
        {
            test("ROUND(DAYS360(C7,(E132+30),TRUE)/360,2)");   
        }

        [TestMethod]
        public void Space_in_Sheetname()
        {
            test("D19+'Required Funds'!B16");   
        }

        [TestMethod]
        public void NonAlphaSheetname()
        {
            test("'Welcome!+-&'!B16");   
        }

        [TestMethod]
        public void Text_with_spaces()
        {
            test("IF(C36>0,\"Hello World\",\"Or not\")");   
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException), "If string contains quote")]
        public void Text_with_Quotes()
        {
            // Incorrect " within string, should be rejected.
            test("IF(C36>0,\"Hello\" World\",\"Or not\")");
        }

        [TestMethod]
        public void Text_with_escaped_quotes()
        {
            test("\"a\"\"ap\"");   
        }

        [TestMethod]
        public void Text_with_linebreak()
        {
            test("\"line1" + Environment.NewLine + "line two\"");   
        }

        [TestMethod]
        public void Text_with_dot()
        {
            test("IF(C36>0,\". says Hello Nurse\")");   
        }

        [TestMethod]
        public void Text_with_comma()
        {
            test("IF(C36>0,\"Hello Nurse, says .\")");   
        }

        [TestMethod]
        public void Testcase_Arie()
        {
            test("IF(C5<>1,noshow, basis + 9 * (AC5 + bonuspunten) / ($AC$2 - ignored))");   
        }

        [TestMethod]
        public void Testcase_120()
        {
            test("VLOOKUP($B6,Invoeren!$B:$I,2,FALSE)");   
        }

        [TestMethod]
        public void Testcase_130_Inequality()
        {
            test("IF(D2>=voldoende, vresult, oresult)");
        }

        [TestMethod]
        public void Testcase_130_Compare_in_String()
        {   
            test("COUNTIF(AD5:AD181, \">= 5.75\")");
        }

        [TestMethod]
        public void Testcase_439_VRange()
        {
            test("+SUMIF('# cl gespec zorg'!A:A,A:A,'# cl gespec zorg'!B:B)");   
        }

        [TestMethod]
        public void Testcase_130_Empty_String()
        {
            test("\"\"");
        }

        [TestMethod]
        public void Testcase_437_NamedRange()
        {
            test("ROUND((SUM(PC3eur))/1000,)");   
        }

        [TestMethod]
        public void Testcase_79_EmptyArgument()
        {
            test("IF($K$20<K21,SUM(L21:N21,),0)");   
        }

        [TestMethod]
        public void Testcase_524()
        {   
            test("RIGHT([1]!SheetName(),(LEN([1]!SheetName())-20))");   
        }

        [TestMethod]
        public void BracesInFormula()
        {   
            test("OR(MONTH(GQ$2) = {3,9})");   
        }

        [TestMethod]
        public void BracesSemicolonInFormula()
        {
            test("VLOOKUP($B192,'Sheet1'!$B:$CD,{35;36},FALSE)");   
        }

        [TestMethod]
        public void xlmnInFormula()
        {
            test("SUMIF('150000'!A:A,_xlnm.Print_Titles,'150000'!C:C)");
        }

        [TestMethod]
        public void CellReferenceFunctionInFormula()
        {
            test("XIRR($Y$7:INDEX($D$3:$AS$7,5,MATCH(Assumptions!$C$18,Returns!$D$3:$AS$3)),$Y$3:INDEX($D$3:$AS$3,1,MATCH(Assumptions!$C$18,Returns!$D$3:$AS$3)))");   
            test("IF(O3=1,1,LINEST(OFFSET(O$9,ABS($F$2-$F$1)+1,0):OFFSET(O$9,$F$1,0),OFFSET($N$9,ABS($F$2-$F$1)+1,0):OFFSET($N9,$F$1,0)))");
            test("AVERAGE(INDIRECT(\"D\"&$A$2):INDIRECT(\"D\"&$A$7))");
        }

        [TestMethod]
        public void xllInFormula()
        {
            test("_xll.HEAT($B9,$C9)");
        }

        [TestMethod]
        public void BigNumberInFormula()
        {
            test("30426000000/E7/1000");
        }

        [TestMethod]
        public void UDFWithNumericCharacters()
        {
            test("ASTRIP2_m(E9,E10)");
        }

        [TestMethod]
        public void UDFLikeNamedRangeCombination()
        {
            test("Prob1OptimalRiskyWeight(C7,C6,E7,G5)");
        }

        [TestMethod]
        public void UDFWithDot()
        {
            test("Functions.BScall(C3,C4,C5,C6,C7,C8)");
        }

        [TestMethod]
        public void NamedRangeReference()
        {
            test("SUM(TestRange)");
            test("SUM(Sheet1!TestRange)");   
        }

        [TestMethod]
        public void SheetWithUnderscore()
        {
            test("aap_noot!B12");
        }

        [TestMethod]
        public void SheetWithPeriod()
        {
            test("vrr2011_Omz.!M84");
            
        }

        [TestMethod]
        public void SheetAsString()
        {
            test("'[20]Algemene info + \"Overview\"'!T95");
        }

        [TestMethod]
        public void SheetWithQuote()
        {
            test("'Owner''s Engineer'!$A$2");
        }

        [TestMethod]
        public void ExternalSheetWithQuote()
        {
            test("'[1]Stacey''s Reconciliation'!C3");
        }

        [TestMethod]
        public void DirectSheetReference()
        {
            test("Sheet1!F7");
        }

        [TestMethod]
        public void SheetReference()
        {
            test("SUM(Sheet1!X1)");
        }

        [TestMethod]
        public void FileReference()
        {
            test("[2]Sheet1!X1");   
        }

        [TestMethod]
        public void FileReferenceInRange()
        {
            test("[2]Sheet1!X1:X10");    
        }

        [TestMethod]
        public void QuotedFileReference()
        {
            test("'[2]Sheet1'!X1");
        }


        [TestMethod]
        public void SheetReferenceRange()
        {
            test("SUM(Sheet1!X1:X6)");
        }

        [TestMethod]
        public void MultipleSheetsReferenceCell()
        {
            test("SUM(Sheet1:Sheet2!A1)");   
        }

        [TestMethod]
        public void MultipleSheetsReferenceRange()
        {
            test("SUM(Sheet1:Sheet2!A1:A3)");
        }

        [TestMethod]
        public void MultipleSheetsInFileReferenceCell()
        {   
            test("SUM([1]Sheet1:Sheet2!B15)");   
        }

        [TestMethod]
        public void RangeWithPrefixedRightLimitReference ()
        {   
            test("SUM(Deals!F9:Deals!F16)");   
        }

        [TestMethod]
        public void ReferencesInSingleColumn()
        {
            test("Sheet1!A:A");   
        }

        [TestMethod]
        public void ReferencesInLargeRange()
        {
            test("'B-Com'!A1:Z1048576");
        }

        [TestMethod]
        public void ReferenceFunctionAsArgument()
        {   
            test("ROUND(INDEX(A:A,1,1:1),1)");   
        }

        [TestMethod]
        public void RangeWithReferenceFunction()
        {
            test("SUM(A1:INDEX(A:A,1,1:1))");   
        }

        [TestMethod]
        public void UnionArgument()
        {
            test("LARGE((F38,C38:C48),1)");   
        }

        
        [TestMethod]
        public void DDE()
        {
            test("[1]!'INDU Index,[PX_close_5d]'");   
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException), "If formula can't be parsed")]
        public void ParsingFails()
        {
            test("]");
        }

        [TestMethod]
        public void EqualsIsFunction()
        {
            test("A1=A2", node => node.IsFunction() && node.GetFunction() == "=");
        }

        [TestMethod]
        public void PercentIsFunction()
        {
            test("1%", node => node.IsFunction() && node.GetFunction() == "%");
        }

        [TestMethod]
        public void FunctionsAsRefExpressions()
        {
            test("IF(TRUE,A1,A2):B5", "INDEX():B5", "MyUDFunction:B5", "Sheet!MyUDFunction:B5");
        }

        [TestMethod]
        public void Bug()
        {
            test("SUM(B5,2)");
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
            test(@"='C:\mypath\[myfile.xlsm]Sheet'!A1");
        }

        [TestMethod]
        public void TestFileNameString()
        {
            test("=[sheet]!A1", "=[sheet.xls]!A1");
        }

        [TestMethod]
        public void TestStructuredReferences()
        {
            // Examples from msdn support document about structured references: https://support.office.com/en-us/article/Using-structured-references-with-Excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
            test();
        }
    }
}
