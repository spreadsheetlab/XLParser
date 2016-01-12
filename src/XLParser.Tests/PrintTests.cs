using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Text.RegularExpressions;
using XLParser;
using Irony.Parsing;

namespace XLParser.Tests
{
    // Parameterized tests would be better for this class, but MSTest doesn't support them
    [TestClass]
    public class PrintTests
    {
        [TestMethod]
        public void TestNumber()
        {
            test("1", "1.5", "100");
            // exponential notation gets parsed to normal numbers so cannot be printed back
            // test("1e3");
        }

        [TestMethod]
        public void TestUDF()
        {
            test("FOO()", "FOO(BAR())", "FOO(1+1,BAR(1,5,100,\"abc\"))");
        }

        [TestMethod]
        public void TestExcelFunction()
        {
            test("VLOOKUP(1)", "SUM(A1:A5)");
        }

        [TestMethod]
        public void TestExcelFunctionCellRef()
        {
            test("INDEX(1)");
        }

        [TestMethod]
        public void TestReference()
        {
            test("A1", "ZZZ555",
            "A1:A5", "1:5", "A:Z",
            "Sheet!A1", "[1]Sheet!A1");
        }

        [TestMethod]
        public void TestRanges()
        {
            test("A1:A5", "1:5", "A:Z", "A1:SheetName!A5", "Sheet1!A1:Sheet5!A1");
        }

        [TestMethod]
        public void TestNamedRanges()
        {
            test("TEST", "Sheet1!TEST", "VLOOKUP", "[1]!TEST", "[1]Sheet!TEST");
        }

        [TestMethod]
        public void TestErrors()
        {
            test("#DIV/0!", "#NAME?", "#NUM!", "#NULL!", "#REF!", "#VALUE!");
        }

        [TestMethod]
        public void TestInfixPostfix()
        {
            test("1 + 1", "1 * 2 + 3", "100%", "A1:A5 1:5", "1 < 5", "1=1");
        }

        [TestMethod]
        public void TestIntersection()
        {
            test("A5:A10 B5:E5", false);
        }

        [TestMethod]
        public void TestArrayFormula()
        {
            test("{1,1}");
        }

        [TestMethod]
        public void TestBrackets()
        {
            test("(1)" , "((A1:A5))");
        }

        [TestMethod]
        public void TestEmptyArgument()
        {
            test("SUM()", "SUM(1,)", "SUM((1))", "SUM((A1))", "SUM((A1,A1),)", "SUM(A1,,A1)");
        }

        // From http://homepages.mcs.vuw.ac.nz/~elvis/db/Excel.shtml
        [TestMethod]
        public void TestElvisExamples()
        {
            test("1",
                "1+1",
                "A1",
                "$B$2",
                "SUM(B5:B15)",
                "SUM(B5:B15,D5:D15)",
                "SUM(B5:B15 A7:D7)",
                "SUM(sheet1!$A$1:$B$2)",
                "[1]sheet1!$A$1",
                "SUM((A:A 1:1))",
                "SUM((A:A,1:1))",
                "SUM((A:A A1:B1))",
                "SUM((D9:D11,(E9:E11,F9:F11)))",
                "=IF(P5=1.0,\"NA\",IF(P5=2.0,\"A\",IF(P5=3.0,\"B\",IF(P5=4.0,\"C\",IF(P5=5.0,\"D\",IF(P5=6.0,\"E\",IF(P5=7.0,\"F\",IF(P5=8.0,\"G\"))))))))",
                "{=SUM(B2:D2*B3:D3)}"
            );
        }

        [TestMethod]
        // Examples I had lying around
        public void TestVarious()
        {
            test("0","1", "-5",
              "0.01", "1.03",
              "a", "abc",
              "1 + 1", "5 / 8.9 + 9",
              "3%", "-3", "+3",
              "SUM(1,5,6,7,8)",
              "A1", "AZ55",
              "$B$9", "$B50",
              "B15:C99", "B$15:$C99",
              "E:E",
              "5:5",
              "Sheet1!A5",
              "#REF!",
              "(3 + 3) * 5",
              "1 * 3 * 5",
              "(5 * 5)",
              "2 ^ (3 * 4)",
              "(+D10-F10)/F10"
            );
        }

        [TestMethod]
        public void TestCalcCSVFails()
        {
            test("'Data Entry'!G31/('Data Entry'!G29/100)"
                ,"(('Data Entry'!G31-IFA!H57)/('Data Entry'!G30/100))-0.86"
                ,"ROUND(IF('Data Entry'!G45=\"Y\",0,IF(OR(D6>4999.999,'Data Entry'!G32<12),0,(1+((5000-D6)*0.000025))*D9)),0)"
                ,"ROUND((+'Data Entry'!D39*4)*ROUND('Calc Data'!D12,0),0)"
                , "IF('[1]Data Entry'!$F$61=0,'[1]Additional Aid'!$G$26,'[2]Savings Due to $295,000'!$G$37)"
                , "IF(E12>G12,E12-G12,0)"
                , "LN(E23*E28/E27)-0.5*E31^2"
                , "POWER((2*B13*EXP((B13+B5)*0.5*B10))/((B13+B5)*B14+2*B13),2*B5*B6/(B11^2))"
                ,"SUM(IF((DelPoint= \"hidalgo\")*(DType = \"pre\")*(OFFSET(DelPoint,0,B3+2)>0),OFFSET(DelPoint,0,B3+2),0))"
                ,"SUM(IF((DelPoint= \"PV\")*IF((DType = \"firm\")+(DType = \"econ\")>0,1,0)*(OFFSET(DelPoint,0,B3+2)>0),OFFSET(DelPoint,0,B3+2),0))"
                );
        }

        private static void test(params string[] formulas)
        {
            test((IEnumerable<string>)formulas);
        }

        private static void test(IEnumerable<string> formulas)
        {
            foreach (var formula in formulas)
            {
                test(formula);
            }
        }

        internal static void test(string formula, bool ignorewhitespace = true, ParseTreeNode parsed = null)
        {
            if (parsed == null)
            {
                parsed = ExcelFormulaParser.Parse(formula);
            }
            try
            {
                var printed = parsed.Print();
                if (ignorewhitespace)
                {
                    formula = Regex.Replace(formula, @"\s+", "");
                    printed = Regex.Replace(printed, @"\s+", "");
                }
                Assert.AreEqual(formula, printed, "Printed parsed formula differs from original.\nOriginal: '{0}'\nPrinted: '{1}'", formula, printed);
            }
            catch (ArgumentException e)
            {
                Assert.Fail("Parse Tree contains a node for which the Print function is not defined.\n{0}", e.Message);
            }
        }
    }
}
