using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Irony.Parsing;
using System.Globalization;
using XLParser;

namespace XLParser.Tests
{
    [TestClass]
    public class FormulaAnalysisTest
    {
        #region Numbers()
        [TestMethod]
        public void FixedNumbers()
        {
            var fa = new FormulaAnalyzer("SUM(A1,3.1+4)");
            var nums = fa.Numbers().ToList();

            CollectionAssert.Contains(nums, 3.1);
            CollectionAssert.Contains(nums, 4.0);
            Assert.AreEqual(2, nums.Count());
        }

        [TestMethod]
        public void NoFixedNumbers()
        {
            var fa = new FormulaAnalyzer("A1+B2");
            var nums = fa.Numbers().ToList();
            Assert.AreEqual(0, nums.Count());
        }

        [TestMethod]
        public void FixedInt()
        {
            var fa = new FormulaAnalyzer("3");
            var nums = fa.Numbers().ToList();
            CollectionAssert.Contains(nums, 3.0);
            Assert.AreEqual(1, nums.Count());
        }

        [TestMethod]
        public void FixedReal()
        {
            var fa = new FormulaAnalyzer("0.1");
            var nums = fa.Numbers().ToList();
            CollectionAssert.Contains(nums, 0.1);
            Assert.AreEqual(1, nums.Count());
        }

        [TestMethod]
        public void Duplicate()
        {
            var fa = new FormulaAnalyzer("3+3");
            var nums = fa.Numbers().ToList();
            Assert.AreEqual(2, nums.Count());
        }

        // Test for issue #21: https://github.com/spreadsheetlab/XLParser/issues/21
        [TestMethod]
        public void NegativeNumber()
        {
            var fa = new FormulaAnalyzer("=-8+A1");
            var nums = fa.Numbers().ToList();
            CollectionAssert.AreEqual(new[] { -8.0 }, nums);
        }
        #endregion

        #region Functions()
        [TestMethod]
        public void CountInfixOperations()
        {
            var fa = new FormulaAnalyzer("3+4/5");
            var ops = fa.Functions().ToList();
            Assert.AreEqual(2, ops.Count());
            CollectionAssert.Contains(ops, "+");
            CollectionAssert.Contains(ops, "/");
        }

        [TestMethod]
        public void CountFunctionOperations()
        {
            var fa = new FormulaAnalyzer("SUM(3,5)");
            var ops = fa.Functions().ToList();
            Assert.AreEqual(1, ops.Count());
            CollectionAssert.Contains(ops, "SUM");
        }

        [TestMethod]
        public void CountPostfixOperations()
        {
            var fa = new FormulaAnalyzer("A6%");
            var ops = fa.Functions().ToList();
            Assert.AreEqual(1, ops.Count());
            CollectionAssert.Contains(ops, "%");
        }

        [TestMethod]
        public void DontCountSheetReferenes()
        {
            string formula = "Weight!B1";
            var fa = new FormulaAnalyzer(formula);
            Assert.AreEqual(0, fa.Functions().Count());
        }

        [TestMethod]
        public void CountComparisons()
        {
            var fa = new FormulaAnalyzer("IF(A1>A2,3,4)");
            var ops = fa.Functions().Distinct().ToList();
            Assert.AreEqual(2, ops.Count());
            CollectionAssert.Contains(ops, ">");
            CollectionAssert.Contains(ops, "IF");
        }

        [TestMethod]
        public void ComparisonIsFunction()
        {
            var fa = new FormulaAnalyzer("IF(A1<=A2,A1+1,A2)");
            CollectionAssert.Contains(fa.Functions().ToList(), "<=");
        }

        [TestMethod]
        public void IntersectIsFunction()
        {
            var fa = new FormulaAnalyzer("SUM(A1:A3 A2:A3)");
            var functions = fa.Functions().Distinct().ToList();
            CollectionAssert.Contains(functions, "INTERSECT");
        }

        [TestMethod]
        public void IntersectIsAtCorrectPosition()
        {
            var fa = new FormulaAnalyzer("SUM(A1:A3 A2:A3)");
            ParseTreeNode intersect = fa.AllNodes.FirstOrDefault(node => node.Token?.Terminal?.Name == "INTERSECT");
            Assert.IsNotNull(intersect);
            Assert.AreEqual(9, intersect.Span.Location.Position);
        }
        #endregion

        #region References()

        [TestMethod]
        public void OnlyDirectReferences()
        {
            // Make sure A1:A10 isn't returned as "A1:A10", "A1" and "A10"
            var fa = new FormulaAnalyzer("SUM(A1:A10)");
            var references = fa.References().ToList();
            CollectionAssert.AreEqual(references.Select(ExcelFormulaParser.Print).ToList(), new [] { "A1:A10" });
            fa = new FormulaAnalyzer("(A1)+2");
            references = fa.References().ToList();
            CollectionAssert.AreEqual(references.Select(ExcelFormulaParser.Print).ToList(), new [] {"A1" });
        }

        [TestMethod]
        public void SimpleReference()
        {
            List<ParserReference> References = new FormulaAnalyzer("B1").ParserReferences().ToList();

            //we get a cell reference
            Assert.AreEqual(true, References.First().ReferenceType == ReferenceType.Cell);

            Assert.AreEqual(1, References.Count);
            Assert.AreEqual("B1", References.First().MinLocation.ToString());
        }

        [TestMethod]
        public void RangeReference()
        {
            List<ParserReference> References = new FormulaAnalyzer("SUM(B1:C5)").ParserReferences().ToList();

            //we get a range reference
            Assert.AreEqual(false, References.First().ReferenceType == ReferenceType.Cell);
        }

        [TestMethod]
        public void NamedRangeReference()
        {
            List<ParserReference> References = new FormulaAnalyzer("SUM(TestRange)").ParserReferences().ToList();
            List<ParserReference> PrefixedReferences = new FormulaAnalyzer("SUM(Sheet1!TestRange)").ParserReferences().ToList();

            //we get a range reference
            Assert.AreEqual(false, References.First().ReferenceType == ReferenceType.Cell);

            Assert.AreEqual(1, References.Count);
            Assert.AreEqual("TestRange", References.First().Name);

            Assert.AreEqual(false, PrefixedReferences.First().ReferenceType == ReferenceType.Cell);
            Assert.AreEqual("TestRange", PrefixedReferences.First().Name);
        }

        [TestMethod]
        public void NamedRangeWithUnderscoreReference()
        {
            List<ParserReference> result = new FormulaAnalyzer("_XX1/100").ParserReferences().ToList();
            Assert.AreEqual(1, result.Count);
            Assert.IsFalse(result.First().ReferenceType == ReferenceType.Cell);
            Assert.IsTrue(result.First().ReferenceType == ReferenceType.UserDefinedName);
        }

        [TestMethod]
        public void TableReference()
        {
            List<ParserReference> result = new FormulaAnalyzer("COUNTA(Table1)").ParserReferences().ToList();
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(ReferenceType.UserDefinedName, result.First().ReferenceType);
            Assert.AreEqual("Table1", result.First().Name);
        }

        [TestMethod]
        public void StructuredTableReferenceHeader()
        {
            List<ParserReference> result = new FormulaAnalyzer("COUNTA(Table1[#Header])").ParserReferences().ToList();
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(ReferenceType.Table, result.First().ReferenceType);
            Assert.AreEqual("Table1", result.First().Name);
        }

        [TestMethod]
        public void StructuredTableReferenceCurrentRow()
        {
            List<ParserReference> result = new FormulaAnalyzer("COUNTA(Table1[@])").ParserReferences().ToList();
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(ReferenceType.Table, result.First().ReferenceType);
            Assert.AreEqual("Table1", result.First().Name);
        }

        [TestMethod]
        public void StructuredTableReferenceColumns()
        {
            List<ParserReference> result = new FormulaAnalyzer("COUNTA(Table1[[Date]:[Color]])").ParserReferences().ToList();
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(ReferenceType.Table, result.First().ReferenceType);
            Assert.AreEqual("Table1", result.First().Name);
        }

        [TestMethod]
        public void StructuredTableReferenceRowAndColumn()
        {
            List<ParserReference> result = new FormulaAnalyzer("Table1[[#Totals],[Qty]]").ParserReferences().ToList();
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(ReferenceType.Table, result.First().ReferenceType);
            Assert.AreEqual("Table1", result.First().Name);
        }

        [TestMethod]
        public void SheetWithUnderscore()
        {
            ParseTree parseResult = ExcelFormulaParser.ParseToTree("aap_noot!B12");
            Assert.AreEqual(3, parseResult.Tokens.Count);
            Assert.AreEqual("aap_noot!", ((Token)parseResult.Tokens.First()).Text);

            Assert.AreEqual(GrammarNames.Formula, parseResult.Root.Term.Name);
            Assert.AreNotEqual(ParseTreeStatus.Error, parseResult.Status);

            Assert.AreEqual(1, parseResult.Root.ChildNodes.Count());
            ParseTreeNode reference = parseResult.Root.ChildNodes.First();
            Assert.AreEqual(GrammarNames.Reference, reference.Term.Name);

            Assert.AreEqual(2, reference.ChildNodes.Count());
            ParseTreeNode prefix = reference.ChildNodes.First();
            Assert.AreEqual(GrammarNames.Prefix, prefix.Term.Name);
            ParseTreeNode cellOrRange = reference.ChildNodes.ElementAt(1);
            Assert.AreEqual(GrammarNames.Cell, cellOrRange.Term.Name);

            Assert.AreEqual(1, prefix.ChildNodes.Count());
            ParseTreeNode sheet = prefix.ChildNodes.First();
            Assert.AreEqual(GrammarNames.TokenSheet, sheet.Term.Name);
            Assert.AreEqual("aap_noot!", sheet.Token.Value);
        }

        [TestMethod]
        public void SheetWithPeriod()
        {
            ParseTree parseResult = ExcelFormulaParser.ParseToTree("vrr2011_Omz.!M84");
            Assert.AreNotEqual(ParseTreeStatus.Error, parseResult.Status);
        }

        [TestMethod]
        public void SheetAsString()
        {
            ParseTree parseResult = ExcelFormulaParser.ParseToTree("'[20]Algemene info + \"Overview\"'!T95");
            Assert.AreNotEqual(ParseTreeStatus.Error, parseResult.Status);
        }

        [TestMethod]
        public void SheetWithQuote()
        {
            List<ParserReference> result = new FormulaAnalyzer("'Owner''s Engineer'!$A$2").ParserReferences().ToList();
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("Owner's Engineer", result.First().Worksheet);
        }

        [TestMethod]
        public void ExternalSheetWithQuote()
        {
            List<ParserReference> result = new FormulaAnalyzer("'[1]Stacey''s Reconciliation'!C3").ParserReferences().ToList();
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("1", result.First().FileName);
            Assert.AreEqual("Stacey's Reconciliation", result.First().Worksheet);
        }

        [TestMethod]
        public void NonNumericExternalSheet()
        {
            List<ParserReference> result = new FormulaAnalyzer("[externalFile.xlsx]Book1!C3").ParserReferences().ToList();
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("externalFile.xlsx", result.First().FileName);
            Assert.AreEqual("Book1", result.First().Worksheet);
        }

        [TestMethod]
        public void DirectSheetReference()
        {
            List<ParserReference> References = new FormulaAnalyzer("Sheet1!F7").ParserReferences().ToList();

            //we get a cell reference
            Assert.AreEqual(true, References.First().ReferenceType == ReferenceType.Cell);

            Assert.AreEqual(1, References.Count);
            Assert.AreEqual("F7", References.First().MinLocation.ToString());
            Assert.AreEqual("Sheet1", References.First().Worksheet);
        }

        [TestMethod]
        public void SheetReference()
        {
            List<ParserReference> References = new FormulaAnalyzer("SUM(Sheet1!X1)").ParserReferences().ToList();

            //we get a cell reference
            Assert.AreEqual(true, References.First().ReferenceType == ReferenceType.Cell);

            Assert.AreEqual(1, References.Count);
            Assert.AreEqual("X1", References.First().MinLocation.ToString());
            Assert.AreEqual("Sheet1", References.First().Worksheet);
        }

        [TestMethod]
        public void FileReference()
        {
            List<ParserReference> References = new FormulaAnalyzer("[2]Sheet1!X1").ParserReferences().ToList();

            //we get a cell reference
            Assert.IsTrue(References.First().ReferenceType == ReferenceType.Cell);
            var cellref = References.First();

            Assert.AreEqual(1, References.Count);
            Assert.AreEqual("X1", cellref.MinLocation.ToString());
            Assert.AreEqual("Sheet1", cellref.Worksheet);
            Assert.AreEqual("2", cellref.FileName);
        }

        [TestMethod]
        public void FileReferenceInRange()
        {
            List<ParserReference> References = new FormulaAnalyzer("[2]Sheet1!X1:X10").ParserReferences().ToList();

            //we get a cell reference
            Assert.IsFalse(References.First().ReferenceType == ReferenceType.Cell);
            var cellref = References.First();

            Assert.AreEqual(1, References.Count);
            Assert.AreEqual("X1", cellref.MinLocation.ToString());
            Assert.AreEqual("Sheet1", cellref.Worksheet);
            Assert.AreEqual("2", cellref.FileName);
        }

        [TestMethod]
        public void QuotedFileReference()
        {
            List<ParserReference> references = new FormulaAnalyzer("'[2]Sheet1'!X1").ParserReferences().ToList();

            //we get a cell reference
            Assert.IsNotNull(references);
            Assert.IsTrue(references.First().ReferenceType == ReferenceType.Cell);
            var cellref = references.First();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual("X1", cellref.MinLocation.ToString());
            Assert.AreEqual("2", cellref.FileName);
            Assert.AreEqual("Sheet1", cellref.Worksheet);
        }


        [TestMethod]
        public void SheetReferenceRange()
        {
            List<ParserReference> References = new FormulaAnalyzer("SUM(Sheet1!X1:X6)").ParserReferences().ToList();

            //we get a cell reference
            Assert.IsFalse(References.First().ReferenceType == ReferenceType.Cell);
        }

        [TestMethod]
        public void MultipleSheetsReferenceCell()
        {
            String formula = "SUM(Sheet1:Sheet2!A1)";

            List<ParserReference> References = new FormulaAnalyzer(formula).ParserReferences().ToList();

            Assert.AreEqual(1, References.Count);
            Assert.AreEqual(true, References.First().ReferenceType == ReferenceType.Cell);
            Assert.AreEqual("A1", References.First().MinLocation.ToString());
            Assert.AreEqual("Sheet1", References.First().Worksheet);
            Assert.AreEqual("Sheet2", References.First().LastWorksheet);
        }

        [TestMethod]
        public void MultipleSheetsReferenceRange()
        {
            String formula = "SUM(Sheet1:Sheet2!A1:A3)";

            List<ParserReference> References = new FormulaAnalyzer(formula).ParserReferences().ToList();

            Assert.AreEqual(1, References.Count);
            Assert.IsTrue(References.First().ReferenceType == ReferenceType.CellRange);
            Assert.AreEqual("A1", References.First().MinLocation.ToString());
            Assert.AreEqual("A3", References.First().MaxLocation.ToString());
            Assert.AreEqual("Sheet1", References.First().Worksheet);
            Assert.AreEqual("Sheet2", References.First().LastWorksheet);
        }

        [TestMethod]
        public void MultipleSheetsInFileReferenceCell()
        {
            String formula = "SUM([1]Sheet1:Sheet2!B15)";

            List<ParserReference> References = new FormulaAnalyzer(formula).ParserReferences().ToList();

            Assert.AreEqual(1, References.Count);
            Assert.AreEqual(true, References.First().ReferenceType == ReferenceType.Cell);
            Assert.AreEqual("B15", References.First().MinLocation.ToString());
            Assert.AreEqual("Sheet1", References.First().Worksheet);
            Assert.AreEqual("Sheet2", References.First().LastWorksheet);
        }

        [TestMethod]
        public void RangeWithPrefixedRightLimitReference()
        {
            List<ParserReference> References = new FormulaAnalyzer("SUM(Deals!F9:Deals!F16)").ParserReferences().ToList();
            Assert.IsFalse(References.First().ReferenceType == ReferenceType.Cell);
            Assert.AreEqual("Deals!F9:Deals!F16", References.First().LocationString);
            Assert.AreEqual("Deals", References.First().Worksheet);

            //quotes should be omitted
            References = new FormulaAnalyzer("SUM(Deals!F9:'Deals'!F16)").ParserReferences().ToList();
            Assert.IsFalse(References.First().ReferenceType == ReferenceType.Cell);
            Assert.AreEqual("Deals!F9:'Deals'!F16", References.First().LocationString);
            Assert.AreEqual("Deals", References.First().Worksheet);

            //external file references
            References = new FormulaAnalyzer("SUM([1]Deals!F9:[1]Deals!F16)").ParserReferences().ToList();
            Assert.IsFalse(References.First().ReferenceType == ReferenceType.Cell);
            Assert.AreEqual("[1]Deals!F9:[1]Deals!F16", References.First().LocationString);
            Assert.AreEqual("Deals", References.First().Worksheet);
            Assert.AreEqual("1", References.First().FileName);
        }

        [TestMethod]
        public void ReferencesInSingleColumn()
        {
            List<ParserReference> result = new FormulaAnalyzer("Sheet1!A:A").ParserReferences().ToList();
            Assert.AreEqual(1, result.Count);
            Assert.IsFalse(result.First().ReferenceType == ReferenceType.Cell);
        }

        [TestMethod]
        public void ReferencesInLargeRange()
        {
            // expanding large range completely may lead to out of 
            // memory exception. Verify this is not the case.
            List<ParserReference> result = new FormulaAnalyzer("'B-Com'!A1:Z1048576").ParserReferences().ToList();
            Assert.AreEqual(1, result.Count());
            Assert.IsFalse(result.First().ReferenceType == ReferenceType.Cell);
        }

        [TestMethod]
        public void ReferenceFunctionAsArgument()
        {
            List<ParserReference> References = new FormulaAnalyzer("ROUND(INDEX(A:A,1,1:1),1)").ParserReferences().ToList();

            //no reference for the index function
            Assert.AreEqual(2, References.Count);
            Assert.AreEqual("A:A", References.First().LocationString);
            Assert.AreEqual("1:1", References.Last().LocationString);
        }

        [TestMethod]
        public void RangeWithReferenceFunction()
        {
            List<ParserReference> References = new FormulaAnalyzer("SUM(A1:INDEX(A:A,1,1:1))").ParserReferences().ToList();

            //no reference for the range with the index function
            Assert.AreEqual(3, References.Count);
            Assert.AreEqual("A1", References[0].LocationString);
            Assert.AreEqual("A:A", References[1].LocationString);
            Assert.AreEqual("1:1", References[2].LocationString);
        }

        [TestMethod]
        public void ArrayAsArgument()
        {
            List<ParserReference> References = new FormulaAnalyzer("LARGE((F38,C38:C48),1)").ParserReferences().ToList();

            Assert.AreEqual(2, References.Count);
            Assert.AreEqual("F38", References.First().MinLocation.ToString());
            Assert.AreEqual("C38:C48", References.Last().LocationString);
        }

        #endregion

        #region Depth() and ConditionalComplexity()

        [TestMethod]
        public void TestDepth()
        {
            var fa = new FormulaAnalyzer("SUM(1,2+SUM(3),3)");
            Assert.AreEqual(fa.Depth(), 4);
            Assert.AreEqual(fa.OperatorDepth(), 3);
        }

        [TestMethod]
        public void TestConditionalComplexity()
        {
            var fa1 = new FormulaAnalyzer("1");
            Assert.AreEqual(fa1.ConditionalComplexity(), 0);
            var fa2 = new FormulaAnalyzer("IF(TRUE,IF(FALSE,1,0),0)");
            Assert.AreEqual(fa2.ConditionalComplexity(), 2);
        }
        #endregion

        #region Constants()

        [TestMethod]
        public void TestConstants()
        {
            var fa = new FormulaAnalyzer("1*3-8-(-9)+VLOOKUP($A:$B,5,10)&\"ABC\"+TRUE*-3");
            var constants = fa.Constants().ToList();
            CollectionAssert.AreEqual(new[] { "1", "3", "8", "-9", "5", "10", "\"ABC\"", "TRUE", "-3"}, constants);
        }
        #endregion
    }
}