using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Irony.Parsing;

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
            List<ParserReference> references = new FormulaAnalyzer("B1").ParserReferences().ToList();
            
            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Cell, references.First().ReferenceType);
            Assert.AreEqual("B1", references.First().MinLocation);
        }

        [TestMethod]
        public void RangeReference()
        {
            List<ParserReference> references = new FormulaAnalyzer("SUM(B1:C5)").ParserReferences().ToList();
            
            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.CellRange, references.First().ReferenceType);
            Assert.AreEqual("B1", references.First().MinLocation);
            Assert.AreEqual("C5", references.First().MaxLocation);
        }

        [TestMethod]
        public void VRangeReferenceLong()
        {
            List<ParserReference> references = new FormulaAnalyzer("SUM(XBC:ABC)").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.VerticalRange, references.First().ReferenceType);
            Assert.AreEqual("XBC", references.First().MinLocation);
            Assert.AreEqual("ABC", references.First().MaxLocation);
        }

        [TestMethod]
        public void VRangeReferenceLarge()
        {
            List<ParserReference> references = new FormulaAnalyzer("SUM(XEZ:XFD)").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.VerticalRange, references.First().ReferenceType);
            Assert.AreEqual("XEZ", references.First().MinLocation);
            Assert.AreEqual("XFD", references.First().MaxLocation);
        }

        [TestMethod]
        public void VRangeReferenceShort()
        {
            List<ParserReference> references = new FormulaAnalyzer("SUM(AB:ZZ)").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.VerticalRange, references.First().ReferenceType);
            Assert.AreEqual("AB", references.First().MinLocation);
            Assert.AreEqual("ZZ", references.First().MaxLocation);
        }

        [TestMethod]
        public void NamedRangeReference()
        {
            List<ParserReference> references = new FormulaAnalyzer("SUM(TestRange)").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.UserDefinedName, references.First().ReferenceType);
            Assert.AreEqual("TestRange", references.First().Name);
        }

        [TestMethod]
        public void NamedRangeWithPrefixReference()
        {
            List<ParserReference> references = new FormulaAnalyzer("SUM(Sheet1!TestRange)").ParserReferences().ToList();
            
            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.UserDefinedName, references.First().ReferenceType);
            Assert.AreEqual("TestRange", references.First().Name);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
        }

        [TestMethod]
        public void NamedRangeWithUnderscoreReference()
        {
            List<ParserReference> references = new FormulaAnalyzer("_XX1/100").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.UserDefinedName, references.First().ReferenceType);
            Assert.AreEqual("_XX1", references.First().Name);
        }

        [TestMethod]
        public void TableReference()
        {
            List<ParserReference> references = new FormulaAnalyzer("COUNTA(Table1)").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.UserDefinedName, references.First().ReferenceType);
            Assert.AreEqual("Table1", references.First().Name);
        }

        [TestMethod]
        public void StructuredTableReferenceHeaders()
        {
            List<ParserReference> references = new FormulaAnalyzer("COUNTA(Table1[#Headers])").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Table, references.First().ReferenceType);
            Assert.AreEqual("Table1", references.First().Name);
            CollectionAssert.AreEqual(new[] {"#Headers"}, references.First().TableSpecifiers);
            CollectionAssert.AreEqual(new string[] {}, references.First().TableColumns);
        }

        [TestMethod]
        public void StructuredTableReferenceThisRow()
        {
            List<ParserReference> references = new FormulaAnalyzer("COUNTA(Table1[[#This Row],[b]])").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Table, references.First().ReferenceType);
            Assert.AreEqual("Table1", references.First().Name);
            CollectionAssert.AreEqual(new[] {"#This Row"}, references.First().TableSpecifiers);
            CollectionAssert.AreEqual(new[] {"b"}, references.First().TableColumns);
        }

        [TestMethod]
        public void StructuredTableReferenceCurrentRow()
        {
            List<ParserReference> references = new FormulaAnalyzer("COUNTA(Table1[@Region])").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Table, references.First().ReferenceType);
            Assert.AreEqual("Table1", references.First().Name);
            CollectionAssert.AreEqual(new[] {"@"}, references.First().TableSpecifiers);
            CollectionAssert.AreEqual(new[] {"Region"}, references.First().TableColumns);
        }

        // See https://github.com/spreadsheetlab/XLParser/issues/173
        [TestMethod]
        public void StructuredTableReferenceWholeRow()
        {
            List<ParserReference> references = new FormulaAnalyzer("=INDEX(Table1[@],2)").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Table, references.First().ReferenceType);
            Assert.AreEqual("Table1", references.First().Name);
            CollectionAssert.AreEqual(new[] {"@"}, references.First().TableSpecifiers);
            CollectionAssert.AreEqual(new string[] {}, references.First().TableColumns);
        }

        [TestMethod]
        public void StructuredTableReferenceColumns()
        {
            List<ParserReference> references = new FormulaAnalyzer("COUNTA(Table1[[Date]:[Color]])").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Table, references.First().ReferenceType);
            Assert.AreEqual("Table1", references.First().Name);
            CollectionAssert.AreEqual(new string[] {}, references.First().TableSpecifiers);
            CollectionAssert.AreEqual(new[] {"Date", "Color"}, references.First().TableColumns);
        }

        [TestMethod]
        public void StructuredTableReferenceRowAndColumn()
        {
            List<ParserReference> references = new FormulaAnalyzer("Table1[[#Totals],[Qty]]").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Table, references.First().ReferenceType);
            Assert.AreEqual("Table1", references.First().Name);
            CollectionAssert.AreEqual(new[] {"#Totals"}, references.First().TableSpecifiers);
            CollectionAssert.AreEqual(new[] {"Qty"}, references.First().TableColumns);
        }

        [TestMethod]
        public void StructuredTableReferenceWithEscapedCharacterPound()
        {
            List<ParserReference> references = new FormulaAnalyzer("=SUBTOTAL(109,Table1['#NumItems])").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Table, references.First().ReferenceType);
            Assert.AreEqual("Table1", references.First().Name);
            CollectionAssert.AreEqual(new string[] {}, references.First().TableSpecifiers);
            CollectionAssert.AreEqual(new[] {"#NumItems"}, references.First().TableColumns);
        }

        [TestMethod]
        public void StructuredTableReferenceWithEscapedCharacterBracket()
        {
            List<ParserReference> references = new FormulaAnalyzer("COUNTA(Table1['[Header']])").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Table, references.First().ReferenceType);
            Assert.AreEqual("Table1", references.First().Name);
            CollectionAssert.AreEqual(new string[] {}, references.First().TableSpecifiers);
            CollectionAssert.AreEqual(new[] {"[Header]"}, references.First().TableColumns);
        }

        [TestMethod]
        public void StructuredTableReferenceWithEscapedCharacterQuote()
        {
            List<ParserReference> references = new FormulaAnalyzer("COUNTA(Table1[''Test''])").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Table, references.First().ReferenceType);
            Assert.AreEqual("Table1", references.First().Name);
            CollectionAssert.AreEqual(new string[] {}, references.First().TableSpecifiers);
            CollectionAssert.AreEqual(new[] {"'Test'"}, references.First().TableColumns);
        }

        [TestMethod]
        public void StructuredTableReferenceUnqualified()
        {
            List<ParserReference> references = new FormulaAnalyzer("=SUBTOTAL(109,[Sales])").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Table, references.First().ReferenceType);
            Assert.AreEqual(null, references.First().Name);
            CollectionAssert.AreEqual(new string[] {}, references.First().TableSpecifiers);
            CollectionAssert.AreEqual(new[] {"Sales"}, references.First().TableColumns);
        }

        [TestMethod]
        public void StructuredTableReferenceUnqualifiedWithSpaceInName()
        {
            List<ParserReference> references = new FormulaAnalyzer("=SUBTOTAL(109,[Sales Amount])").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Table, references.First().ReferenceType);
            Assert.AreEqual(null, references.First().Name);
            CollectionAssert.AreEqual(new string[] {}, references.First().TableSpecifiers);
            CollectionAssert.AreEqual(new[] {"Sales Amount"}, references.First().TableColumns);
        }

        [TestMethod]
        public void StructuredTableReferenceUnqualifiedWithNumbers()
        {
            List<ParserReference> references = new FormulaAnalyzer("=SUBTOTAL(109,[2016])").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Table, references.First().ReferenceType);
            Assert.AreEqual(null, references.First().Name);
            CollectionAssert.AreEqual(new string[] {}, references.First().TableSpecifiers);
            CollectionAssert.AreEqual(new[] {"2016"}, references.First().TableColumns);
        }

        [TestMethod]
        public void StructuredTableReferenceSpecifierAndColumns()
        {
            List<ParserReference> references = new FormulaAnalyzer("=DeptSales[[#All],[Sales Amount]:[Commission Amount]]").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Table, references.First().ReferenceType);
            Assert.AreEqual("DeptSales", references.First().Name);
            CollectionAssert.AreEqual(new[] {"#All"}, references.First().TableSpecifiers);
            CollectionAssert.AreEqual(new[] {"Sales Amount", "Commission Amount"}, references.First().TableColumns);
        }

        [TestMethod]
        public void StructuredTableReferenceSpecifiersAndColumn()
        {
            List<ParserReference> references = new FormulaAnalyzer("=DeptSales[[#Headers],[#Data],[% Commission]]").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Table, references.First().ReferenceType);
            Assert.AreEqual("DeptSales", references.First().Name);
            CollectionAssert.AreEqual(new[] {"#Headers", "#Data"}, references.First().TableSpecifiers);
            CollectionAssert.AreEqual(new[] {"% Commission"}, references.First().TableColumns);
        }

        [TestMethod]
        public void StructuredTableReferenceMultipleRows()
        {
            List<ParserReference> references = new FormulaAnalyzer("=SUM([@Jan]:[@Feb])").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Table, references.First().ReferenceType);
            Assert.AreEqual(null, references.First().Name);
            CollectionAssert.AreEqual(new[] {"@"}, references.First().TableSpecifiers);
            CollectionAssert.AreEqual(new[] {"Jan", "Feb"}, references.First().TableColumns);
        }

        [TestMethod]
        public void StructuredTableReferenceMultipleColumns()
        {
            List<ParserReference> references = new FormulaAnalyzer("=XLOOKUP($G7,Sales[[Region]:[Region]],Sales[Mar])").ParserReferences().ToList();

            Assert.AreEqual(3, references.Count);

            Assert.AreEqual(ReferenceType.Cell, references[0].ReferenceType);
            Assert.AreEqual("$G7", references[0].MinLocation);

            Assert.AreEqual(ReferenceType.Table, references[1].ReferenceType);
            Assert.AreEqual("Sales", references[1].Name);
            CollectionAssert.AreEqual(new string[] {}, references[1].TableSpecifiers);
            CollectionAssert.AreEqual(new[] {"Region", "Region"}, references[1].TableColumns);

            Assert.AreEqual(ReferenceType.Table, references[2].ReferenceType);
            Assert.AreEqual("Sales", references[2].Name);
            CollectionAssert.AreEqual(new string[] {}, references[2].TableSpecifiers);
            CollectionAssert.AreEqual(new[] {"Mar"}, references[2].TableColumns);
        }

        [TestMethod]
        public void StructuredTableReferenceMultipleHeaders()
        {
            List<ParserReference> references = new FormulaAnalyzer("=COUNTA(Sales_5[[#Headers],[Jan]]:Sales_5[[#Headers],[Mar]])").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);

            Assert.AreEqual(ReferenceType.Table, references[0].ReferenceType);
            Assert.AreEqual("Sales_5", references[0].Name);
            CollectionAssert.AreEqual(new[] {"#Headers"}, references[0].TableSpecifiers);
            CollectionAssert.AreEqual(new[] {"Jan", "Mar"}, references[0].TableColumns);
        }

        [TestMethod]
        public void StructuredTableReferenceIntersectColumns()
        {
            List<ParserReference> references = new FormulaAnalyzer("=SUM(Sales_2[Jan]:Sales_2[Feb])").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Table, references.First().ReferenceType);
            Assert.AreEqual("Sales_2", references.First().Name);
            CollectionAssert.AreEqual(new string[] {}, references.First().TableSpecifiers);
            CollectionAssert.AreEqual(new[] {"Jan", "Feb"}, references.First().TableColumns);
        }

        [TestMethod]
        public void StructuredTableReferenceIntersectDifferentSpecifiers()
        {
            List<ParserReference> references = new FormulaAnalyzer("=COUNTA(Sales_2[[#Headers],[Jan]]:Sales_2[[#Data],[Feb]])").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Table, references.First().ReferenceType);
            Assert.AreEqual("Sales_2", references.First().Name);
            CollectionAssert.AreEqual(new string[] {}, references.First().TableSpecifiers);
            CollectionAssert.AreEqual(new[] {"Jan", "Feb"}, references.First().TableColumns);
        }

        [TestMethod]
        public void StructuredTableReferenceIntersectDifferentTables()
        {
            List<ParserReference> references = new FormulaAnalyzer("=COUNTA(Sales_2[Jan]:Sales_4[Feb])").ParserReferences().ToList();

            Assert.AreEqual(2, references.Count);

            Assert.AreEqual(ReferenceType.Table, references[0].ReferenceType);
            Assert.AreEqual("Sales_2", references[0].Name);
            CollectionAssert.AreEqual(new string[] {}, references[0].TableSpecifiers);
            CollectionAssert.AreEqual(new[] {"Jan"}, references[0].TableColumns);

            Assert.AreEqual(ReferenceType.Table, references[1].ReferenceType);
            Assert.AreEqual("Sales_4", references[1].Name);
            CollectionAssert.AreEqual(new string[] {}, references[1].TableSpecifiers);
            CollectionAssert.AreEqual(new[] {"Feb"}, references[1].TableColumns);
        }

        [TestMethod]
        public void StructuredTableReferenceWithSheetReference()
        {
            // See https://github.com/spreadsheetlab/XLParser/issues/170
            List<ParserReference> references = new FormulaAnalyzer("=VLOOKUP([@ProductNumber],Sheet2!A:B,1,FALSE)").ParserReferences().ToList();

            Assert.AreEqual(2, references.Count);

            Assert.AreEqual(ReferenceType.Table, references[0].ReferenceType);
            Assert.AreEqual(null, references[0].Name);
            CollectionAssert.AreEqual(new[] {"@"}, references[0].TableSpecifiers);
            CollectionAssert.AreEqual(new[] {"ProductNumber"}, references[0].TableColumns);

            Assert.AreEqual(ReferenceType.VerticalRange, references[1].ReferenceType);
            Assert.AreEqual("Sheet2", references[1].Worksheet);
            Assert.AreEqual("A", references[1].MinLocation);
            Assert.AreEqual("B", references[1].MaxLocation);
        }

        [TestMethod]
        public void SheetWithUnderscore()
        {
            ParseTree parseResult = ExcelFormulaParser.ParseToTree("aap_noot!B12");
            Assert.AreEqual(3, parseResult.Tokens.Count);
            Assert.AreEqual("aap_noot!", ((Token)parseResult.Tokens.First()).Text);

            Assert.AreEqual(GrammarNames.Formula, parseResult.Root.Term.Name);
            Assert.AreNotEqual(ParseTreeStatus.Error, parseResult.Status);

            Assert.AreEqual(1, parseResult.Root.ChildNodes.Count);
            ParseTreeNode reference = parseResult.Root.ChildNodes.First();
            Assert.AreEqual(GrammarNames.Reference, reference.Term.Name);

            Assert.AreEqual(2, reference.ChildNodes.Count);
            ParseTreeNode prefix = reference.ChildNodes.First();
            Assert.AreEqual(GrammarNames.Prefix, prefix.Term.Name);
            ParseTreeNode cellOrRange = reference.ChildNodes.ElementAt(1);
            Assert.AreEqual(GrammarNames.Cell, cellOrRange.Term.Name);

            Assert.AreEqual(1, prefix.ChildNodes.Count);
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
            List<ParserReference> references = new FormulaAnalyzer("'Owner''s Engineer'!$A$2").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual("Owner's Engineer", references.First().Worksheet);
        }

        [TestMethod]
        public void ExternalSheetWithQuote()
        {
            List<ParserReference> references = new FormulaAnalyzer("'[1]Stacey''s Reconciliation'!C3").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual("1", references.First().FileName);
            Assert.AreEqual("Stacey's Reconciliation", references.First().Worksheet);
        }

        [TestMethod]
        public void NonNumericExternalSheet()
        {
            List<ParserReference> references = new FormulaAnalyzer("[externalFile.xlsx]Book1!C3").ParserReferences().ToList();
            
            Assert.AreEqual(1, references.Count);
            Assert.AreEqual("externalFile.xlsx", references.First().FileName);
            Assert.AreEqual("Book1", references.First().Worksheet);
        }

        [TestMethod]
        public void ExternalWorkbook()
        {
            List<ParserReference> references = new FormulaAnalyzer(@"='C:\Users\Test\Desktop\[Book1.xlsx]Sheet1'!$A$1").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(@"C:\Users\Test\Desktop\", references.First().FilePath);
            Assert.AreEqual("Book1.xlsx", references.First().FileName);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
        }

        [TestMethod]
        public void ExternalWorkbookNetworkPath()
        {
            List<ParserReference> references = new FormulaAnalyzer(@"='\\TEST-01\Folder\[Book1.xlsx]Sheet1'!$A$1").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(@"\\TEST-01\Folder\", references.First().FilePath);
            Assert.AreEqual("Book1.xlsx", references.First().FileName);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
        }

        [TestMethod]
        public void ExternalWorkbookUrlPathHttp()
        {
            List<ParserReference> references = new FormulaAnalyzer(@"='http://example.com/test/[Book1.xlsx]Sheet1'!$A$1").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(@"http://example.com/test/", references.First().FilePath);
            Assert.AreEqual("Book1.xlsx", references.First().FileName);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
        }

        [TestMethod]
        public void ExternalWorkbookUrlPathHttpWithSpaceInPath()
        {
            // See [#138](https://github.com/spreadsheetlab/XLParser/issues/138)
            List<ParserReference> references = new FormulaAnalyzer(@"='http://example.com/test folder/[Book1.xlsx]Sheet1'!$A$1").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(@"http://example.com/test folder/", references.First().FilePath);
            Assert.AreEqual("Book1.xlsx", references.First().FileName);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
        }

        [TestMethod]
        public void ExternalWorkbookUrlPathHttpWithSpaceInDocument()
        {
            // See [#138](https://github.com/spreadsheetlab/XLParser/issues/138)
            List<ParserReference> references = new FormulaAnalyzer(@"='http://example.com/testfolder/[Book 1.xlsx]Sheet1'!$A$1").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(@"http://example.com/testfolder/", references.First().FilePath);
            Assert.AreEqual("Book 1.xlsx", references.First().FileName);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
        }

        [TestMethod]
        public void ExternalWorkbookNetworkPathWithSpace()
        {
            // See [#142](https://github.com/spreadsheetlab/XLParser/issues/142)
            List<ParserReference> references = new FormulaAnalyzer(@"='\\networkshare\test folder$\[Book 1.xlsx]Sheet1'!$C$33").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(@"\\networkshare\test folder$\", references.First().FilePath);
            Assert.AreEqual("Book 1.xlsx", references.First().FileName);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
        }

        [TestMethod]
        public void ExternalWorkbookNetworkPathWithoutSpaceInFolder()
        {
            // See [#142](https://github.com/spreadsheetlab/XLParser/issues/142)
            List<ParserReference> references = new FormulaAnalyzer(@"='\\networkshare\testfolder$\[Book 1.xlsx]Sheet1'!$C$33").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(@"\\networkshare\testfolder$\", references.First().FilePath);
            Assert.AreEqual("Book 1.xlsx", references.First().FileName);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
        }

        [TestMethod]
        public void ExternalWorkbookNetworkPathWithQuotes()
        {
            // See [#135](https://github.com/spreadsheetlab/XLParser/issues/135)
            List<ParserReference> references = new FormulaAnalyzer(@"='\\SRV01\[TestFile.xls]TestSheet'!#REF!").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(@"\\SRV01\", references.First().FilePath);
            Assert.AreEqual("TestFile.xls", references.First().FileName);
            Assert.AreEqual("TestSheet", references.First().Worksheet);
        }

        [TestMethod]
        public void ExternalWorkbookNetworkPathWithQuotesAndSpaces()
        {
            // See [#135](https://github.com/spreadsheetlab/XLParser/issues/135)
            List<ParserReference> references = new FormulaAnalyzer(@"='\\SRV01\[Test File.xls]Test Sheet'!#REF!").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(@"\\SRV01\", references.First().FilePath);
            Assert.AreEqual("Test File.xls", references.First().FileName);
            Assert.AreEqual("Test Sheet", references.First().Worksheet);
        }

        [TestMethod]
        public void ExternalWorkbookNetworkPathWithQuotesAndSpaceInFolder()
        {
            // See [#135](https://github.com/spreadsheetlab/XLParser/issues/135)
            List<ParserReference> references = new FormulaAnalyzer(@"='\\SRV01\Test Folder\[TestFile.xls]Test Sheet'!#REF!").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(@"\\SRV01\Test Folder\", references.First().FilePath);
            Assert.AreEqual("TestFile.xls", references.First().FileName);
            Assert.AreEqual("Test Sheet", references.First().Worksheet);
        }

        [TestMethod]
        public void ExternalWorkbookNetworkPathWithQuotesAndSpaceInDocument()
        {
            // See [#135](https://github.com/spreadsheetlab/XLParser/issues/135)
            List<ParserReference> references = new FormulaAnalyzer(@"='\\SRV01\TestFolder\[Test File.xls]Test Sheet'!#REF!").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(@"\\SRV01\TestFolder\", references.First().FilePath);
            Assert.AreEqual("Test File.xls", references.First().FileName);
            Assert.AreEqual("Test Sheet", references.First().Worksheet);
        }

        [TestMethod]
        public void ExternalWorkbookUrlPathHttps()
        {
            List<ParserReference> references = new FormulaAnalyzer(@"='https://d.docs.live.net/3fade139bf25879f/Documents/[Tracer.xlsx]Sheet2'!$C$5").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(@"https://d.docs.live.net/3fade139bf25879f/Documents/", references.First().FilePath);
            Assert.AreEqual("Tracer.xlsx", references.First().FileName);
            Assert.AreEqual("Sheet2", references.First().Worksheet);
        }

        [TestMethod]
        public void ExternalWorkbookDocumentWithoutBracketsWithSpace()
        {
            // See [#137](https://github.com/spreadsheetlab/XLParser/issues/137)
            List<ParserReference> references = new FormulaAnalyzer(@"='https://preview.perfectxl.com/Financial Sample.xlsx'!financials").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(@"https://preview.perfectxl.com/", references.First().FilePath);
            Assert.AreEqual("Financial Sample.xlsx", references.First().FileName);
            Assert.AreEqual("financials", references.First().Name);
        }

        [TestMethod]
        public void ExternalWorkbookDocumentWithoutBrackets()
        {
            // See [#137](https://github.com/spreadsheetlab/XLParser/issues/135)
            List<ParserReference> references = new FormulaAnalyzer(@"='https://preview.perfectxl.com/FinancialSample.xlsx'!financials").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(@"https://preview.perfectxl.com/", references.First().FilePath);
                Assert.AreEqual("FinancialSample.xlsx", references.First().FileName);
            Assert.AreEqual("financials", references.First().Name);
        }

        [TestMethod]
        public void ExternalWorkbookRelativePath()
        {
            List<ParserReference> references = new FormulaAnalyzer(@"='Test\Folder\[Book1.xlsx]Sheet1'!$A$1").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(@"Test\Folder\", references.First().FilePath);
            Assert.AreEqual("Book1.xlsx", references.First().FileName);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
        }

        [TestMethod]
        public void ExternalWorkbookRelativePathIsDefinedName()
        {
            List<ParserReference> references = new FormulaAnalyzer(@"=Test\Folder").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(@"Test\Folder", references.First().Name);
            Assert.AreEqual(ReferenceType.UserDefinedName, references.First().ReferenceType);
        }

        [TestMethod]
        public void ExternalWorkbookUrlPathHttpWithRoundBracketsInDocument()
        {
            // See [#140](https://github.com/spreadsheetlab/XLParser/issues/140)
            List<ParserReference> references = new FormulaAnalyzer(@"='http://example.com/testfolder(brackets)/[Book 1.xlsx]Sheet1'!$A$1").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(@"http://example.com/testfolder(brackets)/", references.First().FilePath);
            Assert.AreEqual("Book 1.xlsx", references.First().FileName);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
        }

        [TestMethod]
        public void ExternalWorkbookUrlPathHttpWithRoundBrackets()
        {
            // See [#140](https://github.com/spreadsheetlab/XLParser/issues/140)
            List<ParserReference> references = new FormulaAnalyzer(@"='http://example.com/testfolder(brackets)/[Book (1).xlsx]Sheet1'!$A$1").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(@"http://example.com/testfolder(brackets)/", references.First().FilePath);
            Assert.AreEqual("Book (1).xlsx", references.First().FileName);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
        }

        [TestMethod]
        public void ExternalWorkbookPathWithRoundBracketsInDocument()
        {
            // See [#140](https://github.com/spreadsheetlab/XLParser/issues/140)
            List<ParserReference> references = new FormulaAnalyzer(@"='c:\My documents\[Book 1 (copy).xlsx]Sheet1'!$A$1").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(@"c:\My documents\", references.First().FilePath);
            Assert.AreEqual("Book 1 (copy).xlsx", references.First().FileName);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
        }

        [TestMethod]
        public void ExternalWorkbookPathWithRoundBrackets()
        {
            // See [#140](https://github.com/spreadsheetlab/XLParser/issues/140)
            List<ParserReference> references = new FormulaAnalyzer(@"='c:\My documents (copy)\[Book 1 (copy).xlsx]Sheet1'!$A$1").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(@"c:\My documents (copy)\", references.First().FilePath);
            Assert.AreEqual("Book 1 (copy).xlsx", references.First().FileName);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
        }

        [TestMethod]
        public void ExternalWorkbookSingleCell()
        {
            List<ParserReference> references = new FormulaAnalyzer(@"='C:\Users\Test\Desktop\[Book1.xlsx]Sheet1'!$A$1").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Cell, references.First().ReferenceType);
            Assert.AreEqual("Book1.xlsx", references.First().FileName);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
        }

        [TestMethod]
        public void ExternalWorkbookCellRange()
        {
            List<ParserReference> references = new FormulaAnalyzer(@"='C:\Users\Test\Desktop\[Book1.xlsx]Sheet1'!$A$1:$A$10").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.CellRange, references.First().ReferenceType);
            Assert.AreEqual("Book1.xlsx", references.First().FileName);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
        }

        [TestMethod]
        public void ExternalWorkbookDefinedNameLocalScope()
        {
            List<ParserReference> references = new FormulaAnalyzer(@"='C:\Users\Test\Desktop\[Book1.xlsx]Sheet1'!FirstItem").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.UserDefinedName, references.First().ReferenceType);
            Assert.AreEqual("Book1.xlsx", references.First().FileName);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
            Assert.AreEqual("FirstItem", references.First().Name);
        }

        [TestMethod]
        public void ExternalWorkbookDefinedNameGlobalScope()
        {
            List<ParserReference> references = new FormulaAnalyzer(@"='C:\Users\Test\Desktop\Book1.xlsx'!Items").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.UserDefinedName, references.First().ReferenceType);
            Assert.AreEqual("Book1.xlsx", references.First().FileName);
            Assert.AreEqual("Items", references.First().Name);
        }

        [TestMethod]
        public void MultipleExternalWorkbookSingleCell()
        {
            List<ParserReference> references = new FormulaAnalyzer(@"=SUM('C:\Users\Test\Desktop\[Book1.xlsx]Sheet1'!$A$1, 'C:\Users\Test\Desktop\[Book1.xlsx]Sheet1'!$A$2)").ParserReferences().ToList();
            Assert.AreEqual(2, references.Count);
            
            Assert.AreEqual(ReferenceType.Cell, references[0].ReferenceType);
            Assert.AreEqual("Book1.xlsx", references[0].FileName);
            Assert.AreEqual("Sheet1", references[0].Worksheet);
            Assert.AreEqual("$A$1", references[0].MinLocation);

            Assert.AreEqual(ReferenceType.Cell, references[1].ReferenceType);
            Assert.AreEqual("Book1.xlsx", references[1].FileName);
            Assert.AreEqual("Sheet1", references[1].Worksheet);
            Assert.AreEqual("$A$2", references[1].MinLocation);
        }

        [TestMethod]
        public void MultipleExternalWorkbookCellRange()
        {
            List<ParserReference> references = new FormulaAnalyzer(@"=SUM('C:\Users\Test\Desktop\[Book1.xlsx]Sheet1'!$A$1:$A$10, 'C:\Users\Test\Desktop\[Book1.xlsx]Sheet1'!$A$11:$A$20)").ParserReferences().ToList();
            Assert.AreEqual(2, references.Count);
            
            Assert.AreEqual(ReferenceType.CellRange, references[0].ReferenceType);
            Assert.AreEqual("Book1.xlsx", references[0].FileName);
            Assert.AreEqual("Sheet1", references[0].Worksheet);
            Assert.AreEqual("$A$1", references[0].MinLocation);
            Assert.AreEqual("$A$10", references[0].MaxLocation);

            Assert.AreEqual(ReferenceType.CellRange, references[1].ReferenceType);
            Assert.AreEqual("Book1.xlsx", references[1].FileName);
            Assert.AreEqual("Sheet1", references[1].Worksheet);
            Assert.AreEqual("$A$11", references[1].MinLocation);
            Assert.AreEqual("$A$20", references[1].MaxLocation);
        }

        [TestMethod]
        public void MultipleExternalWorkbookDefinedNameLocalScope()
        {
            List<ParserReference> references = new FormulaAnalyzer(@"=SUM('C:\Users\Test\Desktop\[Book1.xlsx]Sheet1'!FirstItem, 'C:\Users\Test\Desktop\[Book1.xlsx]Sheet1'!SecondItem)").ParserReferences().ToList();
            Assert.AreEqual(2, references.Count);
            
            Assert.AreEqual(ReferenceType.UserDefinedName, references[0].ReferenceType);
            Assert.AreEqual("Book1.xlsx", references[0].FileName);
            Assert.AreEqual("Sheet1", references[0].Worksheet);
            Assert.AreEqual("FirstItem", references[0].Name);

            Assert.AreEqual(ReferenceType.UserDefinedName, references[1].ReferenceType);
            Assert.AreEqual("Book1.xlsx", references[1].FileName);
            Assert.AreEqual("Sheet1", references[1].Worksheet);
            Assert.AreEqual("SecondItem", references[1].Name);
        }

        [TestMethod]
        public void MultipleExternalWorkbookDefinedNameGlobalScope()
        {
            List<ParserReference> references = new FormulaAnalyzer(@"=SUM('C:\Users\Test\Desktop\Book1.xlsx'!Items, 'C:\Users\Test\Desktop\Book1.xlsx'!Items2)").ParserReferences().ToList();
            Assert.AreEqual(2, references.Count);
            
            Assert.AreEqual(ReferenceType.UserDefinedName, references[0].ReferenceType);
            Assert.AreEqual("Book1.xlsx", references[0].FileName);
            Assert.AreEqual("Items", references[0].Name);

            Assert.AreEqual(ReferenceType.UserDefinedName, references[1].ReferenceType);
            Assert.AreEqual("Book1.xlsx", references[1].FileName);
            Assert.AreEqual("Items2", references[1].Name);
        }

        [TestMethod]
        public void ExternalWorkbookWithQuoteInSheet()
        {
            List<ParserReference> references = new FormulaAnalyzer(@"='[Book1''s.xlsm]Sheet1'!$A$1").ParserReferences().ToList();
            Assert.AreEqual(1, references.Count);

            Assert.AreEqual(ReferenceType.Cell, references[0].ReferenceType);
            Assert.AreEqual("Book1''s.xlsm", references[0].FileName);
        }

        [TestMethod]
        public void ExternalWorkbooksWithATSign()
        {
            List<ParserReference> references = new FormulaAnalyzer(@"='\\example.com@ssl\sites\[Data.xls]Hoi'!$A$1:$B$1").ParserReferences().ToList();
            Assert.AreEqual(1, references.Count);

            Assert.AreEqual(ReferenceType.CellRange, references[0].ReferenceType);
            Assert.AreEqual("Data.xls", references[0].FileName);
        }

        [TestMethod]
        public void ExternalWorkbookWithColonAndSlash()
        {
            List<ParserReference> references = new FormulaAnalyzer(@"='C:\Users\somebody\http:\\\my.sharepoint.com\docs\[Articles.xlsx]My Articles'!A1:D24").ParserReferences().ToList();
            Assert.AreEqual(1, references.Count);

            Assert.AreEqual(ReferenceType.CellRange, references[0].ReferenceType);
            Assert.AreEqual("Articles.xlsx", references[0].FileName);
        }


        [TestMethod]
        public void ExternalWorkbookQuotesInPath()
        {
            List<ParserReference> references = new FormulaAnalyzer(@"=SUM('\\Users\Test\Desktop''s\Book1.xlsx'!Items)").ParserReferences().ToList();
            Assert.AreEqual(1, references.Count);

            Assert.AreEqual(ReferenceType.UserDefinedName, references[0].ReferenceType);
            Assert.AreEqual("Book1.xlsx", references[0].FileName);
        }

        [TestMethod]
        public void MultipleExternalWorkbooksQuoteInPath()
        {
            List<ParserReference> references = new FormulaAnalyzer(@" =SUM('\\Users\Test\Desktop\Book1.xlsx'!Items, '\\Users\Test\Desktop\Books''s2.xlsx'!Items2)").ParserReferences().ToList();
            Assert.AreEqual(2, references.Count);

            Assert.AreEqual(ReferenceType.UserDefinedName, references[0].ReferenceType);
            Assert.AreEqual("Book1.xlsx", references[0].FileName);

            Assert.AreEqual(ReferenceType.UserDefinedName, references[1].ReferenceType);
            Assert.AreEqual("Books''s2.xlsx", references[1].FileName);
        }

        [TestMethod]
        public void DirectSheetReference()
        {
            List<ParserReference> references = new FormulaAnalyzer("Sheet1!F7").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Cell, references.First().ReferenceType);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
            Assert.AreEqual("F7", references.First().MinLocation);
        }

        [TestMethod]
        public void SheetReference()
        {
            List<ParserReference> references = new FormulaAnalyzer("SUM(Sheet1!X1)").ParserReferences().ToList();
            
            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Cell, references.First().ReferenceType);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
            Assert.AreEqual("X1", references.First().MinLocation);
        }

        [TestMethod]
        public void FileReference()
        {
            List<ParserReference> references = new FormulaAnalyzer("[2]Sheet1!X1").ParserReferences().ToList();

            Assert.AreEqual(ReferenceType.Cell, references.First().ReferenceType);

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Cell, references.First().ReferenceType);
            Assert.AreEqual("2", references.First().FileName);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
            Assert.AreEqual("X1", references.First().MinLocation);
        }

        [TestMethod]
        public void FileReferenceInRange()
        {
            List<ParserReference> references = new FormulaAnalyzer("[2]Sheet1!X1:X10").ParserReferences().ToList();

            Assert.AreEqual(ReferenceType.CellRange, references.First().ReferenceType);

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.CellRange, references.First().ReferenceType);
            Assert.AreEqual("2", references.First().FileName);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
            Assert.AreEqual("X1", references.First().MinLocation);
            Assert.AreEqual("X10", references.First().MaxLocation);
        }

        [TestMethod]
        public void QuotedFileReference()
        {
            List<ParserReference> references = new FormulaAnalyzer("'[2]Sheet1'!X1").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Cell, references.First().ReferenceType);
            Assert.AreEqual("2", references.First().FileName);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
            Assert.AreEqual("X1", references.First().MinLocation);
        }


        [TestMethod]
        public void SheetReferenceRange()
        {
            List<ParserReference> references = new FormulaAnalyzer("SUM(Sheet1!X1:X6)").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.CellRange, references.First().ReferenceType);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
            Assert.AreEqual("X1", references.First().MinLocation);
            Assert.AreEqual("X6", references.First().MaxLocation);
        }

        [TestMethod]
        public void MultipleSheetsReferenceCell()
        {
            List<ParserReference> references = new FormulaAnalyzer("SUM(Sheet1:Sheet2!A1)").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Cell, references.First().ReferenceType);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
            Assert.AreEqual("Sheet2", references.First().LastWorksheet);
            Assert.AreEqual("A1", references.First().MinLocation);
        }

        [TestMethod]
        public void MultipleSheetsReferenceRange()
        {
            List<ParserReference> references = new FormulaAnalyzer("SUM(Sheet1:Sheet2!A1:A3)").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.CellRange, references.First().ReferenceType);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
            Assert.AreEqual("Sheet2", references.First().LastWorksheet);
            Assert.AreEqual("A1", references.First().MinLocation);
            Assert.AreEqual("A3", references.First().MaxLocation);
        }

        [TestMethod]
        public void MultipleSheetsInFileReferenceCell()
        {
            List<ParserReference> references = new FormulaAnalyzer("SUM([1]Sheet1:Sheet2!B15)").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.Cell, references.First().ReferenceType);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
            Assert.AreEqual("Sheet2", references.First().LastWorksheet);
            Assert.AreEqual("B15", references.First().MinLocation);
        }

        [TestMethod]
        public void RangeWithPrefixedRightLimitReference()
        {
            List<ParserReference> references = new FormulaAnalyzer("SUM(Deals!F9:Deals!F16)").ParserReferences().ToList();

            Assert.AreEqual(ReferenceType.CellRange, references.First().ReferenceType);
            Assert.AreEqual("Deals", references.First().Worksheet);
            Assert.AreEqual("Deals!F9:Deals!F16", references.First().LocationString);

            //quotes should be omitted
            references = new FormulaAnalyzer("SUM(Deals!F9:'Deals'!F16)").ParserReferences().ToList();

            Assert.AreEqual(ReferenceType.CellRange, references.First().ReferenceType);
            Assert.AreEqual("Deals", references.First().Worksheet);
            Assert.AreEqual("Deals!F9:'Deals'!F16", references.First().LocationString);

            //external file references
            references = new FormulaAnalyzer("SUM([1]Deals!F9:[1]Deals!F16)").ParserReferences().ToList();

            Assert.AreEqual(ReferenceType.CellRange, references.First().ReferenceType);
            Assert.AreEqual("1", references.First().FileName);
            Assert.AreEqual("Deals", references.First().Worksheet);
            Assert.AreEqual("[1]Deals!F9:[1]Deals!F16", references.First().LocationString);
        }

        [TestMethod]
        public void ReferencesInSingleColumn()
        {
            List<ParserReference> references = new FormulaAnalyzer("Sheet1!A:A").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.VerticalRange, references.First().ReferenceType);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
            Assert.AreEqual("A", references.First().MinLocation);
            Assert.AreEqual("A", references.First().MaxLocation);
        }

        [TestMethod]
        public void ReferencesInMultipleColumns()
        {
            List<ParserReference> references = new FormulaAnalyzer("Sheet1!B:F").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.VerticalRange, references.First().ReferenceType);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
            Assert.AreEqual("B", references.First().MinLocation);
            Assert.AreEqual("F", references.First().MaxLocation);
        }

        [TestMethod]
        public void ReferencesInSingleRow()
        {
            List<ParserReference> references = new FormulaAnalyzer("Sheet1!1:1").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.HorizontalRange, references.First().ReferenceType);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
            Assert.AreEqual("1", references.First().MinLocation);
            Assert.AreEqual("1", references.First().MaxLocation);
        }

        [TestMethod]
        public void ReferencesInMultipleRows()
        {
            List<ParserReference> references = new FormulaAnalyzer("Sheet1!2:36").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.HorizontalRange, references.First().ReferenceType);
            Assert.AreEqual("Sheet1", references.First().Worksheet);
            Assert.AreEqual("2", references.First().MinLocation);
            Assert.AreEqual("36", references.First().MaxLocation);
        }

        [TestMethod]
        public void ReferencesInLargeRange()
        {
            // expanding large range completely may lead to out of 
            // memory exception. Verify this is not the case.
            List<ParserReference> references = new FormulaAnalyzer("'B-Com'!A1:Z1048576").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.CellRange, references.First().ReferenceType);
            Assert.AreEqual("B-Com", references.First().Worksheet);
            Assert.AreEqual("A1", references.First().MinLocation);
            Assert.AreEqual("Z1048576", references.First().MaxLocation);
        }

        [TestMethod]
        public void AbsoluteColumnReference()
        {
            List<ParserReference> references = new FormulaAnalyzer("$A:$A").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.VerticalRange, references.First().ReferenceType);
            Assert.AreEqual("$A", references.First().MinLocation);
            Assert.AreEqual("$A", references.First().MaxLocation);
        }

        [TestMethod]
        public void AbsoluteRowReference()
        {
            List<ParserReference> references = new FormulaAnalyzer("$1:$1").ParserReferences().ToList();

            Assert.AreEqual(1, references.Count);
            Assert.AreEqual(ReferenceType.HorizontalRange, references.First().ReferenceType);
            Assert.AreEqual("$1", references.First().MinLocation);
            Assert.AreEqual("$1", references.First().MaxLocation);
        }

        [TestMethod]
        public void ReferenceFunctionAsArgument()
        {
            List<ParserReference> references = new FormulaAnalyzer("ROUND(INDEX(A:A,1,1:1),1)").ParserReferences().ToList();

            //no reference for the index function
            Assert.AreEqual(2, references.Count);
            Assert.AreEqual("A:A", references.First().LocationString);
            Assert.AreEqual("1:1", references.Last().LocationString);
        }

        [TestMethod]
        public void RangeWithReferenceFunction()
        {
            List<ParserReference> references = new FormulaAnalyzer("SUM(A1:INDEX(A:A,1,1:1))").ParserReferences().ToList();

            //no reference for the range with the index function
            Assert.AreEqual(3, references.Count);
            Assert.AreEqual("A1", references[0].LocationString);
            Assert.AreEqual("A:A", references[1].LocationString);
            Assert.AreEqual("1:1", references[2].LocationString);
        }

        [TestMethod]
        public void ArrayAsArgument()
        {
            List<ParserReference> references = new FormulaAnalyzer("LARGE((F38,C38:C48),1)").ParserReferences().ToList();

            Assert.AreEqual(2, references.Count);
            Assert.AreEqual("F38", references.First().MinLocation);
            Assert.AreEqual("C38:C48", references.Last().LocationString);
        }

        [TestMethod]
        public void UnionWithinParentheses()
        {
            List<ParserReference> references = new FormulaAnalyzer(@"=(A1:A3,C1:C3)").ParserReferences().ToList();
            Assert.AreEqual(2, references.Count);

            Assert.AreEqual("A1:A3", references[0].LocationString);
            Assert.AreEqual("C1:C3", references[1].LocationString);
        }

        [TestMethod]
        public void UnionWithoutParentheses()
        {
            // Valid when defining a name or conditional format range
            List<ParserReference> references = new FormulaAnalyzer(@"=A1:A3,C1:C3").ParserReferences().ToList();
            Assert.AreEqual(2, references.Count);

            Assert.AreEqual("A1:A3", references[0].LocationString);
            Assert.AreEqual("C1:C3", references[1].LocationString);
        }

        [TestMethod]
        public void MultiAreaRangeFormula()
        {
            List<ParserReference> references = new FormulaAnalyzer(@"=Sheet1!$A$1,Sheet1!$B$2").ParserReferences().ToList();
            Assert.AreEqual(2, references.Count);

            Assert.AreEqual("Sheet1!$A$1", references[0].LocationString);
            Assert.AreEqual("Sheet1!$B$2", references[1].LocationString);
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