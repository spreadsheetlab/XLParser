using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace XLParser.Tests
{
    [TestClass]
    public class PrefixInfoTests
    {
        private PrefixInfo test(string prefix)
        {
            var pi = ExcelFormulaParser.Parse($"{prefix}A1").AllNodes(GrammarNames.Prefix).First().GetPrefixInfo();
            Assert.AreEqual(prefix, pi.ToString());
            return pi;
        }

        [TestMethod]
        public void TestSheet()
        {
            var t = test("Sheet!");
            Assert.IsTrue(t.HasSheet);
            Assert.IsFalse(t.HasFile || t.HasFileName || t.HasFileNumber || t.HasFilePath || t.HasMultipleSheets || t.IsQuoted);
            Assert.AreEqual("Sheet", t.Sheet);
        }

        [TestMethod]
        public void TestSheetQuoted()
        {
            var t = test("'Sheet'!");
            Assert.IsTrue(t.HasSheet && t.IsQuoted);
            Assert.IsFalse(t.HasFile || t.HasFileName || t.HasFileNumber || t.HasFilePath || t.HasMultipleSheets);
        }

        [TestMethod]
        public void TestFileNumericSheet()
        {
            var t = test("[1]Sheet!");
            Assert.IsTrue(t.HasSheet && t.HasFileNumber && t.HasFile);
            Assert.IsFalse(t.HasFileName || t.HasFilePath || t.HasMultipleSheets || t.IsQuoted);
            Assert.AreEqual(1, t.FileNumber);
            Assert.AreEqual("Sheet", t.Sheet);
        }

        [TestMethod]
        public void TestFileStringSheet()
        {
            var t = test("[file.xls]Sheet!");
            Assert.IsTrue(t.HasSheet && t.HasFileName && t.HasFile);
            Assert.IsFalse(t.HasFileNumber || t.HasFilePath || t.HasMultipleSheets || t.IsQuoted);
            Assert.AreEqual("file.xls", t.FileName);
            Assert.AreEqual("Sheet", t.Sheet);
        }

        [TestMethod]
        public void TestQuotedFileStringSheet()
        {
            var t = test("'[file.xls]Sheet'!");
            Assert.IsTrue(t.HasSheet && t.HasFileName && t.IsQuoted && t.HasFile);
            Assert.IsFalse(t.HasFileNumber || t.HasFilePath || t.HasMultipleSheets);
            Assert.AreEqual("file.xls", t.FileName);
            Assert.AreEqual("Sheet", t.Sheet);
        }

        [TestMethod]
        public void TestQuotedPathFileStringSheet()
        {
            var t = test(@"'C:\path\[file.xls]Sheet'!");
            Assert.IsTrue(t.HasSheet && t.HasFileName && t.IsQuoted && t.HasFilePath && t.HasFile);
            Assert.IsFalse(t.HasFileNumber || t.HasMultipleSheets);
            Assert.AreEqual(@"C:\path\", t.FilePath);
            Assert.AreEqual("file.xls", t.FileName);
            Assert.AreEqual("Sheet", t.Sheet);
        }

        [TestMethod]
        public void TestMultiplesheets()
        {
            var t = test("Sheet1:Sheet3!");
            Assert.IsTrue(t.HasMultipleSheets);
            Assert.IsFalse(t.HasFile || t.HasFileName || t.HasFileNumber || t.HasFilePath || t.HasSheet || t.IsQuoted);
            Assert.AreEqual("Sheet1:Sheet3", t.MultipleSheets);
        }

        [TestMethod]
        public void TestFileMultiplesheets()
        {
            var t = test("[1]Sheet1:Sheet3!");
            Assert.IsTrue(t.HasMultipleSheets && t.HasFile && t.HasFileNumber);
            Assert.IsFalse(t.HasFileName || t.HasFilePath || t.HasSheet || t.IsQuoted);
            Assert.AreEqual("Sheet1:Sheet3", t.MultipleSheets);
        }
    }
}
