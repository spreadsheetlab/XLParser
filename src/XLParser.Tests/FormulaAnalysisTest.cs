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
    }
}
