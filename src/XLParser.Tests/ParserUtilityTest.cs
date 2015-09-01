using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace XLParser.Tests
{
    [TestClass]
    public class ParserUtilityTest
    {
        [TestMethod]
        public void TestIsBuiltInFunction()
        {
            var builtins = ExcelFormulaParser.Parse("VLOOKUP(1)")
                .AllNodes()
                .Where(ExcelFormulaParser.IsBuiltinFunction)
                .Select(ExcelFormulaParser.GetFunction);
            CollectionAssert.AreEqual(builtins.ToList(), new [] {"VLOOKUP"});
        }
    }
}
