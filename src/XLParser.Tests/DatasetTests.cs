using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace XLParser.Tests
{
    [TestClass]
    public class DatasetTests
    {
        public TestContext TestContext { get; set; }

        private int failedParses;
        private const int MaxParseErrors = 10;

        [TestInitialize]
        public void CheckFails()
        {
            // After enough parse errors, skip the rest
            if (failedParses > MaxParseErrors) Assert.Inconclusive("Skipping due to high number of parse errors");
        }

        [TestCleanup]
        public void CountFails()
        {
            if (TestContext.CurrentTestOutcome != UnitTestOutcome.Passed) failedParses++;
        }


        [TestMethod,
            DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV",
                "data\\enron_formulas.csv", "enron_formulas#csv", DataAccessMethod.Sequential)
        ]
        [TestCategory("Slow")]
        // Comment this to execute the test
        [Ignore]
        public void EnronFormulasParseTest()
        {
            var formula = TestContext.DataRow[0].ToString();
            ExcelFormulaParser.Parse(formula);
        }

        
        [TestMethod,
            DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV",
                "data\\euses_formulas.csv", "euses_formulas#csv", DataAccessMethod.Sequential)
        ]
        [TestCategory("Slow")]
        // Comment this to execute the test
        [Ignore]
        public void EusesFormulasParseTest()
        {
            var formula = TestContext.DataRow[0].ToString();
            ExcelFormulaParser.Parse(formula);
        }
    }
}
