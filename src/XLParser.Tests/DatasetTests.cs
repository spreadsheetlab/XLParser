using System;
using System.Collections;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Irony.Parsing;

namespace XLParser.Tests
{
    [TestClass]
    // Visual studio standard datasources where tried for this class, but it was found very slow
    public class DatasetTests
    {
        public TestContext TestContext { get; set; }

        private const int MaxParseErrors = 10;

        [TestMethod]
        [TestCategory("Slow")]
        // Uncomment this to execute the test
        //[Ignore]
        public void EnronFormulasParseTest()
        {
            parseCSVDataSet("data/enron_formulas.csv", "data/enron_knownfails.csv");
        }

        [TestCategory("Slow")]
        // Uncomment this to execute the test
        //[Ignore]
        public void EusesFormulasParseTest()
        {
            parseCSVDataSet("data/euses_formulas.csv", "data/enron_knownfails.csv");
        }

        private void parseCSVDataSet(string filename, string knownfailsfile = null)
        {
            ISet<string> knownfails = new HashSet<string>(readFormulaCSV(knownfailsfile));
            int parseErrors = 0;
            var LOCK = new object();
            Parallel.ForEach(readFormulaCSV(filename), (formula, control, linenr) =>
            {
                if (parseErrors > MaxParseErrors)
                {
                    control.Stop();
                    return;
                }
                try
                {
                    ExcelFormulaParser.Parse(formula);
                }
                catch (ArgumentException e)
                {
                    if (!knownfails.Contains(formula))
                    {
                        lock (LOCK)
                        {
                            TestContext.WriteLine(String.Format("Failed parsing line {0} <<{1}>>", linenr, formula));
                            parseErrors++;
                        }
                    }
                }
            });
            if (parseErrors > 0) Assert.Fail("Parse Errors on file " + filename);
        }

        private static IEnumerable<string> readFormulaCSV(string f)
        {
            if (f == null) return Enumerable.Empty<string>();
            return File.ReadLines(f)
                .Where(line => line != "")
                .Select(unQuote)
                ;
        }

        private static string unQuote(string line)
        {
            return line.Length > 0 && line[0] == '"' ?
                    line.Substring(1, line.Length - 2).Replace("\"\"", "\"")
                  : line;
        }
    }
}
