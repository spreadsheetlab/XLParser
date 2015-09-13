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
    // Visual studio standard datasources where tried for this class, but it was found to be very slow
    public class DatasetTests
    {
        public TestContext TestContext { get; set; }

        private const int MaxParseErrors = 10;

        [TestMethod]
        [TestCategory("Slow")]
        public void EnronFormulasParseTest()
        {
            parseCSVDataSet("data/enron/formulas.txt", "data/enron/knownfails.txt");
        }

        [TestMethod]
        [TestCategory("Slow")]
        public void EusesFormulasParseTest()
        {
            parseCSVDataSet("data/euses/formulas.txt", "data/euses/knownfails.txt");
        }

        [TestMethod]
        public void ParseTestFormulasStructuredReferences()
        {
            parseCSVDataSet("data/testformulas/structured_references.txt");
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
                    ParserTests.test(formula);
                }
                catch (ArgumentException)
                {
                    if (!knownfails.Contains(formula))
                    {
                        lock (LOCK)
                        {
                            TestContext.WriteLine($"Failed parsing line {linenr} <<{formula}>>");
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
            // using ReadAllLines instead of ReadLines shaves about 10s of the enron test, so it's worth the memory usage.
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
