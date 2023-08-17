using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace XLParser.Tests
{
    [TestClass]
    // Visual Studio standard data sources where tried for this class, but it was found to be very slow
    public class DatasetTests
    {
        public TestContext TestContext { get; set; }

        private const int MaxParseErrors = 10;

        [TestMethod]
        [TestCategory("Slow")]
        public void EnronFormulasParseTest()
        {
            ParseCsvDataSet("data/enron/formulas.txt", "data/enron/knownfails.txt");
        }

        [TestMethod]
        [TestCategory("Slow")]
        public void EusesFormulasParseTest()
        {
            ParseCsvDataSet("data/euses/formulas.txt", "data/euses/knownfails.txt");
        }

        [TestMethod]
        public void ParseTestFormulasStructuredReferences()
        {
            ParseCsvDataSet("data/testformulas/structured_references.txt");
        }

        [TestMethod]
        public void ParseTestFormulasUserContributed()
        {
            ParseCsvDataSet("data/testformulas/user_contributed.txt");
        }

        private void ParseCsvDataSet(string filename, string knownFailsFile = null)
        {
            ISet<string> knownfails = new HashSet<string>(ReadFormulaCsv(knownFailsFile));
            var parseErrors = 0;
            var lockObj = new object();

            Parallel.ForEach(ReadFormulaCsv(filename), (formula, control, lineNumber) =>
            {
                if (parseErrors > MaxParseErrors)
                {
                    control.Stop();
                    return;
                }

                try
                {
                    ParserTests.Test(formula);
                }
                catch (ArgumentException)
                {
                    if (!knownfails.Contains(formula))
                    {
                        lock (lockObj)
                        {
#if !_NET6_
                            TestContext.WriteLine($"Failed parsing line {lineNumber} <<{formula}>>");
#endif
                            parseErrors++;
                        }
                    }
                }
            });
            if (parseErrors > 0)
            {
                Assert.Fail("Parse Errors on file " + filename);
            }
        }

        private static IEnumerable<string> ReadFormulaCsv(string f)
        {
            return f == null ? new string[0] : File.ReadLines(f).Where(line => !string.IsNullOrWhiteSpace(line)).Select(UnQuote);
        }

        private static string UnQuote(string line)
        {
            return line.Length > 0 && line[0] == '"' ? line.Substring(1, line.Length - 2).Replace("\"\"", "\"") : line;
        }
    }
}
