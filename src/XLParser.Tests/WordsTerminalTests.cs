using System;
using Irony.Parsing;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace XLParser.Tests
{
    [TestClass]
    public class WordsTerminalTests
    {
        [TestMethod]
        public void RecognizesAnyWordFromList()
        {
            var words = new[] { "PROC", "FUNC" };
            TestTerminalMatching("FUNC(1)", true, words);
            TestTerminalMatching("PROC(1)", true, words);
        }

        [TestMethod]
        public void RecognizesLongestWord()
        {
            TestTerminalMatching("ACOSH(4)", false, new[] { "ACOS", "ACOSH" }, t => t == "ACOSH");
        }

        [TestMethod]
        public void RecognizeShorterWordIfLongestNotThere()
        {
            TestTerminalMatching("ACOS(4)", false, new[] { "ACOS", "ACOSH" }, t => t == "ACOS");
        }

        [TestMethod]
        public void CaseSensitiveModeWontMatchWordsDifferingInCase()
        {
            TestTerminalNotMatching("FUNC", true, new[] { "func" });
            TestTerminalNotMatching("func", true, new[] { "FUNC" });
        }

        [TestMethod]
        public void CaseInsensitiveModeWillMatchWordsDifferingInCase()
        {
            TestTerminalMatching("FUNC", false, new[] { "func" });
            TestTerminalMatching("func", false, new[] { "FUNC" });
        }

        private static void TestTerminalMatching(string input, bool caseSensitive, string[] words, Func<string, bool> checkTerm = null)
        {
            var grammar = new TestGrammar(caseSensitive, new WordsTerminal("words", words));
            var parser = new Parser(grammar);
            var parseTree = parser.Parse(input);
            Assert.AreEqual(ParseTreeStatus.Parsed, parseTree.Status);
            if (checkTerm != null)
                Assert.IsTrue(checkTerm(parseTree.Tokens[0].Text));
        }

        private static void TestTerminalNotMatching(string input, bool caseSensitive, string[] words)
        {
            var grammar = new TestGrammar(caseSensitive, new WordsTerminal("words", words));
            var parser = new Parser(grammar);
            var parseTree = parser.Parse(input);
            Assert.AreEqual(ParseTreeStatus.Error, parseTree.Status);
        }

        private class TestGrammar : Grammar
        {
            public TestGrammar(bool caseSensitive, Terminal testedTerminal) : base(caseSensitive)
            {
                Root = new NonTerminal("start")
                {
                    Rule = testedTerminal + new RegexBasedTerminal("remainder", ".*")
                        | testedTerminal + Eof
                };
            }
        }
    }
}
