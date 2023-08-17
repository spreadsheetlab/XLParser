using Irony.Parsing;
using System;
using System.Collections.Generic;

namespace XLParser.Web.XLParserVersions.v170
{
    /// <summary>
    /// Terminal that can determine, if there the input contains a one of expected words.
    /// </summary>
    /// <remarks>Children of each node are represented as an array to allow direct indexation. Do not use
    /// for words that have a large difference between low and high character of a token.</remarks>
    public class WordsTerminal : Terminal
    {
        private readonly Node _rootNode;
        private readonly List<string> _words;
        private bool _caseSensitive;

        public WordsTerminal(string name, IEnumerable<string> words) : base(name)
        {
            _rootNode = new Node(0);
            _words = new List<string>(words);
        }

        public override void Init(GrammarData grammarData)
        {
            base.Init(grammarData);
            _caseSensitive = Grammar.CaseSensitive;
            foreach (var word in _words)
            {
                AddWordToTree(_caseSensitive ? word : word.ToUpperInvariant());
            }

            if (EditorInfo == null)
            {
                EditorInfo = new TokenEditorInfo(TokenType.Unknown, TokenColor.Text, TokenTriggers.None);
            }
        }

        public override IList<string> GetFirsts() => _words;

        public override Token TryMatch(ParsingContext context, ISourceStream source)
        {
            var node = _rootNode;
            var input = source.Text;
            for (var i = source.PreviewPosition; i < input.Length; ++i)
            {
                var c = _caseSensitive ? input[i] : char.ToUpperInvariant(input[i]);
                var nextNode = node[c];
                if (nextNode is null)
                {
                    break;
                }

                node = nextNode;
            }

            if (!node.IsTerminal)
            {
                return null;
            }

            source.PreviewPosition += node.Length;
            return source.CreateToken(OutputTerminal);
        }

        private void AddWordToTree(string word)
        {
            var node = _rootNode;
            foreach (var c in word)
            {
                node = node.GetOrAddChild(c);
            }

            node.IsTerminal = true;
        }

        private class Node
        {
            private char _lowChar = '\0';
            private char _highChar = '\0';
            private Node[] _children;

            public Node(int length)
            {
                Length = length;
            }

            public bool IsTerminal { get; set; }

            public int Length { get; }

            public Node this[char c]
            {
                get
                {
                    if (_children is null)
                    {
                        return null;
                    }

                    if (c < _lowChar || c > _highChar)
                    {
                        return null;
                    }

                    return _children[c - _lowChar];
                }
            }

            internal Node GetOrAddChild(char c)
            {
                if (_children is null)
                {
                    var node = new Node(Length + 1);
                    _children = new[] { node };
                    _lowChar = c;
                    _highChar = c;
                    return node;
                }

                var newLowChar = (char)Math.Min(_lowChar, c);
                if (newLowChar != _lowChar)
                {
                    var newChildrenCount = _highChar - newLowChar + 1;
                    Array.Resize(ref _children, newChildrenCount);
                    var ofs = _lowChar - newLowChar;
                    Array.Copy(_children, 0, _children, ofs, newChildrenCount - ofs);
                    Array.Clear(_children, 0, ofs);
                    _lowChar = newLowChar;
                    return _children[0] = new Node(Length + 1);
                }

                var newHighChar = (char)Math.Max(_highChar, c);
                if (newHighChar != _highChar)
                {
                    var newChildrenCount = newHighChar - _lowChar + 1;
                    Array.Resize(ref _children, newChildrenCount);
                    _highChar = newHighChar;
                    return _children[newChildrenCount - 1] = new Node(Length + 1);
                }

                var charIdx = c - _lowChar;
                var child = _children[charIdx];
                if (child is null)
                {
                    return _children[charIdx] = new Node(Length + 1);
                }

                return child;
            }
        }
    }
}
