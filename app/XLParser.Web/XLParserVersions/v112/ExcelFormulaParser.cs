using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Irony.Parsing;

namespace XLParser.Web.XLParserVersions.v112
{
    /// <summary>
    /// Excel formula parser <br/>
    /// Contains parser and utilities that operate directly on the parse tree, or makes working with the parse tree easier.
    /// </summary>
public static class ExcelFormulaParser
    {
        /// <summary>
        /// Singleton parser instance
        /// </summary>
        private readonly static Parser p = new Parser(new ExcelFormulaGrammar());

        /// <summary>
        /// Parse a formula, return the the tree's root node
        /// </summary>
        /// <param name="input">The formula to be parsed.</param>
        /// <exception cref="ArgumentException">
        /// If formula could not be parsed
        /// </exception>
        /// <returns>Parse tree root node</returns>
        public static ParseTreeNode Parse(string input)
        {
            return ParseToTree(input).Root;
        }

        /// <summary>
        /// Parse a formula, return the the tree
        /// </summary>
        /// <param name="input">The formula to be parsed.</param>
        /// <exception cref="ArgumentException">
        /// If formula could not be parsed
        /// </exception>
        /// <returns>Parse tree</returns>
        public static ParseTree ParseToTree(string input)
        {
            var tree = p.Parse(input);

            if (tree.HasErrors())
            {
                throw new ArgumentException("Failed parsing input <<" + input + ">>");
            }

            return tree;
        }

        /// <summary>
        /// All non-terminal nodes in depth-first pre-order
        /// </summary>
        // inspiration taken from https://irony.codeplex.com/discussions/213938
        public static IEnumerable<ParseTreeNode> AllNodes(this ParseTreeNode root)
        {
            var stack = new Stack<ParseTreeNode>();
            stack.Push(root);

            while (stack.Count > 0)
            {
                var node = stack.Pop();
                yield return node;

                var children = node.ChildNodes;
                // Push children on in reverse order so that they will
                // be evaluated left -> right when popped.
                for (int i = children.Count - 1; i >= 0; i--)
                {
                    stack.Push(children[i]);
                }
            }
        }

        /// <summary>
        /// All non-terminal nodes of a certain type in depth-first pre-order
        /// </summary>
        public static IEnumerable<ParseTreeNode> AllNodes(this ParseTreeNode root, string type)
        {
            return AllNodes(root.AllNodes(), type);
        }

        internal static IEnumerable<ParseTreeNode> AllNodes(IEnumerable<ParseTreeNode> allNodes, string type)
        {
            return allNodes.Where(node => node.Is(type));
        }

        /// <summary>
        /// Whether this tree contains any nodes of a type
        /// </summary>
        public static bool Contains(this ParseTreeNode root, string type)
        {
            return root.AllNodes(type).Any();
        }

        /// <summary>
        /// The node type/name
        /// </summary>
        public static string Type(this ParseTreeNode node)
        {
            return node.Term.Name;
        }

        /// <summary>
        /// Check if a node is of a particular type
        /// </summary>
        public static bool Is(this ParseTreeNode pt, string type)
        {
            return pt.Type() == type;
        }

        /// <summary>
        /// Checks whether this node is a function
        /// </summary>
        public static Boolean IsFunction(this ParseTreeNode input)
        {
            return input.Is(GrammarNames.FunctionCall)
                || input.Is(GrammarNames.ReferenceFunctionCall)
                || input.Is(GrammarNames.UDFunctionCall)
                // This gives potential problems/duplication on external UDF's, but they are so rare that I think this is acceptable
                || (input.Is(GrammarNames.Reference) && input.ChildNodes.Count == 2 && input.ChildNodes[1].IsFunction())
                ;
        }

        /// <summary>
        /// Whether or not this node represents parentheses "(_)"
        /// </summary>
        public static bool IsParentheses(this ParseTreeNode input)
        {
            switch (input.Type())
            {
                case GrammarNames.Formula:
                    return input.ChildNodes.Count == 1 && input.ChildNodes[0].Is(GrammarNames.Formula);
                case GrammarNames.Reference:
                    return input.ChildNodes.Count == 1 && input.ChildNodes[0].Is(GrammarNames.Reference);
                default:
                    return false;
            }
        }

        public static bool IsBinaryOperation(this ParseTreeNode input)
        {
            return input.IsFunction()
                   && input.ChildNodes.Count() == 3
                   && input.ChildNodes[1].Term.Flags.HasFlag(TermFlags.IsOperator);
        }

        public static bool IsBinaryReferenceOperation(this ParseTreeNode input)
        {
            return input.IsBinaryOperation() && input.Is(GrammarNames.ReferenceFunctionCall);
        }

        public static bool IsUnaryOperation(this ParseTreeNode input)
        {
            return IsUnaryPrefixOperation(input) || IsUnaryPostfixOperation(input);
        }

        public static bool IsUnaryPrefixOperation(this ParseTreeNode input)
        {
            return input.IsFunction()
                   && input.ChildNodes.Count() == 2
                   && input.ChildNodes[0].Term.Flags.HasFlag(TermFlags.IsOperator);
        }

        public static bool IsUnaryPostfixOperation(this ParseTreeNode input)
        {
            return input.IsFunction()
                   && input.ChildNodes.Count() == 2
                   && input.ChildNodes[1].Term.Flags.HasFlag(TermFlags.IsOperator);

        }

        private static string RemoveFinalSymbol(string input)
        {
            input = input.Substring(0, input.Length - 1);
            return input;
        }

        /// <summary>
        /// Get the function or operator name of this function call
        /// </summary>
        public static string GetFunction(this ParseTreeNode input)
        {
            if (input.IsIntersection())
            {
                return GrammarNames.TokenIntersect;
            }
            if (input.IsUnion())
            {
                return GrammarNames.TokenUnionOperator;
            }
            if (input.IsBinaryOperation() || input.IsUnaryPostfixOperation())
            {
                return input.ChildNodes[1].Print();
            }
            if (input.IsUnaryPrefixOperation())
            {
                return input.ChildNodes[0].Print();
            }
            if (input.IsNamedFunction())
            {
                return RemoveFinalSymbol(input.ChildNodes[0].Print()).ToUpper();
            }
            if (input.IsExternalUDFunction())
            {
                return String.Format("{0}{1}", input.ChildNodes[0].Print(), GetFunction(input.ChildNodes[1]));
            }

            throw new ArgumentException("Not a function call", "input");
        }

        /// <summary>
        /// Check if this node is a specific function
        /// </summary>
        public static bool MatchFunction(this ParseTreeNode input, String functionName)
        {
            return IsFunction(input) && GetFunction(input) == functionName;
        }

        /// <summary>
        /// Checks whether this node is a built-in excel function
        /// </summary>
        public static bool IsBuiltinFunction(this ParseTreeNode node)
        {
            return node.IsFunction() && (node.ChildNodes[0].Is(GrammarNames.ExcelFunction) || node.ChildNodes[0].Is(GrammarNames.RefFunctionName));
        }

        /// <summary>
        /// Whether or not this node represents an intersection
        /// </summary>
        public static bool IsIntersection(this ParseTreeNode input)
        {
            try
            {
                return IsBinaryOperation(input) &&
                       input.ChildNodes[1].Token.Terminal.Name == GrammarNames.TokenIntersect;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Whether or not this node represents an union
        /// </summary>
        public static bool IsUnion(this ParseTreeNode input)
        {
            try
            {
                return input.Is(GrammarNames.ReferenceFunctionCall) && input.ChildNodes[0].Is(GrammarNames.Union);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Checks whether this node is a function call with name, and not just a unary or binary operation
        /// </summary>
        public static bool IsNamedFunction(this ParseTreeNode input)
        {
            return (input.Is(GrammarNames.FunctionCall) && input.ChildNodes[0].Is(GrammarNames.FunctionName))
                || (input.Is(GrammarNames.ReferenceFunctionCall) && input.ChildNodes[0].Is(GrammarNames.RefFunctionName))
                || input.Is(GrammarNames.UDFunctionCall);
        }

        public static bool IsExternalUDFunction(this ParseTreeNode input)
        {
            return input.Is(GrammarNames.Reference) && input.ChildNodes.Count == 2 && input.ChildNodes[1].IsNamedFunction();
        }

        /// <summary>
        /// True if this node presents a number constant with a sign
        /// </summary>
        public static bool IsNumberWithSign(this ParseTreeNode input)
        {
            return IsUnaryPrefixOperation(input)
                   && input.ChildNodes[1].ChildNodes[0].Is(GrammarNames.Constant)
                   && input.ChildNodes[1].ChildNodes[0].ChildNodes[0].Is(GrammarNames.Number);
        }

        /// <summary>
        /// Go to the first non-formula child node
        /// </summary>
        public static ParseTreeNode SkipFormula(this ParseTreeNode input)
        {
            while (input.Is(GrammarNames.Formula))
            {
                input = input.ChildNodes.First();
            }
            return input;
        }

        /// <summary>
        /// Go to the first "relevant" child node, i.e. skips wrapper nodes
        /// </summary>
        /// <remarks>
        /// Skips:
        /// * FormulaWithEq and ArrayFormula nodes
        /// * Formula nodes
        /// * Parentheses
        /// * Reference nodes which are just wrappers
        /// </remarks>
        public static ParseTreeNode SkipToRelevant(this ParseTreeNode input)
        {
            switch (input.Type())
            {
                case GrammarNames.FormulaWithEq:
                case GrammarNames.ArrayFormula:
                    return SkipToRelevant(input.ChildNodes[1]);
                case GrammarNames.Formula:
                case GrammarNames.Reference:
                    // This also catches parentheses
                    if (input.ChildNodes.Count == 1)
                    {
                        return SkipToRelevant(input.ChildNodes[0]);
                    }
                    goto default;
                default:
                    return input;
            }
        }

        /// <summary>
        /// Pretty-print a parse tree to a string
        /// </summary>
        public static string Print(this ParseTreeNode input)
        {
            // For terminals, just print the token text
            if (input.Term is Terminal)
            {
                return input.Token.Text;
            }

            // (Lazy) enumerable for printed childs
            var childs = input.ChildNodes.Select(Print);
            // Concrete list when needed
            List<String> childsL;

            // Switch on nonterminals
            switch (input.Term.Name)
            {
                case GrammarNames.Formula:
                    // Check if these are brackets, otherwise print first child
                    return IsParentheses(input) ? String.Format("({0})", childs.First()) : childs.First();

                case GrammarNames.FunctionCall:
                case GrammarNames.ReferenceFunctionCall:
                case GrammarNames.UDFunctionCall:
                    childsL = childs.ToList();

                    if (input.IsNamedFunction())
                    {
                        return String.Join("", childsL) + ")";
                    }

                    if (input.IsBinaryOperation())
                    {
                        // format string for "normal" binary operation
                        string format = "{0} {1} {2}";
                        if (input.IsIntersection())
                        {
                            format = "{0} {2}";
                        }else if (input.IsBinaryReferenceOperation())
                        {
                            format = "{0}{1}{2}";
                        }
                        
                        return String.Format(format, childsL[0], childsL[1], childsL[2]);
                    }

                    if (input.IsUnion())
                    {
                        return String.Format("({0})", String.Join(",", childsL));
                    }

                    if (input.IsUnaryOperation())
                    {
                        return String.Join("", childsL);
                    }

                    throw new ArgumentException("Unknown function type.");

                case GrammarNames.Reference:
                    /*if (IsParentheses(input) || IsUnion(input))
                    {
                        return String.Format("({0})", childs.First());
                    }

                    childsL = childs.ToList();
                    if (IsIntersection(input))
                    {
                        return String.Format("{0} {1}", childsL[0], childsL[2]);
                    }

                    if (IsBinaryOperation(input))
                    {
                        return String.Format("{0}{1}{2}", childsL[0], childsL[1], childsL[2]);
                    }*/
                    if (IsParentheses(input))
                    {
                        return String.Format("({0})", childs.First());
                    }

                    return String.Join("", childs);

                case GrammarNames.File:
                    return String.Format("[{0}]", childs.First());

                case GrammarNames.Prefix:
                    var ret = String.Join("", childs);
                    // The exclamation mark token is not included in the parse tree, so we have to add that if it's a single file
                    if (input.ChildNodes.Count == 1 && input.ChildNodes[0].Is(GrammarNames.File))
                    {
                        ret += "!";
                    }
                    return ret;

                case GrammarNames.ArrayFormula:
                    return "{=" + childs.ElementAt(1) + "}";

                case GrammarNames.DynamicDataExchange:
                    childsL = childs.ToList();
                    return String.Format("{0}!{1}", childsL[0], childsL[1]);

                // Terms for which to print all child nodes concatenated
                case GrammarNames.ArrayConstant:
                case GrammarNames.FormulaWithEq:
                    return String.Join("", childs);

                // Terms for which we print the childs comma-separated
                case GrammarNames.Arguments:
                case GrammarNames.ArrayRows:
                case GrammarNames.Union:
                    return String.Join(",", childs);

                case GrammarNames.ArrayColumns:
                    return String.Join(";", childs);

                case GrammarNames.ConstantArray:
                    return String.Format("{{{0}}}", childs.First());


                default:
                    // If it is not defined above and the number of childs is exactly one, we want to just print the first child
                    if (input.ChildNodes.Count == 1)
                    {
                        return childs.First();
                    }
                    throw new ArgumentException(String.Format("Could not print node of type '{0}'.\nThis probably means the excel grammar was modified without the print function being modified", input.Term.Name));
            }
        }
    }

}

