using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Irony.Parsing;
using System.Collections;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;

namespace XLParser
{
    /// <summary>
    /// This class can do some simple analysis on the trees produced by the parser.
    /// </summary>
    /// <remarks>
    /// To prevent bloating this class, please make a new (sub)class and file when you want to add a coherent set of other analyses.
    /// </remarks>
    public class FormulaAnalyzer
    {
        public ParseTreeNode Root { get; private set; }

        private List<ParseTreeNode> allNodes;

        /// <summary>
        /// Lazy cached version of all nodes
        /// </summary>
        public List<ParseTreeNode> AllNodes
        {
            get
            {
                return allNodes ?? (allNodes = Root.AllNodes().ToList());
            }
        } 

        /// <summary>
        /// Provide formula analysis functions on a tree
        /// </summary>
        public FormulaAnalyzer(ParseTreeNode root)
        {
            Root = root;
        }

        /// <summary>
        /// Provide formula analysis functions
        /// </summary>
        public FormulaAnalyzer(string formula) : this(ExcelFormulaParser.Parse(formula))
        {}

        /// <summary>
        /// Get all references that aren't part of another reference expression
        /// </summary>
        public IEnumerable<ParseTreeNode> References()
        {
            return Root.GetReferenceNodes();
        }

        public IEnumerable<string> Functions()
        {
            return AllNodes
                .Where(node => node.IsFunction())
                .Select(ExcelFormulaParser.GetFunction);
        }

        public IEnumerable<string> Constants()
        {
            return Root.AllNodesConditional(ExcelFormulaParser.IsNumberWithSign)
                .Where(node => node.Is(GrammarNames.Constant) || node.IsNumberWithSign())
                .Select(ExcelFormulaParser.Print);
        } 

        ///<summary>
        /// Return all constant numbers used in this formula
        ///</summary>
        public IEnumerable<double> Numbers()
        {
            // Excel numbers can be a double, short or signed int. double can fully represent all of these
            return Root.AllNodesConditional(ExcelFormulaParser.IsNumberWithSign)
                .Where(node => node.Is(GrammarNames.Number) || node.IsNumberWithSign())
                .Select(node => double.Parse(node.Print(), NumberStyles.Float, CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Return the depth of the parse tree, the number of nested Formulas
        /// </summary>
        public int Depth()
        {
            return Depth(Root);
        }

        /// <summary>
        /// Depth of nested formulas
        /// </summary>
        private static int Depth(ParseTreeNode node)
        {
            // Get the maximum depth of the childnodes
            int depth = node.ChildNodes.Count == 0 ? 0 : node.ChildNodes.Max(n => Depth(n));

            // If this is a formula node, add one to the depth
            if (node.Is(GrammarNames.Formula))
            {
                depth++;
            }

            return depth;
        }

        /// <summary>
        /// Get function/operator depth
        /// </summary>
        /// <param name="operators">If not null, count only specific functions/operators</param>
        public int OperatorDepth(ISet<string> operators = null)
        {
            return OperatorDepth(Root, operators);
        }

        private int OperatorDepth(ParseTreeNode node, ISet<string> operators = null)
        {
            // Get the maximum depth of the childnodes
            int depth = node.ChildNodes.Count == 0 ? 0 : node.ChildNodes.Max(n => OperatorDepth(n, operators));

            // If this is one of the target functions, increase depth by 1
            if(node.IsFunction()
               && (operators == null || operators.Contains(node.GetFunction())))
            {
                depth++;
            }

            return depth;
        }

        private static readonly ISet<string> conditionalFunctions = new HashSet<string>()
                {
                    "IF",
                    "COUNTIF",
                    "COUNTIFS",
                    "SUMIF",
                    "SUMIFS",
                    "AVERAGEIF",
                    "AVERAGEIFS",
                    "IFERROR"
                };
        /// <summary>
        /// Get the conditional complexity of the formula
        /// </summary>
        public int ConditionalComplexity()
        {
            return OperatorDepth(Root, conditionalFunctions);
        }
    }
}
