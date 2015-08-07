using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using XLParser;
using Irony.Parsing;

namespace XLParser.Web
{
    public partial class Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            ParseOutputTreeCollapseButton.Click += (s, ie) => { ParseOutputTree.CollapseAll(); };
            ParseOutputTreeExpandButton.Click += (s, ie) => { ParseOutputTree.ExpandAll(); };
            
            ParseOutputTree.Nodes.Clear();
            if (FormulaInput.Text != "")
            {
                string formula = FormulaInput.Text;
                var root = ExcelFormulaParser.Parse(formula);

                var tvroot = convert(root);
                ParseOutputTree.Nodes.Add(tvroot);
                tvroot.Expand();

                // Styling
                ParseOutputTree.LeafNodeStyle.Font.Name = "monospace";
            }
        }

        private static TreeNode convert(ParseTreeNode node)
        {
            var tvnode = new TreeNode();
            tvnode.SelectAction = TreeNodeSelectAction.None;
            tvnode.Text = OutputNode(node);
            tvnode.Expanded = showExpanded.Contains(node.Type());

            foreach(var tvchild in node.ChildNodes.Select(convert))
            {
                tvnode.ChildNodes.Add(tvchild);
            }
            return tvnode;
        }

        private static string OutputNode(ParseTreeNode node)
        {
            if (node.Term is Terminal)
            {
                // These are simple terminals like + or =, just print them
                if (node.Type().Length <= 2)
                {
                    return node.Print();
                }else {
                    // For other terminals, print the terminal name + contents
                    return String.Format("{0}[\"{1}\"]", node.Type(), node.Print());
                }
            }
            else
            {
                // Only print the name of non-terminals
                return node.Type();
            }
        }

        /// Whether or not to Expand a node type by default
        private static readonly ISet<string> showExpanded = new HashSet<string>()
        {
            GrammarNames.Formula,
            GrammarNames.Reference,
            GrammarNames.Arguments,
            GrammarNames.FormulaWithEq,
            GrammarNames.ArrayFormula,
            GrammarNames.FunctionCall,
            GrammarNames.ReferenceFunctionCall,
            GrammarNames.UDFunctionCall
        };
    }
}