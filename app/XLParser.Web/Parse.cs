using System;
using System.Runtime.InteropServices;
using System.Web;
using Irony.Parsing;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Diagnostics;
using System.Net;
using System.Text.RegularExpressions;

namespace XLParser.Web
{
    public class Parse : IHttpHandler
    {
        private HttpContext ctx;

        private void w(string s)
        {
            ctx.Response.Write(s);
        }

        private static readonly bool disableCache =
        #if(DEBUG)
            true;
        #else
            false;
        #endif

        private const string latestVersion = "114";

        public void ProcessRequest(HttpContext context)
        {
            ctx = context;

            if (!disableCache && context.Request.Params["nocache"] != "true")
            {
                context.Response.Cache.SetCacheability(HttpCacheability.Public);
                context.Response.Cache.SetExpires(DateTime.Now.AddMinutes(5));
                context.Response.Cache.SetMaxAge(new TimeSpan(0, 0, 5));
            }

            // Dynamically load an library version
            var xlparserVersion = context.Request.Params["version"] ?? latestVersion;
            if (!Regex.IsMatch(xlparserVersion, @"^[0-9]{3}[\-a-z0-9]*$"))
            {
                context.Response.StatusCode = (int)HttpStatusCode.BadRequest;
                ctx.Response.ContentType = "text/plain";
                w("Invalid version");
                context.Response.End();
                return;
            }

            Assembly xlparser;
            try
            {
                xlparser = LoadXLParserVersion(xlparserVersion);
            }
            catch (ArgumentException)
            {
                context.Response.StatusCode = (int)HttpStatusCode.NotFound;
                ctx.Response.ContentType = "text/plain";
                w("Version doesn't exist");
                context.Response.End();
                return;
            }
            

            // We want to actually give meaningful HTTP error codes and not have IIS interfere
            context.Response.TrySkipIisCustomErrors = true;

            // check file extention for format
            var format = (System.IO.Path.GetExtension(context.Request.FilePath) ?? ".json").Substring(1);
            string formula = context.Request.Unvalidated["formula"];
            switch (format)
            {
                case "json":
                    ParseToJSON(formula, xlparser);
                    break;
                default:
                    context.Response.StatusCode = 415;
                    ctx.Response.ContentType = "text/plain";
                    w($"Format '{format}' not supported.");
                    context.Response.End();
                    break;
            }
        }

        private void ParseToJSON(string formula, Assembly xlparser)
        {
            var ExcelFormulaParser = xlparser.GetType("XLParser.ExcelFormulaParser", true);

            ctx.Response.ContentType = "application/json";
            if (formula == null)
            {
                ctx.Response.StatusCode = 400;
                w(JsonConvert.SerializeObject(new { error = "no formula supplied"}));
                ctx.Response.End();
                return;
            }
            ParseTreeNode root;
            try
            {
                var Parse = ExcelFormulaParser.GetMethod("Parse", BindingFlags.Static | BindingFlags.Public);
                //root = XLParser.ExcelFormulaParser.Parse(formula);
                root = (ParseTreeNode)Parse.Invoke(null, new object[] { formula });
            }
            catch (ArgumentException)
            {
                // Parse error, return 422 - Unprocessable Entity
                ctx.Response.StatusCode = 422;
                //var r = new Parser(new ExcelFormulaGrammar()).Parse(formula);
                var r = new Parser((Grammar)Activator.CreateInstance(xlparser.GetType("ExcelFormulaGrammar"))).Parse(formula);
                w(JsonConvert.SerializeObject(new
                {
                    error = "Parse error",
                    formula = formula,
                    message = r.ParserMessages.Select(m => new
                    {
                        level = m.Level.ToString(),
                        line = m.Location.Line+1,
                        column = m.Location.Column+1,
                        msg = m.Message
                        //String.Format("{0} at {1}: {2}", m.Level, m.Location, m.Message)
                    }).FirstOrDefault()
                    
                }));
                ctx.Response.End();
                return;
            }

            Func<ParseTreeNode, string> printer =
                node =>
                    NodeText(node,
                        (inode => (string) ExcelFormulaParser.GetMethod("Print").Invoke(null, new object[] {inode})));

            w(JsonConvert.SerializeObject(
                    ToJSON(root, printer),
                    Formatting.Indented,
                    new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore,
                    }
            ));
            ctx.Response.End();
        }

        private static JSONNode ToJSON(ParseTreeNode node, Func<ParseTreeNode,string> printer)
        {
            return new JSONNode
            {
                name = printer(node),
                children = node.ChildNodes.Count == 0 ? null : node.ChildNodes.Select(n=>ToJSON(n,printer))
            };
        }

        private class JSONNode
        {
            public string name;
            public IEnumerable<JSONNode> children;
        }

        private static string NodeText(ParseTreeNode node, Func<ParseTreeNode, string> Print)
        {
            if (node.Term is NonTerminal) return node.Term.Name;

            // These are simple terminals like + or =, just print them
            // For other terminals, print the terminal name + contents
            return node.Term.Name.Length <= 2 ? Print(node) : $"{node.Term.Name}[\"{Print(node)}\"]";
        }

        private IDictionary<string, Assembly> xlparsers = new Dictionary<string, Assembly>();
        private Assembly LoadXLParserVersion(string version)
        {
            if (xlparsers.ContainsKey(version)) return xlparsers[version];
            try
            {
                var uri = new UriBuilder(Assembly.GetExecutingAssembly().CodeBase);
                string dir = Path.GetDirectoryName(Uri.UnescapeDataString(uri.Path));
                string path = Path.Combine(dir, $@"xlparser\{version}.dll");
                return xlparsers[version] = Assembly.LoadFrom(path);
            }
            catch (FileNotFoundException e)
            {
                throw new ArgumentException($"Version {version} doesn't exist", e);
            }
        }

        public bool IsReusable => true;
    }
}
