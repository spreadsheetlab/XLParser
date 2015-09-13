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

        private const string latestVersion = "120";

        public void ProcessRequest(HttpContext context)
        {
            ctx = context;

            if (!disableCache && context.Request.Params["nocache"] != "true")
            {
                context.Response.Cache.SetCacheability(HttpCacheability.Public);
                context.Response.Cache.SetExpires(DateTime.Now.AddMinutes(5));
                context.Response.Cache.SetMaxAge(new TimeSpan(0, 0, 5));
            }

             // Dynamically load a library version
            var xlparserVersion = context.Request.Params["version"] ?? latestVersion;
            if (!Regex.IsMatch(xlparserVersion, @"^[0-9]{3}[\-a-z0-9]*$"))
            {
                context.Response.StatusCode = (int) HttpStatusCode.BadRequest;
                ctx.Response.ContentType = "text/plain";
                w("Invalid version");
                context.Response.End();
                return;
            }

            try
            {
                LoadXLParserVersion(xlparserVersion);
            }
            catch (ArgumentException)
            {
                context.Response.StatusCode = (int) HttpStatusCode.NotFound;
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
                    ParseToJSON(formula);
                    break;
                default:
                    context.Response.StatusCode = 415;
                    ctx.Response.ContentType = "text/plain";
                    w($"Format '{format}' not supported.");
                    context.Response.End();
                    break;
            }
        }

        private void ParseToJSON(string formula)
        {
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
                //root = XLParser.ExcelFormulaParser.Parse(formula);
                root = parse(formula);
            }
            catch (ArgumentException)
            {
                // Parse error, return 422 - Unprocessable Entity
                ctx.Response.StatusCode = 422;
                var r = new Parser((Grammar)Activator.CreateInstance(grammar)).Parse(formula);
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

            w(JsonConvert.SerializeObject(
                    ToJSON(root),
                    Formatting.Indented,
                    new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore,
                    }
            ));
            ctx.Response.End();
        }

        private JSONNode ToJSON(ParseTreeNode node)
        {
            return new JSONNode
            {
                name = NodeText(node),
                children = node.ChildNodes.Count == 0 ? null : node.ChildNodes.Select(ToJSON)
            };
        }

        private class JSONNode
        {
            public string name;
            public IEnumerable<JSONNode> children;
        }

        private string NodeText(ParseTreeNode node)
        {
            if (node.Term is NonTerminal) return node.Term.Name;

            // These are simple terminals like + or =, just print them
            // For other terminals, print the terminal name + contents
            return node.Term.Name.Length <= 2 ? print(node) : $"{node.Term.Name}[\"{print(node)}\"]";
        }

        private Func<string, ParseTreeNode> parse;
        private Func<ParseTreeNode, string> print;
        private Type grammar;

        // Yes, this is f-ugly. Better solutions were tried (dynamically loading through reflection, extern alias and separate appdomains) but failed.
        // Mainly this is because .NET is very very picky about loading multiple versions of libraries with the same name
        private void LoadXLParserVersion(string version)
        {
            switch (version)
            {
                case "100":
                    parse = XLParserVersions.v100.ExcelFormulaParser.Parse;
                    print = XLParserVersions.v100.ExcelFormulaParser.Print;
                    grammar = typeof(XLParserVersions.v100.ExcelFormulaGrammar);
                    break;
                case "112":
                    parse = XLParserVersions.v112.ExcelFormulaParser.Parse;
                    print = XLParserVersions.v112.ExcelFormulaParser.Print;
                    grammar = typeof(XLParserVersions.v112.ExcelFormulaGrammar);
                    break;
                case "113":
                    parse = XLParserVersions.v113.ExcelFormulaParser.Parse;
                    print = XLParserVersions.v113.ExcelFormulaParser.Print;
                    grammar = typeof(XLParserVersions.v113.ExcelFormulaGrammar);
                    break;
                case "114":
                    parse = XLParserVersions.v114.ExcelFormulaParser.Parse;
                    print = XLParserVersions.v114.ExcelFormulaParser.Print;
                    grammar = typeof(XLParserVersions.v114.ExcelFormulaGrammar);
                    break;
                case "120":
                    parse = XLParserVersions.v120.ExcelFormulaParser.Parse;
                    print = XLParserVersions.v120.ExcelFormulaParser.Print;
                    grammar = typeof(XLParserVersions.v120.ExcelFormulaGrammar);
                    break;
                default:
                    throw new ArgumentException($"Version {version} doesn't exist");
            }
        }

        public bool IsReusable => true;
    }
}
