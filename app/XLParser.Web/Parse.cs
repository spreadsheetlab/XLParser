using System;
using System.Runtime.InteropServices;
using System.Web;
using Irony.Parsing;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Linq;

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

        public void ProcessRequest(HttpContext context)
        {
            ctx = context;

            if (!disableCache && context.Request.Params["nocache"] != "true")
            {
                context.Response.Cache.SetCacheability(HttpCacheability.Public);
                context.Response.Cache.SetExpires(DateTime.Now.AddMinutes(5));
                context.Response.Cache.SetMaxAge(new TimeSpan(0, 0, 5));
            }

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
                    w(String.Format("Format '{0}' not supported.", format));
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
                root = ExcelFormulaParser.Parse(formula);
            }
            catch (ArgumentException)
            {
                // Parse error, return 422 - Unprocessable Entity
                ctx.Response.StatusCode = 422;
                var r = new Parser(new ExcelFormulaGrammar()).Parse(formula);
                w(JsonConvert.SerializeObject(new
                {
                    error = "Parse error",
                    messages = r.ParserMessages.Select(m => String.Format("{0} at {1}: {2}", m.Level, m.Location, m.Message)),
                    formula = formula
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

        private static JSONNode ToJSON(ParseTreeNode node)
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

        private static string NodeText(ParseTreeNode node)
        {
            if (node.Term is NonTerminal) return node.Type();

            // These are simple terminals like + or =, just print them
            if (node.Type().Length <= 2) return node.Print();

            // For other terminals, print the terminal name + contents
            return String.Format("{0}[\"{1}\"]", node.Type(), node.Print());
        }

        public bool IsReusable
        {
            // Return false in case your Managed Handler cannot be reused for another request.
            // Usually this would be false in case you have some state information preserved per request.
            get { return true; }
        }
    }
}
