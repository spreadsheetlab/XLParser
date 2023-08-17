using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Web;
using Irony.Parsing;
using Newtonsoft.Json;
using XLParser.Web.XLParserVersions.v100;

namespace XLParser.Web
{
    public class Parse : IHttpHandler
    {
        private HttpContext _httpContext;

        private void WriteResponse(string s)
        {
            _httpContext.Response.Write(s);
        }

        private static readonly bool DisableCache =
#if(DEBUG)
            true;
#else
            false;
#endif

        private const string LatestVersion = "170";

        public void ProcessRequest(HttpContext context)
        {
            _httpContext = context;

            if (!DisableCache && context.Request.Params["nocache"] != "true")
            {
                context.Response.Cache.SetCacheability(HttpCacheability.Public);
                context.Response.Cache.SetExpires(DateTime.Now.AddMinutes(5));
                context.Response.Cache.SetMaxAge(new TimeSpan(0, 0, 5));
            }

            context.Response.AddHeader("Access-Control-Allow-Origin", "*");

            // Dynamically load a library version
            var xlParserVersion = context.Request.Params["version"] ?? LatestVersion;
            if (!Regex.IsMatch(xlParserVersion, @"^[0-9]{3,4}[\-a-z0-9]*$"))
            {
                context.Response.StatusCode = (int) HttpStatusCode.BadRequest;
                _httpContext.Response.ContentType = "text/plain";
                WriteResponse("Invalid version");
                context.Response.End();
                return;
            }

            try
            {
                LoadXlParserVersion(xlParserVersion);
            }
            catch (ArgumentException)
            {
                context.Response.StatusCode = (int) HttpStatusCode.NotFound;
                _httpContext.Response.ContentType = "text/plain";
                WriteResponse("Version doesn't exist");
                context.Response.End();
                return;
            }

            // We want to actually give meaningful HTTP error codes and not have IIS interfere
            context.Response.TrySkipIisCustomErrors = true;

            // check file extension for format
            var format = (Path.GetExtension(context.Request.FilePath) ?? ".json").TrimStart('.');
            var formula = context.Request.Unvalidated["formula"];
            switch (format)
            {
                case "json":
                    ParseToJson(formula);
                    break;
                default:
                    context.Response.StatusCode = 415;
                    _httpContext.Response.ContentType = "text/plain";
                    WriteResponse($"Format '{format}' not supported.");
                    context.Response.End();
                    break;
            }
        }

        private void ParseToJson(string formula)
        {
            _httpContext.Response.ContentType = "application/json";
            if (formula == null)
            {
                _httpContext.Response.StatusCode = 400;
                WriteResponse(JsonConvert.SerializeObject(new {error = "no formula supplied"}));
                _httpContext.Response.End();
                return;
            }

            ParseTreeNode root;
            try
            {
                //root = XLParser.ExcelFormulaParser.Parse(formula);
                root = _parse(formula);
            }
            catch (ArgumentException)
            {
                // Parse error, return 422 - Unprocessable Entity
                _httpContext.Response.StatusCode = 422;
                ParseTree r = new Parser((Grammar) Activator.CreateInstance(_grammar)).Parse(formula);
                WriteResponse(JsonConvert.SerializeObject(new
                {
                    error = "Parse error",
                    formula,
                    message = r.ParserMessages.Select(m => new
                    {
                        level = m.Level.ToString(),
                        line = m.Location.Line + 1,
                        column = m.Location.Column + 1,
                        msg = m.Message
                    }).FirstOrDefault()
                }));
                _httpContext.Response.End();
                return;
            }

            WriteResponse(JsonConvert.SerializeObject(ToJson(root), Formatting.Indented, new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore
            }));
            _httpContext.Response.End();
        }

        private JsonNode ToJson(ParseTreeNode node)
        {
            return new JsonNode
            {
                name = NodeText(node),
                children = node.ChildNodes.Count == 0 ? null : node.ChildNodes.Select(ToJson)
            };
        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private class JsonNode
        {
            public IEnumerable<JsonNode> children;
            public string name;
        }

        private string NodeText(ParseTreeNode node)
        {
            if (node.Term is NonTerminal)
            {
                return node.Term.Name;
            }

            // These are simple terminals like + or =, just print them
            // For other terminals, print the terminal name + contents
            return node.Term.Name.Length <= 2 ? _print(node) : $"{node.Term.Name}[\"{_print(node)}\"]";
        }

        private Func<string, ParseTreeNode> _parse;
        private Func<ParseTreeNode, string> _print;
        private Type _grammar;

        // Yes, this is f-ugly. Better solutions were tried (dynamically loading through reflection, extern alias and separate AppDomains) but failed.
        // Mainly this is because .NET is very very picky about loading multiple versions of libraries with the same name
        private void LoadXlParserVersion(string version)
        {
            switch (version)
            {
                case "100":
                    _parse = ExcelFormulaParser.Parse;
                    _print = ExcelFormulaParser.Print;
                    _grammar = typeof(ExcelFormulaGrammar);
                    break;
                case "114":
                    _parse = XLParserVersions.v114.ExcelFormulaParser.Parse;
                    _print = XLParserVersions.v114.ExcelFormulaParser.Print;
                    _grammar = typeof(XLParserVersions.v114.ExcelFormulaGrammar);
                    break;
                case "120":
                    _parse = XLParserVersions.v120.ExcelFormulaParser.Parse;
                    _print = XLParserVersions.v120.ExcelFormulaParser.Print;
                    _grammar = typeof(XLParserVersions.v120.ExcelFormulaGrammar);
                    break;
                case "139":
                    _parse = XLParserVersions.v139.ExcelFormulaParser.Parse;
                    _print = XLParserVersions.v139.ExcelFormulaParser.Print;
                    _grammar = typeof(XLParserVersions.v139.ExcelFormulaGrammar);
                    break;
                case "1310":
                    _parse = XLParserVersions.v1310.ExcelFormulaParser.Parse;
                    _print = XLParserVersions.v1310.ExcelFormulaParser.Print;
                    _grammar = typeof(XLParserVersions.v1310.ExcelFormulaGrammar);
                    break;
                case "141":
                    _parse = XLParserVersions.v141.ExcelFormulaParser.Parse;
                    _print = XLParserVersions.v141.ExcelFormulaParser.Print;
                    _grammar = typeof(XLParserVersions.v141.ExcelFormulaGrammar);
                    break;
                case "142":
                    _parse = XLParserVersions.v142.ExcelFormulaParser.Parse;
                    _print = XLParserVersions.v142.ExcelFormulaParser.Print;
                    _grammar = typeof(XLParserVersions.v142.ExcelFormulaGrammar);
                    break;
                case "150":
                    _parse = XLParserVersions.v150.ExcelFormulaParser.Parse;
                    _print = XLParserVersions.v150.ExcelFormulaParser.Print;
                    _grammar = typeof(XLParserVersions.v150.ExcelFormulaGrammar);
                    break;
                case "151":
                    _parse = XLParserVersions.v151.ExcelFormulaParser.Parse;
                    _print = XLParserVersions.v151.ExcelFormulaParser.Print;
                    _grammar = typeof(XLParserVersions.v151.ExcelFormulaGrammar);
                    break;
                case "152":
                    _parse = XLParserVersions.v152.ExcelFormulaParser.Parse;
                    _print = XLParserVersions.v152.ExcelFormulaParser.Print;
                    _grammar = typeof(XLParserVersions.v152.ExcelFormulaGrammar);
                    break;
                case "160":
                    _parse = XLParserVersions.v160.ExcelFormulaParser.Parse;
                    _print = XLParserVersions.v160.ExcelFormulaParser.Print;
                    _grammar = typeof(XLParserVersions.v160.ExcelFormulaGrammar);
                    break;
                case "161":
                    _parse = XLParserVersions.v161.ExcelFormulaParser.Parse;
                    _print = XLParserVersions.v161.ExcelFormulaParser.Print;
                    _grammar = typeof(XLParserVersions.v161.ExcelFormulaGrammar);
                    break;
                case "162":
                    _parse = XLParserVersions.v162.ExcelFormulaParser.Parse;
                    _print = XLParserVersions.v162.ExcelFormulaParser.Print;
                    _grammar = typeof(XLParserVersions.v162.ExcelFormulaGrammar);
                    break;
                case "163":
                    _parse = XLParserVersions.v163.ExcelFormulaParser.Parse;
                    _print = XLParserVersions.v163.ExcelFormulaParser.Print;
                    _grammar = typeof(XLParserVersions.v163.ExcelFormulaGrammar);
                    break;
                case "170":
                    _parse = XLParserVersions.v170.ExcelFormulaParser.Parse;
                    _print = XLParserVersions.v170.ExcelFormulaParser.Print;
                    _grammar = typeof(XLParserVersions.v170.ExcelFormulaGrammar);
                    break;
                default:
                    throw new ArgumentException($"Version {version} doesn't exist");
            }
        }

        public bool IsReusable => true;
    }
}