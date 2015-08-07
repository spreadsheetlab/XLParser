<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="XLParser.Web.Default" %>
<!DOCTYPE html>
<html>
<head runat="server">
    <meta charset="utf-8"/>
    <title>XLParser web demo</title>
    <link rel="stylesheet" href="xlparser-web.css" type="text/css" />
    <script src="http://d3js.org/d3.v3.min.js" charset="utf-8"></script>
    <!--<script src="http://d3js.org/d3.v3.js" charset="utf-8"></script>-->
    <script src="d3vizsvg.js" type="text/javascript" defer></script>
</head>
<body>
    <div id="borderwrapper">
        <h2>XLParser web demo</h2>
    
        <p>
            Formula: <input type="text" size="100" id="formulainput" /> <br/> <br />
            <button onclick="newTree(document.getElementById('formulainput').value)">Parse</button>
        </p>
    
        <div id="bugreport">
            <a href="javascript:;" onclick=" javascript: if(document.getElementById('bug_explanation').style.display !== 'block') {document.getElementById('bug_explanation').style.display = 'block'} else {document.getElementById('bug_explanation').style.display = 'none'};" style="color: black; font-weight: bold;">Found a bug?</a> <br/> <br/>
            <div id="bug_explanation" style="display: none;">
                Great! <a href="https://github.com/PerfectXL/XLParser/issues">Please report it as a Github issue!</a> <br/> <br/>

                If the bug is with a specific formula/excel file, please include that too.<br/>
                Generally bugs in XLParser are one of the following, please include this type in the report:<br/>
                <ul>
                    <li>The parser can't parse a formula that Excel accepts</li>
                    <li>The parser parses a formula that Excel doesn't accept.</li>
                    <li>The parser interprets a formula wrong, that is it produces a parse tree that doesn't correspond with how Excel behaves.</li>
                    <li>There is a "normal" bug in the code around the core parser.</li>
                </ul>
            
            </div>
        </div>
    </div>
    <div id="borderwrapper2">
        <!-- Based on https://mbostock.github.io/d3/talk/20111018/tree.html and https://gist.github.com/d3noob/8326869-->
        <p>Parse Tree (<a id="imgdatasvg">SVG</a>, <a id="imgdatapng">PNG</a>):</p>
        <div id="d3viz">
        </div>
    </div>
    <div id="imgdata"></div>
</body>
</html>
