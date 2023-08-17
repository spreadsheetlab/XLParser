<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="XLParser.Web.Default" %>
<!DOCTYPE html>
<html>
<head runat="server">
    <meta charset="utf-8"/>
    <title>XLParser web demo</title>
    <link rel="stylesheet" href="xlparser-web.css" type="text/css"/>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/d3/3.5.6/d3.min.js" charset="utf-8"></script>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.4/jquery.min.js"></script>
    <script src="d3vizsvg.js" type="text/javascript" defer>
    </script>
</head>
<body>
<div id="borderwrapper">
    <div id="leftoflogos">
        <h2><a href="https://github.com/PerfectXL/XLParser">XLParser</a> web demo</h2>

        <table>
            <colgroup>
                <col/>
                <col style="width: 100px;"/>
                <col style="width: 100px;"/>
            </colgroup>
            <thead>
            <tr>
                <th style="text-align: left;">Formula:</th>
                <th style="text-align: left;">Version:</th><th></th>
            </tr>
            </thead>
            <tbody>
            <tr>
                <td>
                    <textarea rows="1" id="formulainput"></textarea>
                </td>
                <td>
                    <select id="parserversionselected">
                        <option value="170" selected>1.7.0</option>
                        <option value="163">1.6.3</option>
                        <option value="162">1.6.2</option>
                        <option value="161">1.6.1</option>
                        <option value="160">1.6.0</option>
                        <option value="152">1.5.2</option>
                        <option value="151">1.5.1</option>
                        <option value="150">1.5.0</option>
                        <option value="142">1.4.2</option>
                        <option value="141">1.4.1</option>
                        <option value="1310">1.3.10</option>
                        <option value="139">1.3.9</option>
                        <option value="120">1.2.0</option>
                        <option value="114">1.1.4</option>
                        <option value="100">1.0.0</option>
                    </select>
                </td>
                <td>
                    <button id="parsebutton" onclick="newTree(document.getElementById('formulainput').value, document.getElementById('parserversionselected').value)">Parse</button>
                </td>
            </tr>
            </tbody>
        </table>

        <div id="bugreport">
            <a href="javascript:;" onclick=" javascript: if (document.getElementById('bug_explanation').style.display !== 'block') { document.getElementById('bug_explanation').style.display = 'block' } else { document.getElementById('bug_explanation').style.display = 'none' };" style="color: black; font-weight: bold;">Found a bug?</a> <br/> <br/>
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

    <div id="logos">
        <a href="https://tudelft.nl">
            <img src="img/logo-tudelft.png" alt="TU Delft logo"/>
        </a><br/>
        <img src="img/logo-spreadsheet-lab.png" alt="Spreadsheet lab logo"/><br/>
        <a href="https://www.infotron.nl/">
            <img src="img/logo-infotron.png" alt="Infotron logo"/>
        </a>
    </div>
</div>
<div id="borderwrapper2">
    <script type="text/javascript">
        var ua = window.navigator.userAgent;
        if (ua.indexOf("MSIE ") > 0 || ua.indexOf('Trident/') > 0) {
            document.write(
                "<em>Note: image downloading does not work properly in Internet Explorer 11 and lower.</em><br />");
        }
    </script>
    <!-- Based on https://mbostock.github.io/d3/talk/20111018/tree.html and https://gist.github.com/d3noob/8326869-->
    <p>Parse Tree (<a id="imgdatasvg">SVG</a>, <a id="imgdatapng">PNG</a>):</p>
    <div id="d3viz"></div>
</div>
</body>
</html>