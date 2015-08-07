<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="XLParser.Web.Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>XLParser web demo</title>
    <link rel="stylesheet" href="xlparser-web.css" type="text/css" />
</head>
<body>
    <h2>XLParser web demo</h2>
    <form id="form1" runat="server">
        <div>
            <asp:Label runat="server" text="Formula:"></asp:Label>
            <asp:TextBox runat="server" id="FormulaInput" Text="SUM(B5,2)" Columns="100" /> <br/>
            <asp:Button runat="server" id="ParseButton" text="Parse" />
        </div>
        
        <div id="ParseOutput">
            <asp:Button runat="server" ID="ParseOutputTreeExpandButton" Text="Expand all"/>
            <asp:Button runat="server" ID="ParseOutputTreeCollapseButton" Text="Collapse all"/>
            <asp:TreeView runat="server" id="ParseOutputTree"
                LeafNodeStyle-Font-Name="monospace"
                >
            </asp:TreeView>
        </div>
    </form>
    
    <div id="bugreport">
        <br/>
        <a href="javascript:;" onclick=" javascript:document.getElementById('bug_explanation').style.display = 'block';" style="color: black; font-weight: bold;">Found a bug?</a> <br/> <br/>
        <div id="bug_explanation" style="display: none;">
            Great! <a href="https://github.com/PerfectXL/XLParser/issues">Please report it as a Github issue!</a> <br/> <br/>

            If the bug is with a specific formula/excel file, please include that too.<br/>
            Generally bugs in XLParser are one of the following, please include this type in the report:<br/>
            <ul>
                <li>The parser can't parse a formula that Excel accepts</li>
                <li>The parser parses a formula that Excel doesn't accept.<br/>Note that it might not be always possible to fix these types of bugs.</li>
                <li>The parser interprets a formula wrong, that is it produces a parse tree that doesn't correspond with how Excel behaves.</li>
                <li>There is a "normal" bug in the code around the core parser.</li>
            </ul>
            
        </div>
    </div>
    <div>
        
    </div>
</body>
</html>
