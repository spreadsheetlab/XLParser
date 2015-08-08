﻿var default_formula = "SUM(B5,2)";

var margin = { top: 20, right: 20, bottom: 20, left: 20 },
        width = 500 - margin.right - margin.left,
        height = 500 - margin.top - margin.bottom;

var i;
var tree;
var root;

var diagonal = d3.svg.diagonal()
    .projection(function (d) { return [d.x, d.y]; });

var vis;

function newTree(formula) {
    var encodedFormula = encodeURIComponent(formula);
    var url = "Parse.json?formula=" + encodedFormula

    d3.json(url, function (request, json) {
        //console.log(json)
        //console.log(request)
        if (json !== undefined) {
            var tw = treeWidth(json);
            var th = treeHeight(json);
            //console.log("W: " + tw + " H: " + th);
            //console.log(json);
            var w = Math.max(tw * 75, width);
            var h = Math.max(10 + th * 60, height);
            var imgw = w + margin.right + margin.left;
            var imgh = h + margin.top + margin.bottom;

            tree = d3.layout.tree().size([w, h]);
            i = 0;

            d3.select("#d3viz").html("");

            var svg = d3.select("#d3viz")
            .append("svg")
            .attr("id", "dynamicSVGParsetree")
            .attr("version", 1.1)
            .attr("xmlns", "http://www.w3.org/2000/svg")
            .attr("width", imgw)
            .attr("height", imgh)
            ;

            svg.append("style")
                .attr("type", "text/css")
                //.text("<![CDATA[\n" + svgcss + "\n]]>");
                .text(svgcss);

            vis = svg.append("g")
            .attr("transform", "translate(" + margin.left + "," + margin.top + ")");

            root = json;
            update(root);

            generateImageData(imgw, imgh);
        } else {
            json = JSON.parse(request.response);
            var msg = "<strong>Error:</strong> <code>" + json.error + "</code><br />";
            // Convert to entities to prevent XSS
            msg += "Input: <input id='errorformulainput' disabled value='" + json.formula.replace(/./gm, function (s) { return "&#" + s.charCodeAt(0) + ";"; }) + "'/><br />";
            if (json.messages !== undefined) {
                msg += "<textarea disabled id='errormessages'>" + json.messages + '</textarea>';
            }
            d3.select("#d3viz")
            .html(msg);
        }
    });

    if (ga !== undefined) {
        ga('send', 'pageview', url);
        var imgdatasvg = $('#imgdatasvg');
        imgdatasvg.off('click');
        imgdatasvg.on('click', function () { ga('send', 'pageview', 'parsetree.svg?formula=' + encodedFormula); })
        var imgdatapng = $('#imgdatapng');
        imgdatapng.off('click');
        imgdatapng.on('click', function () { ga('send', 'pageview', 'parsetree.png?formula=' + encodedFormula); })
    }
}

newTree(default_formula);
d3.select('#formulainput').attr("value", default_formula);


function update(source) {

    // Compute the new tree layout.
    var nodes = tree.nodes(root).reverse();
    var links = tree.links(nodes);

    // Normalize for fixed-depth.
    nodes.forEach(function (d) {
        d.y = 10 + d.depth * 60;
    });

    // Update the nodes…
    var node = vis.selectAll("g.node")
        .data(nodes, function (d) { return d.id || (d.id = ++i); });

    // Enter any new nodes at the parent's previous position.
    var nodeEnter = node.enter().append("g")
        .attr("class", "node")
        .attr("transform", function (d) {
            return "translate(" + d.x + "," + d.y + ")";
        })

    nodeEnter.append("circle")
        //.attr("r", 1e-6)
        .attr("r", 8)
        //.style("fill", function(d) { return d._children ? "lightsteelblue" : "#fff"; });
        .style("fill", "#fff");

    nodeEnter.append("text")
        //.attr("x", function(d) { return d.children || d._children ? -10 : 10; })
        .attr("y", function (d) {
            // Put nodes without children (terminals) below, nodes with (nonterminals) above
            return d.children || d._children ? -20 : 20;
        })
        .attr("dy", ".31em")
        .attr("text-anchor", "middle")
        .text(function (d) { return d.name; })
        .style("fill-opacity", 1);

    // Declare the links…
    var link = vis.selectAll("path.link")
        .data(links, function (d) { return d.target.id; });

    // Enter the links.
    link.enter().insert("path", "g")
        .attr("class", "link")
        .attr("d", diagonal);
    // Transition nodes to their new position.
}

function treeWidth(node) {
    if (node.children == undefined) return 1;
    var sum = 0;
    for (var i = 0; i < node.children.length; i++) {
        sum += treeWidth(node.children[i]);
    }
    return sum;
}

function treeHeight(node) {
    if (node.children == undefined) return 1;
    var max = 0;
    for (var i = 0; i < node.children.length; i++) {
        max = Math.max(max, treeHeight(node.children[i]));
    }
    return max + 1;
}

// From: http://techslides.com/save-svg-as-an-image
function generateImageData(imgw, imgh) {
    var html = d3.select("#dynamicSVGParsetree")
          .node().parentNode.innerHTML;

    //console.log(html);
    var svgsrc = 'data:image/svg+xml;base64,' + btoa(html);
    //var img = '<img src="' + imgsrc + '">';
    //d3.select("#imgdata").html(img);
    var imgdatasvg = $('#imgdatasvg')
    imgdatasvg.attr('crossOrigin', 'anonymous');
    imgdatasvg.attr('href', svgsrc);
    imgdatasvg.attr('download', "parsetree.svg");

    var image = new Image;
    image.src = svgsrc;
    image.onload = function () {
        var imgdatapng = $('#imgdatapng');
        try {
            var canvas = document.createElement("canvas");
            canvas.width = imgw;
            canvas.height = imgh;
            canvas.style.backgroundColor = "white";
            var canvasctx = canvas.getContext("2d");
            canvasctx.drawImage(image, 0, 0);
            var pngsrc = canvas.toDataURL("image/png");
            
            imgdatapng.attr('href', pngsrc);
            imgdatapng.attr('download', "parsetree.png");
        }
        catch(e) {
            imgdatapng.attr('href', "javascript: void(0)");
            imgdatapng.off('click');
            imgdatapng.on('click', function() { 
                alert("An error occured while creating PNG.\n\n" +
                    "If you are using Internet Explorer 10 or 11, this page doesn't have enough privileges to allow PNG creation. Increase trust level for this page.\n\n" +
                    "Are you using an older browser? If so try a newer one.\n\n" +
                    "Confirmed to work in Firefox 39 and Chrome 44.");
                return false;
            });
        }
    }

};

var svgcss = ".node circle {\n"+
"    fill: #fff;\n"+
"    stroke: steelblue;\n"+
"    stroke-width: 1.5px;\n"+
"}\n"+
".node text {\n" +
"   font-family: 'Helvetica Neue', Helvetica, sans-serif;"+
"   font-size: 14px;\n"+
"}\n"+
"path.link {\n"+
"    fill: none;\n"+
"    stroke: #cfcfcf;\n"+
"    stroke-width: 1.5px;\n"+
"}";