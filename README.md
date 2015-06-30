## XLParser
A C# Excel formula parser

### Introduction

XLParser is the reference implementation of the Excel grammar published in the upcoming paper "A Grammar for Spreadsheet Formulas Evaluated on Two Large Datasets" by E. Aivaloglou, D. Hoepelman and F. Hermans.

XLParser can parse Excel formulas and is intended to facilitate the analysis of spreadsheet formulas, and for that purpose produces compact parse trees.
XLParser  has a 99.99% success rate on the [Enron](http://www.felienne.com/archives/3634) and [EUSES](http://eusesconsortium.org/resources.php) datasets.
Note however that XLParser is not very restrictive, and thus might parse formulas that Excel would reject as invalid, keep this in mind when parsing user input with XLParser.

XLParser is based on the C# [Irony parser framework](https://irony.codeplex.com/).

### License

All files of XLParser are released under the Mozilla Public License 2.0.

Roughly this means that you can make any alterations you like and can use this library in any project, even closed-source and statically linked, as long as you publish any modifications to the library.

### How to build

Open the `XLParser.sln` file in `src/` in Visual Studio 2013 or higher and press build. The dependencies are already included in compiled form in this repository.

### How to use in your project

The `ExcelFormulaParser` class is your main entry point. You can parse a formula through `ExcelFormulaParser.Parse("yourformula")`.

`ExcelFormulaParser` has several useful methods that operate directly on the parse tree like `AllNodes` to traverse the whole tree or `GetFunction` to get the function name of a node that represents a function call. You can `Print` any node. In visual studio you can see the printed version of any node during debugging by adding `yournode.Print(),ac` in the watch window.

`FormulaAnalyzer` contains some example functionality for analyzing the parse tree.

### How to debug or experiment

Irony, the parser framework XLParser uses, includes a tool called the "grammar explorer". This is a great way to play around with the grammar and parse trees.
To use this tool, you first need to build it once by opening the Irony solution (`lib/Irony/Irony_All.2012.sln`) and building it with release configuration. After that you can use the binary in `lib/Irony/Irony.GrammarExplorer/bin/Release/Irony.GrammarExplorer.exe`.

To load the XLParser grammar, first make sure you have built XLParser. Then open the GrammarExplorer and add the grammar (`...` button) from `src/XLParser/bin/Debug/XLParser.dll`.