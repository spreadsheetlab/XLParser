## XLParser
A C# Excel formula parser with the following properties:

* **High compatiblity**<br/>
  XLParser has been tested on over a million real-world formulas and has a 99.9% succesful parse rate.
* **Compact parse trees**<br/>
  XLParser was designed for and is used in research of Excel spreadsheets and refactoring, all of which are easier if parse trees are smaller
* **Compact grammar**<br />
  [Our grammar](https://github.com/PerfectXL/XLParser/blob/master/src/XLParser/ExcelFormulaGrammar.cs) contains less than 100 tokens and production rules, and is thus easy to implement in another language or parser generator.

### Quickstart

You can parse formulas and view parse trees in the [**online demo**](http://xlparser.perfectxl.nl/demo).

**or**

1. Download the [latest release](https://github.com/PerfectXL/XLParser/releases/latest)
2. Extract somewhere convenient
3. Build Irony GrammarExplorer, see the Debugging section below, and open it
4. Click on the `...` button at the top and select `Add Grammar`
5. Point to the `XLParser.dll` file in the folder you extracted and click ok
6. You can now parse formulas and see the trees in the `Test` tab

### Background

XLParser is the reference implementation of the Excel grammar published in the paper ["A Grammar for Spreadsheet Formulas Evaluated on Two Large Datasets" by Efthimia Aivaloglou, David Hoepelman and Felienne Hermans](https://drive.google.com/file/d/0B79P2Uym3JjvMjlaWWtnTWRLQmc/view?usp=sharing).

XLParser can parse Excel formulas and is intended to facilitate the analysis of spreadsheet formulas, and for that purpose produces compact parse trees.
XLParser  has a 99.99% success rate on the [Enron](http://www.felienne.com/archives/3634) and [EUSES](http://eusesconsortium.org/resources.php) datasets.
Note however that XLParser is not very restrictive, and thus might parse formulas that Excel would reject as invalid, keep this in mind when parsing user input with XLParser.

XLParser is based on the C# [Irony parser framework](https://irony.codeplex.com/).

## Using XLParser

### Building

Open the `XLParser.sln` file in `src/` in Visual Studio 2015 or higher and press build. The dependencies are already included in compiled form in this repository.

### Using in your project

The easiest way to add the dependency to your project is with [NuGet](https://www.nuget.org/packages/XLParser/)

The `ExcelFormulaParser` class is your main entry point. You can parse a formula through `ExcelFormulaParser.Parse("yourformula")`.

`ExcelFormulaParser` has several useful methods that operate directly on the parse tree like `AllNodes` to traverse the whole tree or `GetFunction` to get the function name of a node that represents a function call. You can `Print` any node.

`FormulaAnalyzer` contains some example functionality for analyzing the parse tree.

### Debugging

Irony, the parser framework XLParser uses, includes a tool called the "grammar explorer". This is a great way to play around with the grammar and parse trees.
To use this tool, you first need to build it once by opening the Irony solution (`lib/Irony/Irony_All.2012.sln`) and building it with release configuration. After that you can use the binary in `lib/Irony/Irony.GrammarExplorer/bin/Release/Irony.GrammarExplorer.exe`.

To load the XLParser grammar, first make sure you have built XLParser. Then open the GrammarExplorer and add the grammar (`...` button) from `src/XLParser/bin/Debug/XLParser.dll`.

In Visual Studio you can see the printed version of any node during debugging by adding `yournode.Print(),ac` in the watch window.

### Documentation

See [the paper](https://drive.google.com/file/d/0B79P2Uym3JjvMjlaWWtnTWRLQmc/view?usp=sharing) and [the API documentation](https://spreadsheetlab.github.io/XLParser/api/namespace_x_l_parser.html).

### License

All files of XLParser are released under the Mozilla Public License 2.0.

Roughly this means that you can make any alterations you like and can use this library in any project, even closed-source and statically linked, as long as you publish any modifications to the library.
