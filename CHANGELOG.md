# Changelog

## 1.2

Fixes [#16](https://github.com/PerfectXL/XLParser/issues/16), [#17](https://github.com/PerfectXL/XLParser/issues/17), [#19](https://github.com/PerfectXL/XLParser/issues/3)

* Made it easier to modify the grammar in your own class by extending the grammar class
* Can now parse non-numeric filenames (`=[file]Sheet!A1`)
* Parsing of the `Prefix` nonterminal is changed and is now a little bit more uniform. `ExcelFormulaParser.GetPrefixInfo` gives prefix information in an easy to use format.
* Can now parse [Structured References](https://support.office.com/en-us/article/Using-structured-references-with-Excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e). See [#16](https://github.com/PerfectXL/XLParser/issues/16) for caveats.
* You can now select the XLParser version to use in the web demo

## 1.1.4

* Added some missing methods that test for specific types of operators
* Added tests and fixes if necessary for methods that were missing tests

## 1.1.3

Reference implementation of the Excel grammar published in the upcoming paper "A Grammar for Spreadsheet Formulas Evaluated on Two Large Datasets" by E. Aivaloglou, D. Hoepelman and F. Hermans.

* Added all formulas from EUSES and Enron datasets and tests to check if they all parse
* Made parser thread safe
* Fixed [#9](https://github.com/PerfectXL/XLParser/issues/9): some tokens which would not accept all unicode characters (like UDF) now do so
* `'Sheet1:Sheet5'` will now correctly parse as `MULTIPLESHEETS` instead of a single sheet


## 1.1.2

Fixed [#1](https://github.com/PerfectXL/XLParser/issues/1), [#2](https://github.com/PerfectXL/XLParser/issues/2), [#4](https://github.com/PerfectXL/XLParser/issues/4).

* Added a web demo in app/XLParser.Web which generates parse tree images
* All UDF's now use the same nonterminal
* Non-Prefixed UDFs can now be part of a reference expression
* IF and CHOOSE functions can now be part of a reference expression
* Reference functions INDEX,OFFSET and INDIRECT can no longer have a prefix
* Operator precedence for reference operators (: , and intersection) is now correct
* Fixed printing of reference operators

## 1.0.0

First public release.
Corresponds to pre-print/reviewer version of the paper
