# Changelog

<!--*Changes which are in this source code, but not yet in a release*:-->
## 1.3.0

* Build for .NET 4.5.2, 4.6.1 and standard 1.6, thanks [igitur](https://github.com/spreadsheetlab/XLParser/pull/61).
* Remove embedded Irony dependency in favor of [daxnet](https://github.com/daxnet)s [updated fork](https://github.com/daxnet/irony).

## 1.2.4
Reference implementation of the Excel grammar published in the Journal of Systems and Software SCAM special issue paper "A Grammar for Spreadsheet Formulas Evaluated on Two Large Datasets" by E. Aivaloglou, D. Hoepelman and F. Hermans.

* Fixed several errors in which names/named ranges were allowed
  * Question marks are now allowed
  * Can now start with all unicode letters (e.g. `=äbc`)
  * Corrected characters which are allowed if the name starts with a cell name or TRUE/FALSE (e.g. `=A1.MYNAME`)
* Allow for whitespace-only sheetnames (e.g. `='    '!A1`), altough they will always be returned as `" "` by `PrefixInfo` 
* Made some corrections in how multiple sheet references (`=Sheet1:Sheet3!A1`) are parsed
* Removed escape sequences in strings (e.g. `"Line1\nLine2"`) as these are not part of the Excel formula language
* Added support for structured references to a complete table (e.g. `=MyTable[]`)

## 1.2.3
* Adds support for special characters in structured references.

## 1.2.2

* Adds equality to `PrefixInfo` class
* Fixes parse error if external reference file path contains a space (`='C:\My Dir\[file.xlsx]Sheet'!A1`)
* `ExcelFormulaParser.SkipToRelevant` no longer skips references without a prefix.
<br> This is a breaking change, but the old behavior is arguably a bug. An argument is added to restore old behavior, defaults to new behavior.

## 1.2.1

* Adds `GetReferenceNodes` method to `ExcelFormulaParser`

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

Reference implementation of the Excel grammar published in the paper "A Grammar for Spreadsheet Formulas Evaluated on Two Large Datasets" by E. Aivaloglou, D. Hoepelman and F. Hermans.

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
