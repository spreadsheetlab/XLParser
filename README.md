XLParser: Excel formula parser
==========

XLParser is the reference implementation of the Excel grammar published in the upcoming paper "A Grammar for Spreadsheet Formulas Evaluated on Two Large Datasets" by E. Aivaloglou, D. Hoepelman and F. Hermans.

XLParser can parse Excel formulas and is intended to facilitate the analysis of spreadsheet formulas, and for that purpose produces compact parse trees.
XLParser  has a 99.99% success rate on the [Enron](http://www.felienne.com/archives/3634) and [EUSES](http://eusesconsortium.org/resources.php) datasets.
Note however that XLParser is not very restrictive, and thus might parse formulas that Excel would reject as invalid, keep this in mind when parsing user input with XLParser.

XLParser is based on the C# [Irony parser framework](https://irony.codeplex.com/).

**XLParser will become available in the week of June 29th.** 