This package is a rewrite for the DIF package that handles China Life trustee's Excel files for Diversified income fund, Balanced fund, Guarantee fund and Growth fund.

The purpose is to have a more clear structure in the program, namely, now file handling follows the approach of: read file to lines -- group lines to sections -- section to records. This is much clearer than the old package's approach, read the file line by line, parse the line while reading it.

The package is divided into two modules:

dif.py: read the Excel file and save the holdings as a list of dictionary objects, i.e., records. To access certain type of records, say cash or equity, use the record's type to filter them out. The module also reads portfolio summary, and test whether sums of different types of positions (cash, equity, bond, futures) equal their subtotal in summary.

geneva.py: use the records from dif.py and save them as csv files to be uploaded for reconciliation with Advent Geneva system. It has a open_dif() function that has the same interface as DIF.open_dif.py's open_dif() function, so that the new open_dif() function can be used by the recon_helper.py in the reconciliation package.


To be improved:

1. Futures positions don't have any dates converted to yyyy-mm-dd format yet. Because the maturity date of futures is like '2018 Sep' instead of an exact date. Don't know how to process it. See 'samples/CL Franklin DIF 2018-05-28(2nd Revised).xls'.

2. The reconciliation csv's have lots of unnecessary fields. But to delete them requires changing Geneva reconciliation setup. So I'll wait until the reconciliation is stable.



+++++++++++++++++++++
ver 0.1 @ 2018-6-4
+++++++++++++++++++++
1. Tested with Excel files of Diversified income fund, Balanced fund and Guarantee fund.
2. Can work with recon_helper.py.
