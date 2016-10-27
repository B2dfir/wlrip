# WLrip (beta) #
Version 0.1
By Barney Skeggs (b2dfir.blogspot.com)

‘WaitList.dat’ (WaitList) is data file which has been found to contain stripped text from email, contact and document files as a function of the Windows Search Indexer.

WLrip is a Python 3.5 program that will extract the metadata and body text of each indexed record to a new .txt file, and produce a metadata report in .csv format.
Running WLrip with the ‘-x’ option will produce a .xlsx report with hyperlinks to each .txt file created. This is the recommended method to run WLrip, however it requires the Python ‘XLSXWriter’ module (https://github.com/jmcnamara/XlsxWriter).

Recommended execution of WLrip.py is as follows:
>Wlrip.py -c -x -f filename -o output directory

####Arguments:

| Argument       | Description           |
| ------------- |-------------|
|-c | Removes various null characters, in an attempt to clean up the text output |
|-x |Produces a .xlsx report, as well as the default .csv report.    | 
|-k |Kills the ‘Microsoft Windows Search Indexer’ process, which will lock the WaitList.dat file on a live system. Requires administrator privileges. |
|-f |Specify WaitList.dat file location for processing|
|-o |Specifies an output directory. If not included, the report will be generated within a new folder in the current directory|

I have done my best to write this program in a way that allows it to capture new values (which I have not yet encountered) in the ‘other’ field. Values captured in the ‘other’ field will be appended with a [type], to indicate the field value stored in the data structure. Please inform me or propose changes to the git should you come across new value types for implementation in future release.

#####For Community Contribution

Two values are included in the output report for community analysis. 

The first is Unkn*, which I suspect is actually padding and/or zeros from another value of larger size, as it has had a value of 00 for all tested records so far.
The second is DocID, which I suspect is an ID for the document which was indexed by Windows Search Indexer.

Please see b2dfir.blogspot.com for a description of the data structure understanding that underpins this parser.
