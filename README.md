# ODS-PHP-Parser
Simple ODS Parser written in PHP programming Language.

This Parser is simple.

It reads only one Level of Tables. This means that Tables are not read recursively, Sub-Tables inside Tables are not properly read by this Parser if they are found. I personally think that it is a bad Practice to put Tables inside the Cells of a Table, as it makes Mess in the Spreadsheet Document. 

This Parser reads only most important Tags. 

At this Moment it supports reading ordinary Tables (Sheets), Rows, Columns and Cells. Cells which are covered by big spanned Cells are read as well. Some special Tags as Groups of Rows, Groups of Columns and Headers are not supported at this Moment. 

This Parser is made using the built-in PHP's XML Parser. 

The ODS Format is an open Format, known as 'OpenDocument' Format. The 'S' in 'ODS' means 'Spreadsheet'. This Parser uses 'OpenDocument' Format Version 1.2, which is the latest to this very Day.

See Examples in the same Folder for better Understanding of this Parser.
