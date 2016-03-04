XLSX I/O
========
Cross-platform C library for reading values from and writing values to .xlsx files.

Description
-----------
XLSX I/O aims to provide a C library for reading and writing .xlsx files.
The .xlsx file format is the native format used by Microsoft(R) Excel(TM) since version 2007.

Goal
----
The library was written with the following goals in mind:
- written in standard C, but allows being used by C++
- simple interface
- small footprint
- portable across different platforms (Windows, *nix)
- minimal dependancies: only depends on expat and libzip (which in turn depends on zlib)
- seperate library for reading and writing .xlsx files

Reading .xlsx files:
- intended to process .xlsx files as a data table, which assumes the following:
  + assumes the first row contains header names
  + assumes the next rows contain values in the same columns as where the header names are supplied
  + only values are processed, anything else is ignored (formulas, layout, graphics, charts, ...)
- the entire shared string table is loaded in memory (can be large for big spreadsheets with a lot of different values)

Writing .xlsx files:
- intended for writing data tables as .xlsx files, which assumes the following:
  + only support for writing data (no support for formulas, layout, graphics, charts, ...)
  + no support for multiple worksheets (only one worksheet per file)
- on the fly file generation without the need to buffer data in memory
- no support for shared strings (all values are written as inline strings)

License
-------
XLSX I/O is released under the terms of the MIT License (MIT), see LICENSE.txt.
