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
- minimal dependancies: only depends on expat (only for reading) and minizip or libzip (which in turn depend on zlib)
- separate library for reading and writing .xlsx files

Reading .xlsx files:
- intended to process .xlsx files as a data table, which assumes the following:
  + assumes the first row contains header names
  + assumes the next rows contain values in the same columns as where the header names are supplied
  + only values are processed, anything else is ignored (formulas, layout, graphics, charts, ...)
- the entire shared string table is loaded in memory (warning: could be large for big spreadsheets with a lot of different values)
- supports .xlsx files without shared string table
- worksheet data itself is read on the fly without the need to buffer data in memory
- 2 methods are provided
  + a simple method that allows the application to iterate trough rows and cells
  + an advanced method (with less overhead) which calls callback functions for each cell and after each row

Writing .xlsx files:
- intended for writing data tables as .xlsx files, which assumes the following:
  + only support for writing data (no support for formulas, layout, graphics, charts, ...)
  + no support for multiple worksheets (only one worksheet per file)
- on the fly file generation without the need to buffer data in memory
- no support for shared strings (all values are written as inline strings)

Building from source
--------------------

Requirements:
- a C compiler like gcc or clang, on Windows MinGW and MinGW-w64 are supported
- the dependancy libraries (see Dependancies)
- a shell environment, on Windows MSYS is supported
- the make command
- CMake version 2.6 or higher (optional, but preferred)

There are 2 methods to build XLSX I/O:
- using the basic Makefile included
- using CMake (preferred)

Building with make
- build and install by running `make install` optionally followed by:
  + `PREFIX=<path>`	Base path were files will be installed (defaults to /usr/local)
  + `WITH_LIBZIP=1`	Use libzip instead of minizip
  + `WIDE=1`	Also build UTF-16 library (xlsxio_readw)
  + `STATICDLL=1`	Build a static DLL (= doesn't depend on any other DLLs) - only supported on Windows

Building with CMake (preferred method)
- configure by running `cmake -G"Unix Makefiles"` (or `cmake -G"MSYS Makefiles"` on Windows) optionally followed by:
  + `-DCMAKE_INSTALL_PREFIX:PATH=<path>`  Base path were files will be installed
  + `-DBUILD_STATIC:BOOL=OFF`             Don't build static libraries
  + `-DBUILD_SHARED:BOOL=OFF`             Don't build shared libraries
  + `-DBUILD_TOOLS:BOOL=OFF`              Don't build tools (only libraries)
  + `-DBUILD_EXAMPLES:BOOL=OFF`           Don't build examples
  + `-DWITH_LIBZIP:BOOL=ON`               Use libzip instead of Minizip
  + `-DWITH_WIDE:BOOL=ON`                 Also build UTF-16 library (libxlsxio_readw)
- build and install by running `make install` (or `make install/strip` to strip symbols)

For Windows prebuilt binaries are also available for download (both 32-bit and 64-bit)

Command line utilities
----------------------
Some command line utilities are included:
- xlsxio_xlsx2csv: converts all sheets in all specified .xlsx files to individual CSV (Comma Separated Values) files.
- xlsxio_csv2xlsx: converts all specified CSV (Comma Separated Values) files to .xlsx files.

Dependancies
------------
This project has the following depencancies:
- [expat](http://www.libexpat.org/) (only for libxlsxio_read)
and
- [minizip](http://www.winimage.com/zLibDll/minizip.html) (libxlsxio_read and libxlsxio_write)
or
- [libzip](http://www.nih.at/libzip/) (libxlsxio_read and libxlsxio_write)

License
-------
XLSX I/O is released under the terms of the MIT License (MIT), see LICENSE.txt.
