#include <stdlib.h>
#include <stdio.h>
#include <string>
#include "xlsxio_read.h"

const char* filename = "example.xlsx";



class XLSXIOReader
{
 private:
  xlsxioreader handle;
 public:
  XLSXIOReader (const char* filename);
  ~XLSXIOReader ();
  class XLSXIOReaderSheet* OpenSheet (const char* sheetname, unsigned int flags);
};



class XLSXIOReaderSheet
{
  friend class XLSXIOReader;
 private:
  xlsxioreadersheet sheethandle;
  XLSXIOReaderSheet (xlsxioreader xlsxhandle, const char* sheetname, unsigned int flags);
 public:
  ~XLSXIOReaderSheet ();
  bool GetNextRow ();
  char* GetNextCell ();
  bool GetNextCell (std::string& value);
};



XLSXIOReader::XLSXIOReader (const char* filename)
{
  handle = xlsxioread_open(filename);
}

XLSXIOReader::~XLSXIOReader ()
{
  xlsxioread_close(handle);
}



class XLSXIOReaderSheet* XLSXIOReader::OpenSheet (const char* sheetname, unsigned int flags)
{
  return new XLSXIOReaderSheet(handle, sheetname, flags);
}



XLSXIOReaderSheet::XLSXIOReaderSheet (xlsxioreader xlsxhandle, const char* sheetname, unsigned int flags)
{
  sheethandle = xlsxioread_sheet_open(xlsxhandle, sheetname, flags);
}

XLSXIOReaderSheet::~XLSXIOReaderSheet ()
{
  xlsxioread_sheet_close(sheethandle);
}

bool XLSXIOReaderSheet::GetNextRow ()
{
  return (xlsxioread_sheet_next_row(sheethandle) != 0);
}

char* XLSXIOReaderSheet::GetNextCell ()
{
  return xlsxioread_sheet_next_cell(sheethandle);
}

bool XLSXIOReaderSheet::GetNextCell (std::string& value)
{
  char* result = xlsxioread_sheet_next_cell(sheethandle);
  if (!result) {
    value.clear();
    return false;
  }
  value.assign(result);
  free(result);
  return true;
}



int main (int argc, char* argv[])
{
  XLSXIOReader* xlsxfile = new XLSXIOReader(filename);
  XLSXIOReaderSheet* xlsxsheet = xlsxfile->OpenSheet(NULL, XLSXIOREAD_SKIP_EMPTY_ROWS);
  if (xlsxsheet) {
    std::string value;
    while (xlsxsheet->GetNextRow()) {
      while (xlsxsheet->GetNextCell(value)) {
        printf("%s\t", value.c_str());
      }
      printf("\n");
    }
    delete xlsxsheet;
  }
  delete xlsxfile;
  return 0;
}
