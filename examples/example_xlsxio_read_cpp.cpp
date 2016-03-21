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
  bool GetNextCellString (char*& value);
  bool GetNextCellString (std::string& value);
  bool GetNextCellInt (long& value);
  bool GetNextCellFloat (double& value);
  bool GetNextCellDateTime (time_t& value);
  inline XLSXIOReaderSheet& operator >> (char*& value) { GetNextCellString(value); return *this; }
  inline XLSXIOReaderSheet& operator >> (std::string& value) { GetNextCellString(value); return *this; }
  inline XLSXIOReaderSheet& operator >> (int& value) { long l; GetNextCellInt(l); l = value; return *this; }
  inline XLSXIOReaderSheet& operator >> (long& value) { GetNextCellInt(value); return *this; }
  inline XLSXIOReaderSheet& operator >> (long long& value) { long l; GetNextCellInt(l); l = value; return *this; }
  inline XLSXIOReaderSheet& operator >> (unsigned int& value) { long l; GetNextCellInt(l); l = value; return *this; }
  inline XLSXIOReaderSheet& operator >> (unsigned long& value) { long l; GetNextCellInt(l); l = value; return *this; }
  inline XLSXIOReaderSheet& operator >> (unsigned long long& value) { long l; GetNextCellInt(l); l = value; return *this; }
  inline XLSXIOReaderSheet& operator >> (double& value) { GetNextCellFloat(value); return *this; }
  //inline XLSXIOReaderSheet& operator >> (time_t& value) { GetNextCellString(value); return *this; }
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

bool XLSXIOReaderSheet::GetNextCellString (char*& value)
{
  if (!xlsxioread_sheet_next_cell_string(sheethandle, &value)) {
    value = NULL;
    return false;
  }
  return true;
}

bool XLSXIOReaderSheet::GetNextCellString (std::string& value)
{
  char* result;
  if (!xlsxioread_sheet_next_cell_string(sheethandle, &result)) {
    value.clear();
    return false;
  }
  value.assign(result);
  free(result);
  return true;
}

bool XLSXIOReaderSheet::GetNextCellInt (long& value)
{
  if (!xlsxioread_sheet_next_cell_int(sheethandle, &value)) {
    value = 0;
    return false;
  }
  return true;
}

bool XLSXIOReaderSheet::GetNextCellFloat (double& value)
{
  if (!xlsxioread_sheet_next_cell_float(sheethandle, &value)) {
    value = 0;
    return false;
  }
  return true;
}

bool XLSXIOReaderSheet::GetNextCellDateTime (time_t& value)
{
  if (!xlsxioread_sheet_next_cell_datetime(sheethandle, &value)) {
    value = 0;
    return false;
  }
  return true;
}



int main (int argc, char* argv[])
{
  XLSXIOReader* xlsxfile = new XLSXIOReader(filename);
  XLSXIOReaderSheet* xlsxsheet = xlsxfile->OpenSheet(NULL, XLSXIOREAD_SKIP_EMPTY_ROWS);
  if (xlsxsheet) {
    std::string value;
    while (xlsxsheet->GetNextRow()) {
/*
      while (xlsxsheet->GetNextCellString(value)) {
        printf("%s\t", value.c_str());
      }
*/
      std::string s;
      *xlsxsheet >> s;
      printf("%s\t", s.c_str());
      char* n;
      *xlsxsheet >> n;
      printf("%s\t", n);
      free(n);
      int i;
      *xlsxsheet >> i;
      printf("%i\t", i);
      *xlsxsheet >> i;
      printf("%i\t", i);
      double d;
      *xlsxsheet >> d;
      printf("%.6G\t", d);
      printf("\n");
    }
    delete xlsxsheet;
  }
  delete xlsxfile;
  return 0;
}
