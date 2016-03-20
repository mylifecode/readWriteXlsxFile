#include <stdlib.h>
#include <stdio.h>
#include "xlsxio_write.h"

class XLSXIOWriter
{
 private:
  xlsxiowriter handle;
 public:
  XLSXIOWriter (const char* filename, const char* sheetname = NULL, size_t detectionrows = 5);
  ~XLSXIOWriter ();
  void SetRowHeight (size_t height = 0);
  void AddColumn (const char* name, int width = 0);
  void AddCellString (const char* value);
  void AddCellInt (long long value);
  void AddCellFloat (double value);
  void AddCellDateTime (time_t value);
  void NextRow ();
};



XLSXIOWriter::XLSXIOWriter (const char* filename, const char* sheetname, size_t detectionrows)
{
  unlink(filename);
  handle = xlsxiowrite_open(filename, sheetname);
  xlsxiowrite_set_detection_rows(handle, detectionrows);
}

XLSXIOWriter::~XLSXIOWriter ()
{
  xlsxiowrite_close(handle);
}

void XLSXIOWriter::SetRowHeight (size_t height)
{
  xlsxiowrite_set_row_height(handle, height);
}

void XLSXIOWriter::AddColumn (const char* name, int width)
{
  xlsxiowrite_add_column(handle, name, width);
}

void XLSXIOWriter::AddCellString (const char* value)
{
  xlsxiowrite_add_cell_string(handle, value);
}

void XLSXIOWriter::AddCellInt (long long value)
{
  xlsxiowrite_add_cell_int(handle, value);
}

void XLSXIOWriter::AddCellFloat (double value)
{
  xlsxiowrite_add_cell_float(handle, value);
}
void XLSXIOWriter::AddCellDateTime (time_t value)
{
  xlsxiowrite_add_cell_datetime(handle, value);
}

void XLSXIOWriter::NextRow ()
{
  xlsxiowrite_next_row(handle);
}



const char* filename = "example.xlsx";

int main (int argc, char* argv[])
{
  XLSXIOWriter* xlsxfile = new XLSXIOWriter(filename);
  xlsxfile->SetRowHeight(1);
  xlsxfile->AddColumn("Col1");
  xlsxfile->AddColumn("Col2");
  xlsxfile->AddColumn("Col3");
  xlsxfile->AddColumn("Col4");
  xlsxfile->AddColumn("Col5");
  xlsxfile->NextRow();
  int i;
  for (i = 0; i < 1000; i++) {
    xlsxfile->AddCellString("Test");
    xlsxfile->AddCellString((char*)NULL);
    xlsxfile->AddCellInt((long long)i);
    xlsxfile->AddCellDateTime(time(NULL));
    xlsxfile->AddCellFloat(3.1415926);
    xlsxfile->NextRow();
  }
  delete xlsxfile;
  return 0;
}
