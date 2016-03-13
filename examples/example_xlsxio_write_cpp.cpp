#include <stdlib.h>
#include <stdio.h>
#include "xlsxio_write.h"

class XLSXIOWriter
{
 private:
  xlsxiowriter handle;
 public:
  XLSXIOWriter (const char* filename);
  ~XLSXIOWriter ();
  void AddCell (const char* value);
  void AddCell (long long value);
  void AddCell (double value);
  void AddCell (time_t value);
  void NextRow ();
};



XLSXIOWriter::XLSXIOWriter (const char* filename)
{
  unlink(filename);
  handle = xlsxiowrite_open(filename);
}

XLSXIOWriter::~XLSXIOWriter ()
{
  xlsxiowrite_close(handle);
}

void XLSXIOWriter::AddCell (const char* value)
{
  xlsxiowrite_add_cell_string(handle, value);
}

void XLSXIOWriter::AddCell (long long value)
{
  xlsxiowrite_add_cell_int(handle, value);
}

void XLSXIOWriter::AddCell (double value)
{
  xlsxiowrite_add_cell_float(handle, value);
}
void XLSXIOWriter::AddCell (time_t value)
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
  int i;
  for (i = 0; i < 1000; i++) {
    xlsxfile->AddCell("Test");
    xlsxfile->AddCell((char*)NULL);
    xlsxfile->AddCell((long long)i);
    xlsxfile->AddCell(time(NULL));
    xlsxfile->AddCell(3.1415926);
    xlsxfile->NextRow();
  }
  delete xlsxfile;
  return 0;
}
