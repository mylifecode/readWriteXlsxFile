#include <stdlib.h>
#include <stdio.h>
#include <string>
#include "xlsxio_write.h"

/*! \class XLSXIOWriter
 *  \brief class for writing data to an .xlsx file
 *\details C++ wrapper for xlsxiowrite_ functions.
 */
class XLSXIOWriter
{
 private:
  xlsxiowriter handle;
 public:

  /*! \brief XLSXIOWriter constructor, creates and opens .xlsx file
   * \param  filename      path of .xlsx file to open
   * \param  sheetname     name of worksheet
   * \param  detectionrows number of rows to buffer in memory, zero for none, defaults to 5
   * \sa     xlsxiowrite_open()
   */
  XLSXIOWriter (const char* filename, const char* sheetname = NULL, size_t detectionrows = 5);

  /*! \brief XLSXIOWriter destructor, closes .xlsx file
   * \sa     xlsxiowrite_close()
   */
  ~XLSXIOWriter ();

  /*! \brief specify the row height to use for the current and next rows
   * \param  height        row height (in text lines), zero for unspecified
   * Must be called before the first call to any Add method of the current row
   * \sa     xlsxiowrite_set_row_height()
   */
  void SetRowHeight (size_t height = 0);
  void AddColumn (const char* name, int width = 0);
  void AddCellString (const char* value);
  void AddCellInt (long long value);
  void AddCellFloat (double value);
  void AddCellDateTime (time_t value);
  inline XLSXIOWriter& operator << (const char* value) { AddCellString(value); return *this; }
  inline XLSXIOWriter& operator << (const std::string& value) { AddCellString(value.c_str()); return *this; }
  inline XLSXIOWriter& operator << (int value) { AddCellInt(value); return *this; }
  inline XLSXIOWriter& operator << (long value) { AddCellInt(value); return *this; }
  inline XLSXIOWriter& operator << (long long value) { AddCellInt(value); return *this; }
  inline XLSXIOWriter& operator << (unsigned int value) { AddCellInt(value); return *this; }
  inline XLSXIOWriter& operator << (unsigned long value) { AddCellInt(value); return *this; }
  inline XLSXIOWriter& operator << (unsigned long long value) { AddCellInt(value); return *this; }
  inline XLSXIOWriter& operator << (float value) { AddCellFloat(value); return *this; }
  inline XLSXIOWriter& operator << (double value) { AddCellFloat(value); return *this; }
  //inline XLSXIOWriter& operator << (time_t value) { AddCellDateTime(value); return *this; }
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
    *xlsxfile << "Test" << (char*)NULL << i;
    xlsxfile->AddCellDateTime(time(NULL));
    *xlsxfile << 3.1415926;
    xlsxfile->NextRow();
  }
  delete xlsxfile;
  return 0;
}
