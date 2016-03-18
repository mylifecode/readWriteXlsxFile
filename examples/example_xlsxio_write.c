#include <stdlib.h>
#include <stdio.h>
#include <unistd.h>
#include "xlsxio_write.h"

const char* filename = "example.xlsx";

int main (int argc, char* argv[])
{
  xlsxiowriter handle;
  //remove the destination file first if it exists
  unlink(filename);
  //open .xlsx file for writing
  if ((handle = xlsxiowrite_open(filename, "MySheet")) == NULL) {
    fprintf(stderr, "Error creating .xlsx file\n");
    return 1;
  }
  //write column names
  xlsxiowrite_add_column(handle, "Col1", 4);
  xlsxiowrite_add_column(handle, "Col2", 21);
  xlsxiowrite_add_column(handle, "Col3", 12);
  xlsxiowrite_add_column(handle, "Col4", 2);
  xlsxiowrite_add_column(handle, "Col5", 4);
  xlsxiowrite_add_column(handle, "Col6", 16);
  xlsxiowrite_add_column(handle, "Col7", 10);
  xlsxiowrite_next_row(handle);
  //write data
  int i;
  for (i = 0; i < 1000; i++) {
    xlsxiowrite_add_cell_string(handle, "Test");
    xlsxiowrite_add_cell_string(handle, "A b  c   d    e     f\nnew line");
    xlsxiowrite_add_cell_string(handle, "&% <test> \"'");
    xlsxiowrite_add_cell_string(handle, NULL);
    xlsxiowrite_add_cell_int(handle, i);
    xlsxiowrite_add_cell_datetime(handle, time(NULL));
    xlsxiowrite_add_cell_float(handle, 3.1415926);
    xlsxiowrite_next_row(handle);
  }
  //close .xlsx file
  xlsxiowrite_close(handle);
  return 0;
}
