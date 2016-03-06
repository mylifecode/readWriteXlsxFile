#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include "xlsxio_read.h"

//const char* filename = "example.xlsx";
//const char* filename = "1424345137-Sample.xlsx";
const char* filename = "mytest.xlsx";

struct xlsx_list_sheets_data {
  char* firstsheet;
};

int xlsx_list_sheets_callback (const char* name, void* callbackdata)
{
  struct xlsx_list_sheets_data* data = (struct xlsx_list_sheets_data*)callbackdata;
  if (!data->firstsheet)
    data->firstsheet = strdup(name);
  printf(" - %s\n", name);
  return 0;
}

int sheet_row_callback (size_t row, size_t maxcol, void* callbackdata)
{
  printf("\n");
printf("[[%i,%i]]\n", (int)row, (int)maxcol);/////
  return 0;
}

int sheet_cell_callback (size_t row, size_t col, const char* value, void* callbackdata)
{
  if (col > 1)
    printf("\t");
printf("[%i,%i]", (int)row, (int)col);/////
  if (value)
    printf("%s", value);
  return 0;
}

int main (int argc, char* argv[])
{
  xlsxioreader xlsxioread;
  //open .xlsx file for reading
  if ((xlsxioread = xlsxioread_open(filename)) == NULL) {
    fprintf(stderr, "Error opening .xlsx file\n");
    return 1;
  }
  //list available sheets
  struct xlsx_list_sheets_data sheetdata;
  printf("Available sheets:\n");
  sheetdata.firstsheet = NULL;
  xlsxioread_list_sheets(xlsxioread, xlsx_list_sheets_callback, &sheetdata);

  //perform tests
  xlsxioread_process(xlsxioread, sheetdata.firstsheet, XLSXIOREAD_SKIP_EMPTY_ROWS | XLSXIOREAD_SKIP_EXTRA_CELLS, sheet_cell_callback, sheet_row_callback, NULL);

  //clean up
  free(sheetdata.firstsheet);
  xlsxioread_close(xlsxioread);
  return 0;
}
