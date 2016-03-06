#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include "xlsxio_read.h"

const char* filename = "example.xlsx";

int main (int argc, char* argv[])
{
  xlsxioreader xlsxioread;
  //open .xlsx file for reading
  if ((xlsxioread = xlsxioread_open(filename)) == NULL) {
    fprintf(stderr, "Error opening .xlsx file\n");
    return 1;
  }
  char* value;
  xlsxioreadersheet sheet = xlsxioread_sheet_open(xlsxioread, NULL, XLSXIOREAD_SKIP_EMPTY_ROWS);
  while (xlsxioread_sheet_next_row(sheet)) {
    while ((value = xlsxioread_sheet_next_cell(sheet)) != NULL) {
      printf("%s\t", value);
      free(value);
    }
    printf("\n");
  }
  xlsxioread_sheet_close(sheet);

  //clean up
  xlsxioread_close(xlsxioread);
  return 0;
}
