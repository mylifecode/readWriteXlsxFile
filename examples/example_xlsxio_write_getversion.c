#include <stdio.h>
#include "xlsxio_write.h"

int main (int argc, char* argv[])
{
  //display version string
  printf("Version: %s\n", xlsxiowrite_get_version_string());

  //get version numbers
  int major, minor, micro;
  xlsxiowrite_get_version(&major, &minor, &micro);
  printf("Version: %i.%i.%i\n", major, minor, micro);
  return 0;
}
