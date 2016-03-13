#include <stdio.h>
#include "xlsxio_write.h"
#include "xlsxio_version.h"

int main (int argc, char* argv[])
{
  /* The following methods call the library to get information, this is the preferred method. */

  //get version string from library
  printf("Version: %s\n", xlsxiowrite_get_version_string());

  //get version numbers from library
  int major, minor, micro;
  xlsxiowrite_get_version(&major, &minor, &micro);
  printf("Version: %i.%i.%i\n", major, minor, micro);

  /* The following methods use header file xlsxio_version.h to get information, avoid when using shared libraries. */

  //get version string from header
  printf("Version: %s\n", XLSXIO_VERSION_STRING);

  //get version numbers from header
  printf("Version: %i.%i.%i\n", XLSXIO_VERSION_MAJOR, XLSXIO_VERSION_MINOR, XLSXIO_VERSION_MICRO);

  //get library name and version string from header
  printf("Name and version: %s\n", XLSXIO_WRITE_FULLNAME);
  return 0;
}
