#ifndef INCLUDED_XLSXIO_WRITE_H
#define INCLUDED_XLSXIO_WRITE_H

#include <stdlib.h>

#ifndef DLL_EXPORT_XLSXIO
#ifdef _WIN32
#if defined(BUILD_XLSXIO_DLL)
#define DLL_EXPORT_XLSXIO __declspec(dllexport)
#elif !defined(STATIC) && !defined(BUILD_XLSXIO_STATIC)
#define DLL_EXPORT_XLSXIO __declspec(dllimport)
#else
#define DLL_EXPORT_XLSXIO
#endif
#else
#define DLL_EXPORT_XLSXIO
#endif
#endif

#ifdef __cplusplus
extern "C" {
#endif

/*! \brief write handle for .xlsx object */
typedef struct xlsxio_write_struct* xlsxiowriter;

/*! \brief open .xlsx file
 * \param  filename      path of .xlsx file to open
 * \return write handle for .xlsx object or NULL on error
 * \sa     xlsxiowrite_close()
 */
DLL_EXPORT_XLSXIO xlsxiowriter xlsxiowrite_open (const char* filename);

/*! \brief close .xlsx file
 * \param  handle        write handle for .xlsx object
 * \return zero on success, non-zero on error
 * \sa     xlsxiowrite_open()
 */
DLL_EXPORT_XLSXIO int xlsxiowrite_close (xlsxiowriter handle);

/*! \brief add a cell with string data
 * \param  handle        write handle for .xlsx object
 * \param  value         string value
 * \sa     xlsxiowrite_next_row()
 */
DLL_EXPORT_XLSXIO void xlsxiowrite_add_cell_string (xlsxiowriter handle, const char* value);

/*! \brief mark the end of a row (next cell will start on a new row)
 * \param  handle        write handle for .xlsx object
 * \sa     xlsxiowrite_add_cell_string()
 */
DLL_EXPORT_XLSXIO void xlsxiowrite_next_row (xlsxiowriter handle);

#ifdef __cplusplus
}
#endif

#endif
