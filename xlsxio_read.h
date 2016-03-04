#ifndef INCLUDED_XLSXIO_READ_H
#define INCLUDED_XLSXIO_READ_H

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

/*! \brief read handle for .xlsx object */
typedef struct xlsxio_read_data* xlsxioreadhandle;

/*! \brief open .xlsx file
 * \param  filename      path of .xlsx file to open
 * \return read handle for .xlsx object or NULL on error
 * \sa     xlsxioread_close()
 */
DLL_EXPORT_XLSXIO xlsxioreadhandle xlsxioread_open (const char* filename);

/*! \brief close .xlsx file
 * \param  handle        read handle for .xlsx object
 * \sa     xlsxioread_open()
 */
DLL_EXPORT_XLSXIO void xlsxioread_close (xlsxioreadhandle handle);

/*! \brief type of pointer to callback function for listing worksheets
 * \param  name          name of worksheet
 * \param  callbackdata  callback data passed to xlsxioread_list_sheets
 * \return zero to continue, non-zero to abort
 * \sa     xlsxioread_list_sheets()
 */
typedef int (*xlsxioread_list_sheets_callback_fn)(const char* name, void* callbackdata);

/*! \brief list worksheets in .xlsx file
 * \param  handle        read handle for .xlsx object
 * \param  callback      callback function called for each worksheet
 * \param  callbackdata  custom data as passed to quickmail_add_body_custom/quickmail_add_attachment_custom
 * \sa     xlsxioread_list_sheets_callback_fn
 */
DLL_EXPORT_XLSXIO void xlsxioread_list_sheets (xlsxioreadhandle handle, xlsxioread_list_sheets_callback_fn callback, void* callbackdata);

/*! \brief possible values for the flags parameter of xlsxioread_process_sheet()
 * \sa     xlsxioread_process_sheet()
 */
#define XLSXIOREAD_SKIP_EMPTY_ROWS      0x0001
#define XLSXIOREAD_SKIP_EMPTY_CELLS     0x0002
#define XLSXIOREAD_SKIP_NONE            0
#define XLSXIOREAD_SKIP_ALL             (XLSXIOREAD_SKIP_EMPTY_ROWS | XLSXIOREAD_SKIP_EMPTY_CELLS)

/*! \brief type of pointer to callback function for processing a worksheet cell value
 * \param  row           row number (first row is 1)
 * \param  col           column number (first column is 1)
 * \param  value         value of cell (note: formulas are not calculated)
 * \param  callbackdata  callback data passed to xlsxioread_process_sheet
 * \return zero to continue, non-zero to abort
 * \sa     xlsxioread_process_sheet()
 * \sa     xlsxioread_process_sheet_row_callback_fn
 */
typedef int (*xlsxioread_process_sheet_cell_callback_fn)(size_t row, size_t col, const char* value, void* callbackdata);

/*! \brief type of pointer to callback function for processing the end of a worksheet row
 * \param  row           row number (first row is 1)
 * \param  maxcol        maximum column number on this row (first column is 1)
 * \param  callbackdata  callback data passed to xlsxioread_process_sheet
 * \return zero to continue, non-zero to abort
 * \sa     xlsxioread_process_sheet()
 * \sa     xlsxioread_process_sheet_cell_callback_fn
 */
typedef int (*xlsxioread_process_sheet_row_callback_fn)(size_t row, size_t maxcol, void* callbackdata);

/*! \brief process all rows and columns of a worksheet in an .xlsx file
 * \param  handle        read handle for .xlsx object
 * \param  sheetname     worksheet name (NULL for first sheet)
 * \param  flags         XLSXIOREAD_SKIP_ flag(s) to determine how data is processed
 * \param  cell_callback callback function called for each cell
 * \param  row_callback  callback function called after each row
 * \param  callbackdata  callback data passed to xlsxioread_process_sheet
 * \sa     xlsxioread_process_sheet_row_callback_fn
 * \sa     xlsxioread_process_sheet_cell_callback_fn
 */
DLL_EXPORT_XLSXIO void xlsxioread_process_sheet (xlsxioreadhandle handle, const char* sheetname, unsigned int flags, xlsxioread_process_sheet_cell_callback_fn cell_callback, xlsxioread_process_sheet_row_callback_fn row_callback, void* callbackdata);

#ifdef __cplusplus
}
#endif

#endif
