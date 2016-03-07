/*****************************************************************************
Copyright (C)  2016  Brecht Sanders  All Rights Reserved

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
*****************************************************************************/

/**
 * @file xlsxio_write.h
 * @brief XLSX I/O header file for writing .xlsx files.
 * @author Brecht Sanders
 *
 * Include this header file to use XLSX I/O for writing .xlsx files and
 * link with -lxlsxio_write.
 */

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
