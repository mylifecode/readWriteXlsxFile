#include "xlsxio_read.h"
#include <stdlib.h>
#include <string.h>
#ifdef BUILD_XLSXIO_STATIC
#define ZIP_STATIC
#endif
#include <zip.h>
#include <expat.h>

////////////////////////////////////////////////////////////////////////

#define BUFFER_SIZE 256
//#define BUFFER_SIZE 4

//process XML file contents
int expat_process_zip_file (zip_t* zip, const char* filename, XML_StartElementHandler start_handler, XML_EndElementHandler end_handler, XML_CharacterDataHandler data_handler, void* callbackdata, XML_Parser* xmlparser)
{
  zip_file_t* zipfile;
  XML_Parser parser;
  char buf[BUFFER_SIZE];
  zip_int64_t buflen;
  enum XML_Status status = XML_STATUS_ERROR;
  if ((zipfile = zip_fopen(zip, filename, 0)) == NULL) {
    return -1;
  }
  parser = XML_ParserCreate(NULL);
  XML_SetUserData(parser, callbackdata);
  XML_SetElementHandler(parser, start_handler, end_handler);
  XML_SetCharacterDataHandler(parser, data_handler);
  if (xmlparser)
    *xmlparser = parser;
  while ((buflen = zip_fread(zipfile, buf, sizeof(buf))) >= 0) {
    if ((status = XML_Parse(parser, buf, buflen, (buflen < sizeof(buf) ? 1 : 0))) == XML_STATUS_ERROR)
      break;
  }
  if (status != XML_STATUS_ERROR)
    status = XML_Parse(parser, NULL, 0, 1);
  XML_ParserFree(parser);
  zip_fclose(zipfile);
  return (status == XML_STATUS_ERROR ? 1 : 0);
}

//get expat attribute by name, returns NULL if not found
const XML_Char* get_expat_attr_by_name (const XML_Char** atts, const XML_Char* name)
{
  const XML_Char** p = atts;
  while (*p) {
    if (stricmp(*p++, name) == 0)
      return *p;
    p++;
  }
  return NULL;
}

//generate .rels filename, returns NULL on error, caller must free result
char* get_relationship_filename (const char* filename)
{
  char* result;
  size_t filenamelen = strlen(filename);
  if ((result = (char*)malloc(filenamelen + 12))) {
    size_t i = filenamelen;
    while (i > 0) {
      if (filename[i - 1] == '/')
        break;
      i--;
    }
    memcpy(result, filename, i);
    memcpy(result + i, "_rels/", 6);
    memcpy(result + i + 6, filename + i, filenamelen - i);
    strcpy(result + filenamelen + 6, ".rels");
  }
  return result;
}

//join basepath and filename (caller must free result)
char* join_basepath_filename (const char* basepath, const char* filename)
{
  char* result = NULL;
  if (filename && *filename) {
    size_t basepathlen = (basepath ? strlen(basepath) : 0);
    size_t filenamelen = strlen(filename);
    if ((result = (char*)malloc(basepathlen + filenamelen + 1)) != NULL) {
      if (basepathlen > 0)
        memcpy(result, basepath, basepathlen);
      memcpy(result + basepathlen, filename, filenamelen);
      result[basepathlen + filenamelen] = 0;
    }
  }
  return result;
}

//determine column number based on cell coordinate (e.g. "A1"), returns 1-based column number or 0 on error
size_t get_col_nr (const char* A1col)
{
  const char* p = A1col;
  size_t result = 0;
  if (p) {
    while (*p) {
      if (*p >= 'A' && *p <= 'Z')
        result = result * 26 + (*p - 'A') + 1;
      else if (*p >= 'a' && *p <= 'z')
        result = result * 26 + (*p - 'a') + 1;
      else if (*p >= '0' && *p <= '9' && p != A1col)
        return result;
      else
        break;
      p++;
    }
  }
  return 0;
}

//determine row number based on cell coordinate (e.g. "A1"), returns 1-based row number or 0 on error
size_t get_row_nr (const char* A1col)
{
  const char* p = A1col;
  size_t result = 0;
  if (p) {
    while (*p) {
      if ((*p >= 'A' && *p <= 'Z') || (*p >= 'a' && *p <= 'z'))
        ;
      else if (*p >= '0' && *p <= '9' && p != A1col)
        result = result * 10 + (*p - '0');
      else
        return 0;
      p++;
    }
  }
  return result;
}

////////////////////////////////////////////////////////////////////////

struct sharedstringlist {
  char** strings;
  size_t count;
};

struct sharedstringlist* sharedstringlist_create ()
{
  struct sharedstringlist* sharedstrings;
  if ((sharedstrings = (struct sharedstringlist*)malloc(sizeof(struct sharedstringlist))) != NULL) {
    sharedstrings->strings = NULL;
    sharedstrings->count = 0;
  }
  return sharedstrings;
}

void sharedstringlist_destroy (struct sharedstringlist* sharedstrings)
{
  if (sharedstrings) {
    size_t i;
    for (i = 0; i < sharedstrings->count; i++)
      free(sharedstrings->strings[i]);
    free(sharedstrings);
  }
}

size_t sharedstringlist_size (struct sharedstringlist* sharedstrings)
{
  return sharedstrings->count;
}

int sharedstringlist_add_buffer (struct sharedstringlist* sharedstrings, const char* data, size_t datalen)
{
  char* s;
  char** p;
  if (!sharedstrings)
    return 1;
  if (!data) {
    s = NULL;
  } else {
    if ((s = (char*)malloc(datalen + 1)) == NULL)
      return 2;
    memcpy(s, data, datalen);
    s[datalen] = 0;
  }
  if ((p = (char**)realloc(sharedstrings->strings, (sharedstrings->count + 1) * sizeof(sharedstrings->strings[0]))) == NULL)
    return 3;
  sharedstrings->strings = p;
  sharedstrings->strings[sharedstrings->count++] = s;
  return 0;
}

int sharedstringlist_add_string (struct sharedstringlist* sharedstrings, const char* data)
{
  return sharedstringlist_add_buffer(sharedstrings, data, (data ? strlen(data) : 0));
}

const char* sharedstringlist_get (struct sharedstringlist* sharedstrings, size_t index)
{
  if (!sharedstrings || index >= sharedstrings->count)
    return NULL;
  return sharedstrings->strings[index];
}

////////////////////////////////////////////////////////////////////////

struct shared_strings_callback_data {
  XML_Parser xmlparser;
  zip_file_t* zipfile;
  struct sharedstringlist* sharedstrings;
  int insst;
  int insi;
  int intext;
  char* text;
  size_t textlen;
};

void shared_strings_callback_find_sharedstringtable_start (void* callbackdata, const XML_Char* name, const XML_Char** atts);
void shared_strings_callback_find_sharedstringtable_end (void* callbackdata, const XML_Char* name);
void shared_strings_callback_find_shared_stringitem_start (void* callbackdata, const XML_Char* name, const XML_Char** atts);
void shared_strings_callback_find_shared_stringitem_end (void* callbackdata, const XML_Char* name);
void shared_strings_callback_find_shared_string_start (void* callbackdata, const XML_Char* name, const XML_Char** atts);
void shared_strings_callback_find_shared_string_end (void* callbackdata, const XML_Char* name);
void shared_strings_callback_string_data (void* callbackdata, const XML_Char* buf, int buflen);

void shared_strings_callback_find_sharedstringtable_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct shared_strings_callback_data* data = (struct shared_strings_callback_data*)callbackdata;
  if (stricmp(name, "sst") == 0) {
    XML_SetElementHandler(data->xmlparser, shared_strings_callback_find_shared_stringitem_start, NULL);
  }
}

void shared_strings_callback_find_sharedstringtable_end (void* callbackdata, const XML_Char* name)
{
  struct shared_strings_callback_data* data = (struct shared_strings_callback_data*)callbackdata;
  if (stricmp(name, "sst") == 0) {
    XML_SetElementHandler(data->xmlparser, shared_strings_callback_find_sharedstringtable_start, NULL);
  }
}

void shared_strings_callback_find_shared_stringitem_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct shared_strings_callback_data* data = (struct shared_strings_callback_data*)callbackdata;
  if (stricmp(name, "si") == 0) {
    XML_SetElementHandler(data->xmlparser, shared_strings_callback_find_shared_string_start, shared_strings_callback_find_sharedstringtable_end);
  }
}

void shared_strings_callback_find_shared_stringitem_end (void* callbackdata, const XML_Char* name)
{
  struct shared_strings_callback_data* data = (struct shared_strings_callback_data*)callbackdata;
  if (stricmp(name, "si") == 0) {
    XML_SetElementHandler(data->xmlparser, shared_strings_callback_find_shared_stringitem_start, shared_strings_callback_find_sharedstringtable_end);
  } else {
    shared_strings_callback_find_sharedstringtable_end(callbackdata, name);
  }
}

void shared_strings_callback_find_shared_string_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct shared_strings_callback_data* data = (struct shared_strings_callback_data*)callbackdata;
  if (stricmp(name, "t") == 0) {
    XML_SetElementHandler(data->xmlparser, NULL, shared_strings_callback_find_shared_string_end);
    XML_SetCharacterDataHandler(data->xmlparser, shared_strings_callback_string_data);
  }
}

void shared_strings_callback_find_shared_string_end (void* callbackdata, const XML_Char* name)
{
  struct shared_strings_callback_data* data = (struct shared_strings_callback_data*)callbackdata;
  if (stricmp(name, "t") == 0) {
    sharedstringlist_add_buffer(data->sharedstrings, data->text, data->textlen);
    if (data->text)
      free(data->text);
    data->text = NULL;
    data->textlen = 0;
    XML_SetElementHandler(data->xmlparser, shared_strings_callback_find_shared_string_start, shared_strings_callback_find_shared_stringitem_end);
    XML_SetCharacterDataHandler(data->xmlparser, NULL);
  } else {
    shared_strings_callback_find_shared_stringitem_end(callbackdata, name);
  }
}

void shared_strings_callback_string_data (void* callbackdata, const XML_Char* buf, int buflen)
{
  struct shared_strings_callback_data* data = (struct shared_strings_callback_data*)callbackdata;
  if ((data->text = (char*)realloc(data->text, data->textlen + buflen)) == NULL) {
    //memory allocation error
    data->textlen = 0;
  } else {
    memcpy(data->text + data->textlen, buf, buflen);
    data->textlen += buflen;
  }
}

////////////////////////////////////////////////////////////////////////

struct xlsxio_read_data {
  zip_t* zip;
};

DLL_EXPORT_XLSXIO xlsxioreadhandle xlsxioread_open (const char* filename)
{
  struct xlsxio_read_data* result;
  if ((result = (struct xlsxio_read_data*)malloc(sizeof(struct xlsxio_read_data))) != NULL) {
    if ((result->zip = zip_open(filename, ZIP_RDONLY, NULL)) == NULL) {
      free(result);
      return NULL;
    }
  }
  return result;
}

DLL_EXPORT_XLSXIO void xlsxioread_close (xlsxioreadhandle handle)
{
  if (handle)
    zip_close(handle->zip);
}

////////////////////////////////////////////////////////////////////////

//callback function definition for use with list_files_by_contenttype
typedef void (*contenttype_file_callback_fn)(zip_t* zip, const char* filename, const char* contenttype, void* callbackdata);

struct list_files_by_contenttype_callback_data {
  /*XML_Parser xmlparser;*/
  zip_t* zip;
  const char* contenttype;
  contenttype_file_callback_fn filecallbackfn;
  void* filecallbackdata;
};

//expat callback function for element start used by list_files_by_contenttype
void list_files_by_contenttype_expat_callback_element_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct list_files_by_contenttype_callback_data* data = (struct list_files_by_contenttype_callback_data*)callbackdata;
  if (stricmp(name, "Override") == 0) {
    //explicitly specified file
    const XML_Char* contenttype;
    const XML_Char* partname;
    if ((contenttype = get_expat_attr_by_name(atts, "ContentType")) != NULL && stricmp(contenttype, data->contenttype) == 0) {
      if ((partname = get_expat_attr_by_name(atts, "PartName")) != NULL) {
        if (partname[0] == '/')
          partname++;
        data->filecallbackfn(data->zip, partname, contenttype, data->filecallbackdata);
      }
    }
  } else if (stricmp(name, "Default") == 0) {
    //by extension
    const XML_Char* contenttype;
    const XML_Char* extension;
    if ((contenttype = get_expat_attr_by_name(atts, "ContentType")) != NULL && stricmp(contenttype, data->contenttype) == 0) {
      if ((extension = get_expat_attr_by_name(atts, "Extension")) != NULL) {
        const char* filename;
        size_t filenamelen;
        zip_int64_t i;
        zip_int64_t zipnumfiles = zip_get_num_entries(data->zip, 0);
        size_t extensionlen = strlen(extension);
        for (i = 0; i < zipnumfiles; i++) {
          filename = zip_get_name(data->zip, i, ZIP_FL_ENC_GUESS);
          filenamelen = strlen(filename);
          if (filenamelen > extensionlen && filename[filenamelen - extensionlen - 1] == '.' && stricmp(filename + filenamelen - extensionlen, extension) == 0) {
            data->filecallbackfn(data->zip, filename, contenttype, data->filecallbackdata);
          }
        }
      }
    }
  }
}

//list file names by content type
void list_files_by_contenttype (zip_t* zip, const char* contenttype, contenttype_file_callback_fn filecallbackfn, void* filecallbackdata)
{
  struct list_files_by_contenttype_callback_data callbackdata = {
    /*.xmlparser = NULL,*/
    .zip = zip,
    .contenttype = contenttype,
    .filecallbackfn = filecallbackfn,
    .filecallbackdata = filecallbackdata
  };
  expat_process_zip_file(zip, "[Content_Types].xml", list_files_by_contenttype_expat_callback_element_start, NULL, NULL, &callbackdata, NULL/*&callbackdata.xmlparser*/);
}

////////////////////////////////////////////////////////////////////////

//callback structure used by main_sheet_list_expat_callback_element_start
struct main_sheet_list_callback_data {
  XML_Parser xmlparser;
  xlsxioread_list_sheets_callback_fn callback;
  void* callbackdata;
};

//callback used by xlsxioread_list_sheets
void main_sheet_list_expat_callback_element_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct main_sheet_list_callback_data* data = (struct main_sheet_list_callback_data*)callbackdata;
  if (data && data->callback) {
    if (stricmp(name, "sheet") == 0) {
      const XML_Char* sheetname;
      //const XML_Char* relid = get_expat_attr_by_name(atts, "r:id");
      if ((sheetname = get_expat_attr_by_name(atts, "name")) != NULL)
        if (data->callback) {
          if ((*data->callback)(sheetname, data->callbackdata) != 0) {
            XML_StopParser(data->xmlparser, XML_FALSE);
            return;
          }
        }
    }
  }
}

//process contents each sheet listed in main sheet
void xlsxioread_list_sheets_callback (zip_t* zip, const char* filename, const char* contenttype, void* callbackdata)
{
  //get sheet information from file
  expat_process_zip_file(zip, filename, main_sheet_list_expat_callback_element_start, NULL, NULL, callbackdata, &((struct main_sheet_list_callback_data*)callbackdata)->xmlparser);
}

//list all worksheets
DLL_EXPORT_XLSXIO void xlsxioread_list_sheets (xlsxioreadhandle handle, xlsxioread_list_sheets_callback_fn callback, void* callbackdata)
{
  //process contents of main sheet
  struct main_sheet_list_callback_data sheetcallbackdata = {
    .xmlparser = NULL,
    .callback = callback,
    .callbackdata = callbackdata
  };
  list_files_by_contenttype(handle->zip, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", xlsxioread_list_sheets_callback, &sheetcallbackdata);
}

////////////////////////////////////////////////////////////////////////

//callback data structure used by main_sheet_get_sheetfile_callback
struct main_sheet_get_rels_callback_data {
  XML_Parser xmlparser;
  const char* sheetname;
  char* basepath;
  char* sheetrelid;
  char* sheetfile;
  char* sharedstringsfile;
};

//determine relationship id for specific sheet name
void main_sheet_get_relid_expat_callback_element_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct main_sheet_get_rels_callback_data* data = (struct main_sheet_get_rels_callback_data*)callbackdata;
  if (stricmp(name, "sheet") == 0) {
    const XML_Char* name = get_expat_attr_by_name(atts, "name");
    if (!data->sheetname || stricmp(name, data->sheetname) == 0) {
      const XML_Char* relid = get_expat_attr_by_name(atts, "r:id");
      if (relid && *relid) {
        data->sheetrelid = strdup(relid);
        XML_StopParser(data->xmlparser, XML_FALSE);
        return;
      }
    }
  }
}

//determine sheet file name for specific relationship id
void main_sheet_get_sheetfile_expat_callback_element_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct main_sheet_get_rels_callback_data* data = (struct main_sheet_get_rels_callback_data*)callbackdata;
  if (data->sheetrelid) {
    if (stricmp(name, "Relationship") == 0) {
      const XML_Char* reltype;
      if ((reltype = get_expat_attr_by_name(atts, "Type")) != NULL && stricmp(reltype, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet") == 0) {
        const XML_Char* relid = get_expat_attr_by_name(atts, "Id");
        if (stricmp(relid, data->sheetrelid) == 0) {
          const XML_Char* filename = get_expat_attr_by_name(atts, "Target");
          if (filename && *filename) {
            data->sheetfile = join_basepath_filename(data->basepath, filename);
          }
        }
      } else if ((reltype = get_expat_attr_by_name(atts, "Type")) != NULL && stricmp(reltype, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings") == 0) {
        const XML_Char* filename = get_expat_attr_by_name(atts, "Target");
        if (filename && *filename) {
          data->sharedstringsfile = join_basepath_filename(data->basepath, filename);
        }
      }
    }
  }
}

//determine the file name for a specified sheet name
void main_sheet_get_sheetfile_callback (zip_t* zip, const char* filename, const char* contenttype, void* callbackdata)
{
  struct main_sheet_get_rels_callback_data* data = (struct main_sheet_get_rels_callback_data*)callbackdata;
  if (!data->sheetrelid) {
    expat_process_zip_file(zip, filename, main_sheet_get_relid_expat_callback_element_start, NULL, NULL, callbackdata, &data->xmlparser);
  }
  if (data->sheetrelid) {
    char* relfilename;
    //determine base name (including trailing slash)
    size_t i = strlen(filename);
    while (i > 0) {
      if (filename[i - 1] == '/')
        break;
      i--;
    }
    if (data->basepath)
      free(data->basepath);
    data->basepath = (char*)malloc(i + 1);
    memcpy(data->basepath, filename, i);
    data->basepath[i] = 0;
    //find sheet filename in relationship contents
    if ((relfilename = get_relationship_filename(filename)) != NULL) {
      expat_process_zip_file(zip, relfilename, main_sheet_get_sheetfile_expat_callback_element_start, NULL, NULL, callbackdata, &data->xmlparser);
      free(relfilename);
    } else {
      free(data->sheetrelid);
      data->sheetrelid = NULL;
    }
  }
}

////////////////////////////////////////////////////////////////////////

typedef enum {
  none,
  value_string,
  inline_string,
  shared_string
} cell_string_type_enum;

struct data_sheet_callback_data {
  XML_Parser xmlparser;
  struct sharedstringlist* sharedstrings;
  size_t rownr;
  size_t colnr;
  size_t cols;
  char* celldata;
  size_t celldatalen;
  cell_string_type_enum cell_string_type;
  unsigned int flags;
  xlsxioread_process_sheet_row_callback_fn sheet_row_callback;
  xlsxioread_process_sheet_cell_callback_fn sheet_cell_callback;
  void* callbackdata;
};

void data_sheet_expat_callback_find_worksheet_start (void* callbackdata, const XML_Char* name, const XML_Char** atts);
void data_sheet_expat_callback_find_worksheet_end (void* callbackdata, const XML_Char* name);
void data_sheet_expat_callback_find_sheetdata_start (void* callbackdata, const XML_Char* name, const XML_Char** atts);
void data_sheet_expat_callback_find_sheetdata_end (void* callbackdata, const XML_Char* name);
void data_sheet_expat_callback_find_row_start (void* callbackdata, const XML_Char* name, const XML_Char** atts);
void data_sheet_expat_callback_find_row_end (void* callbackdata, const XML_Char* name);
void data_sheet_expat_callback_find_cell_start (void* callbackdata, const XML_Char* name, const XML_Char** atts);
void data_sheet_expat_callback_find_cell_end (void* callbackdata, const XML_Char* name);
void data_sheet_expat_callback_find_value_start (void* callbackdata, const XML_Char* name, const XML_Char** atts);
void data_sheet_expat_callback_find_value_end (void* callbackdata, const XML_Char* name);
void data_sheet_expat_callback_value_data (void* callbackdata, const XML_Char* buf, int buflen);

void data_sheet_expat_callback_find_worksheet_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (stricmp(name, "worksheet") == 0) {
    XML_SetElementHandler(data->xmlparser, data_sheet_expat_callback_find_sheetdata_start, NULL);
  }
}

void data_sheet_expat_callback_find_worksheet_end (void* callbackdata, const XML_Char* name)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (stricmp(name, "worksheet") == 0) {
    XML_SetElementHandler(data->xmlparser, data_sheet_expat_callback_find_worksheet_start, NULL);
  }
}

void data_sheet_expat_callback_find_sheetdata_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (stricmp(name, "sheetData") == 0) {
    XML_SetElementHandler(data->xmlparser, data_sheet_expat_callback_find_row_start, data_sheet_expat_callback_find_sheetdata_end);
  }
}

void data_sheet_expat_callback_find_sheetdata_end (void* callbackdata, const XML_Char* name)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (stricmp(name, "sheetData") == 0) {
    XML_SetElementHandler(data->xmlparser, data_sheet_expat_callback_find_sheetdata_start, data_sheet_expat_callback_find_worksheet_end);
  } else {
    data_sheet_expat_callback_find_worksheet_end(callbackdata, name);
  }
}

void data_sheet_expat_callback_find_row_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (stricmp(name, "row") == 0) {
    data->rownr++;
    data->colnr = 1;
    XML_SetElementHandler(data->xmlparser, data_sheet_expat_callback_find_cell_start, data_sheet_expat_callback_find_row_end);
  }
}

void data_sheet_expat_callback_find_row_end (void* callbackdata, const XML_Char* name)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (stricmp(name, "row") == 0) {
    //determine number of columns based on first row
    if (data->rownr == 1 && data->cols == 0)
      data->cols = data->colnr - 1;
    //add empty columns if needed
    if (!(data->flags | XLSXIOREAD_SKIP_EMPTY_CELLS) && data->sheet_cell_callback) {
      while (data->colnr <= data->cols) {
        if ((*data->sheet_cell_callback)(data->rownr, data->colnr, NULL, data->callbackdata)) {
          XML_StopParser(data->xmlparser, XML_FALSE);
          return;
        }
        data->colnr++;
      }
    }
    //process end of row
    if (data->sheet_row_callback) {
      if ((*data->sheet_row_callback)(data->rownr, data->colnr - 1, callbackdata)) {
        XML_StopParser(data->xmlparser, XML_FALSE);
        return;
      }
    }
    XML_SetElementHandler(data->xmlparser, data_sheet_expat_callback_find_row_start, data_sheet_expat_callback_find_sheetdata_end);
  } else {
    data_sheet_expat_callback_find_sheetdata_end(callbackdata, name);
  }
}

void data_sheet_expat_callback_find_cell_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (stricmp(name, "c") == 0) {
    const XML_Char* t = get_expat_attr_by_name(atts, "r");
    size_t cellcolnr = get_col_nr(t);
    //insert empty rows if needed
    if (data->colnr == 1) {
      size_t cellrownr = get_row_nr(t);
      if (data->flags | XLSXIOREAD_SKIP_EMPTY_ROWS) {
        data->rownr = cellrownr;
      } else {
        while (data->rownr < cellrownr) {
          //insert empty columns
          if (!(data->flags | XLSXIOREAD_SKIP_EMPTY_CELLS) && data->sheet_cell_callback) {
            while (data->colnr <= data->cols) {
              if ((*data->sheet_cell_callback)(data->rownr, data->colnr, NULL, data->callbackdata)) {
                XML_StopParser(data->xmlparser, XML_FALSE);
                return;
              }
              data->colnr++;
            }
          }
          //finish empty row
          if (data->sheet_row_callback) {
            if ((*data->sheet_row_callback)(data->rownr, data->cols, callbackdata)) {
              XML_StopParser(data->xmlparser, XML_FALSE);
              return;
            }
          }
          data->rownr++;
          data->colnr = 1;
        }
      }
    }
    //insert empty columns if needed
    if (data->flags | XLSXIOREAD_SKIP_EMPTY_CELLS) {
      data->colnr = cellcolnr;
    } else {
      while (data->colnr < cellcolnr) {
        if (data->sheet_cell_callback) {
          if ((*data->sheet_cell_callback)(data->rownr, data->colnr, NULL, data->callbackdata)) {
            XML_StopParser(data->xmlparser, XML_FALSE);
            return;
          }
        }
        data->colnr++;
      }
    }
    //determing value type
    if ((t = get_expat_attr_by_name(atts, "t")) != NULL && stricmp(t, "s") == 0)
      data->cell_string_type = shared_string;
    else
      data->cell_string_type = value_string;
    //prepare empty value data
    free(data->celldata);
    data->celldata = NULL;
    data->celldatalen = 0;
    XML_SetElementHandler(data->xmlparser, data_sheet_expat_callback_find_value_start, data_sheet_expat_callback_find_cell_end);
  }
}

void data_sheet_expat_callback_find_cell_end (void* callbackdata, const XML_Char* name)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (stricmp(name, "c") == 0) {
    const char* s = NULL;
    if (data->celldata) {
      data->celldata[data->celldatalen] = 0;
      if (data->cell_string_type == shared_string) {
        char* p = NULL;
        long num = strtol(data->celldata, &p, 10);
        if (!p || (p != data->celldata && *p == 0)) {
          s = sharedstringlist_get(data->sharedstrings, num);
        }
      } else if (data->cell_string_type != none) {
        s = data->celldata;
      }
    }
    //process data
    if (data->sheet_cell_callback) {
      if ((*data->sheet_cell_callback)(data->rownr, data->colnr, s, data->callbackdata)) {
        XML_StopParser(data->xmlparser, XML_FALSE);
        return;
      }
    }
    data->colnr++;
    //reset data
    data->cell_string_type = none;
    free(data->celldata);
    data->celldata = NULL;
    data->celldatalen = 0;
    XML_SetElementHandler(data->xmlparser, data_sheet_expat_callback_find_cell_start, data_sheet_expat_callback_find_row_end);
    XML_SetCharacterDataHandler(data->xmlparser, NULL);
  } else {
    data_sheet_expat_callback_find_row_end(callbackdata, name);
  }
}

void data_sheet_expat_callback_find_value_start (void* callbackdata, const XML_Char* name, const XML_Char** atts)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (stricmp(name, "v") == 0 || stricmp(name, "t") == 0) {
    XML_SetElementHandler(data->xmlparser, NULL, data_sheet_expat_callback_find_value_end);
    XML_SetCharacterDataHandler(data->xmlparser, data_sheet_expat_callback_value_data);
  } if (stricmp(name, "is") == 0) {
    data->cell_string_type = inline_string;
  }
}

void data_sheet_expat_callback_find_value_end (void* callbackdata, const XML_Char* name)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (stricmp(name, "v") == 0 || stricmp(name, "t") == 0) {
    XML_SetElementHandler(data->xmlparser, data_sheet_expat_callback_find_value_start, data_sheet_expat_callback_find_cell_end);
    XML_SetCharacterDataHandler(data->xmlparser, NULL);
  } if (stricmp(name, "is") == 0) {
    data->cell_string_type = none;
  } else {
    data_sheet_expat_callback_find_row_end(callbackdata, name);
  }
}

void data_sheet_expat_callback_value_data (void* callbackdata, const XML_Char* buf, int buflen)
{
  struct data_sheet_callback_data* data = (struct data_sheet_callback_data*)callbackdata;
  if (data->cell_string_type != none) {
    if ((data->celldata = (char*)realloc(data->celldata, data->celldatalen + buflen + 1)) == NULL) {
      //memory allocation error
      data->celldatalen = 0;
    } else {
      //add new data to value buffer
      memcpy(data->celldata + data->celldatalen, buf, buflen);
      data->celldatalen += buflen;
    }
  }
}

////////////////////////////////////////////////////////////////////////

DLL_EXPORT_XLSXIO void xlsxioread_process_sheet (xlsxioreadhandle handle, const char* sheetname, unsigned int flags, xlsxioread_process_sheet_cell_callback_fn cell_callback, xlsxioread_process_sheet_row_callback_fn row_callback, void* callbackdata)
{
  //determine sheet file name
  struct main_sheet_get_rels_callback_data getrelscallbackdata = {
    .sheetname = sheetname,
    .basepath = NULL,
    .sheetrelid = NULL,
    .sheetfile = NULL,
    .sharedstringsfile = NULL
  };

  list_files_by_contenttype(handle->zip, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", main_sheet_get_sheetfile_callback, &getrelscallbackdata);

  //process shared strings
  struct sharedstringlist* sharedstrings = sharedstringlist_create();
  struct shared_strings_callback_data sharedstringsdata = {
    .xmlparser = NULL,
    .zipfile = NULL,
    .sharedstrings = sharedstrings,
    .insst = 0,
    .insi = 0,
    .intext = 0,
    .text = NULL,
    .textlen = 0
  };
  expat_process_zip_file(handle->zip, getrelscallbackdata.sharedstringsfile, shared_strings_callback_find_sharedstringtable_start, NULL, NULL, &sharedstringsdata, &sharedstringsdata.xmlparser);
  free(sharedstringsdata.text);

  //process sheet
  struct data_sheet_callback_data processcallbackdata = {
    .xmlparser = NULL,
    .sharedstrings = sharedstrings,
    .rownr = 0,
    .colnr = 0,
    .cols = 0,
    .celldata = NULL,
    .celldatalen = 0,
    .cell_string_type = none,
    .flags = flags,
    .sheet_row_callback = row_callback,
    .sheet_cell_callback = cell_callback,
    .callbackdata = callbackdata
  };
  expat_process_zip_file(handle->zip, getrelscallbackdata.sheetfile, data_sheet_expat_callback_find_worksheet_start, NULL, NULL, &processcallbackdata, &processcallbackdata.xmlparser);

  //clean up
  free(processcallbackdata.celldata);
  sharedstringlist_destroy(sharedstrings);
  free(getrelscallbackdata.basepath);
  free(getrelscallbackdata.sheetrelid);
  free(getrelscallbackdata.sheetfile);
  free(getrelscallbackdata.sharedstringsfile);
}
