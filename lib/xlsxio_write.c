#include "xlsxio_write.h"
#include "xlsxio_version.h"
#include <stdlib.h>
#include <string.h>
#include <time.h>
#include <unistd.h>
#include <fcntl.h>
#ifdef BUILD_XLSXIO_STATIC
#define ZIP_STATIC
#endif
#include <zip.h>
#if defined(_WIN32) && !defined(USE_PTHREADS)
#define USE_WINTHREADS
#include <windows.h>
#else
#define USE_PTHREADS
#include <pthread.h>
#endif

#ifndef ZIP_RDONLY
typedef struct zip zip_t;
typedef struct zip_source zip_source_t;
#endif

#ifdef _WIN32
#define pipe(fds) _pipe(fds, 4096, _O_BINARY)
#define read _read
#define write _write
#define write _write
#define close _close
#define fdopen _fdopen
#else
#define _fdopen(f) f
#endif

DLL_EXPORT_XLSXIO void xlsxiowrite_get_version (int* major, int* minor, int* micro)
{
  if (major)
    *major = XLSXIO_VERSION_MAJOR;
  if (minor)
    *minor = XLSXIO_VERSION_MINOR;
  if (micro)
    *micro = XLSXIO_VERSION_MICRO;
}

DLL_EXPORT_XLSXIO const char* xlsxiowrite_get_version_string ()
{
  return XLSXIO_VERSION_STRING;
}

////////////////////////////////////////////////////////////////////////

int zip_add_static_content_buffer (zip_t* zip, const char* filename, const char* buf, size_t buflen)
{
  zip_source_t* zipsrc;
  if ((zipsrc = zip_source_buffer(zip, buf, buflen, 0)) == NULL) {
    fprintf(stderr, "Error creating file \"%s\" inside zip file\n", filename);/////
    return 1;
  }
  if (zip_file_add(zip, filename, zipsrc, ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8) < 0) {
    fprintf(stderr, "Error in zip_file_add\n");/////
    zip_source_free(zipsrc);
    return 2;
  }
  return 0;
}

int zip_add_static_content_string (zip_t* zip, const char* filename, const char* data)
{
  return zip_add_static_content_buffer(zip, filename, data, strlen(data));
}

////////////////////////////////////////////////////////////////////////

const char* content_types_xml =
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
  "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
  //"<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>"
  //"<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
  "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
  "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
  "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
  "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>"
  "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>"
  "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
  //"<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>"
  "</Types>";

const char* docprops_core_xml =
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
  "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\">"
  "</cp:coreProperties>";

const char* docprops_app_xml =
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
  "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">"
  "</Properties>";

const char* rels_xml =
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
  "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
  "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>"
  "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>"
  "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
  "</Relationships>";

/*
const char* sharedstrings_xml =
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
  "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"35\" uniqueCount=\"34\">\r\n"
  "<si><t>Name</t></si><si><t>Expires</t></si><si><t>LastLogon</t></si><si><t>LastSetPassword</t></si><si><t>LastBadPassword</t></si><si><t>LockOut</t></si><si><t>Active</t></si><si><t>Created</t></si><si><t>LastChanged</t></si><si><t>Login</t></si><si><t>Logons</t></si><si><t>ADName</t></si><si><t>FirstName</t></si><si><t>LastName</t></si><si><t>PrimaryEmail</t></si><si><t>HomeDirectory</t></si><si><t>Notes</t></si><si><t>PasswordExpires</t></si><si><t/></si><si><t>Never</t></si><si><t>Firefighter (Realdolmen)</t></si><si><t>2016-02-28 00:00:00</t></si><si><t>2016-02-17 12:00:00</t></si><si><t>2015-12-24 09:41:29</t></si><si><t>2016-02-01 11:50:11</t></si><si><t>Enabled+Expired</t></si><si><t>2013-02-27 10:30:55</t></si><si><t>2016-02-26 08:51:27</t></si><si><t>zxBeweRD008_f</t></si><si><t>CN=Firefighter (Realdolmen),OU=Admin Accounts,OU=Domain Security,DC=ISIS,DC=LOCAL</t></si><si><t>Realdolmen</t></si><si><t>Firefighter</t></si><si><t>Realdolmen.Firefighter@alpro.com</t></si><si><t>consultant RealDolmen - managed services - no often used</t></si>\r\n"
  "</sst>\r\n";

const char* styles_xml =
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
  "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\r\n"
  "<fonts count=\"19\"><font><sz val=\"10\"/><name val=\"Courier New\"/><family val=\"2\"/></font><font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font><font><b/><sz val=\"18\"/><color theme=\"3\"/><name val=\"Cambria\"/><family val=\"2\"/><scheme val=\"major\"/></font><font><b/><sz val=\"15\"/><color theme=\"3\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font><font><b/><sz val=\"13\"/><color theme=\"3\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font><font><b/><sz val=\"11\"/><color theme=\"3\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font><font><sz val=\"11\"/><color rgb=\"FF006100\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font><font><sz val=\"11\"/><color rgb=\"FF9C0006\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font><font><sz val=\"11\"/><color rgb=\"FF9C6500\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font><font><sz val=\"11\"/><color rgb=\"FF3F3F76\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font><font><b/><sz val=\"11\"/><color rgb=\"FF3F3F3F\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font><font><b/><sz val=\"11\"/><color rgb=\"FFFA7D00\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font><font><sz val=\"11\"/><color rgb=\"FFFA7D00\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font><font><b/><sz val=\"11\"/><color theme=\"0\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font><font><sz val=\"11\"/><color rgb=\"FFFF0000\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font><font><i/><sz val=\"11\"/><color rgb=\"FF7F7F7F\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font><font><b/><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font><font><sz val=\"11\"/><color theme=\"0\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font><font><b/><sz val=\"10\"/><name val=\"Courier New\"/><family val=\"2\"/></font></fonts>"
  "<fills count=\"33\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill><fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFC6EFCE\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFFC7CE\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFFEB9C\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFFCC99\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFF2F2F2\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFA5A5A5\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFFFFCC\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"4\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"4\" tint=\"0.79998168889431442\"/><bgColor indexed=\"65\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"4\" tint=\"0.59999389629810485\"/><bgColor indexed=\"65\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"4\" tint=\"0.39997558519241921\"/><bgColor indexed=\"65\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"5\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"5\" tint=\"0.79998168889431442\"/><bgColor indexed=\"65\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"5\" tint=\"0.59999389629810485\"/><bgColor indexed=\"65\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"5\" tint=\"0.39997558519241921\"/><bgColor indexed=\"65\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"6\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"6\" tint=\"0.79998168889431442\"/><bgColor indexed=\"65\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"6\" tint=\"0.59999389629810485\"/><bgColor indexed=\"65\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"6\" tint=\"0.39997558519241921\"/><bgColor indexed=\"65\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"7\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"7\" tint=\"0.79998168889431442\"/><bgColor indexed=\"65\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"7\" tint=\"0.59999389629810485\"/><bgColor indexed=\"65\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"7\" tint=\"0.39997558519241921\"/><bgColor indexed=\"65\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"8\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"8\" tint=\"0.79998168889431442\"/><bgColor indexed=\"65\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"8\" tint=\"0.59999389629810485\"/><bgColor indexed=\"65\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"8\" tint=\"0.39997558519241921\"/><bgColor indexed=\"65\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"9\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"9\" tint=\"0.79998168889431442\"/><bgColor indexed=\"65\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"9\" tint=\"0.59999389629810485\"/><bgColor indexed=\"65\"/></patternFill></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"9\" tint=\"0.39997558519241921\"/><bgColor indexed=\"65\"/></patternFill></fill></fills>"
  "<borders count=\"10\"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top/><bottom style=\"thick\"><color theme=\"4\"/></bottom><diagonal/></border><border><left/><right/><top/><bottom style=\"thick\"><color theme=\"4\" tint=\"0.499984740745262\"/></bottom><diagonal/></border><border><left/><right/><top/><bottom style=\"medium\"><color theme=\"4\" tint=\"0.39997558519241921\"/></bottom><diagonal/></border><border><left style=\"thin\"><color rgb=\"FF7F7F7F\"/></left><right style=\"thin\"><color rgb=\"FF7F7F7F\"/></right><top style=\"thin\"><color rgb=\"FF7F7F7F\"/></top><bottom style=\"thin\"><color rgb=\"FF7F7F7F\"/></bottom><diagonal/></border><border><left style=\"thin\"><color rgb=\"FF3F3F3F\"/></left><right style=\"thin\"><color rgb=\"FF3F3F3F\"/></right><top style=\"thin\"><color rgb=\"FF3F3F3F\"/></top><bottom style=\"thin\"><color rgb=\"FF3F3F3F\"/></bottom><diagonal/></border><border><left/><right/><top/><bottom style=\"double\"><color rgb=\"FFFF8001\"/></bottom><diagonal/></border><border><left style=\"double\"><color rgb=\"FF3F3F3F\"/></left><right style=\"double\"><color rgb=\"FF3F3F3F\"/></right><top style=\"double\"><color rgb=\"FF3F3F3F\"/></top><bottom style=\"double\"><color rgb=\"FF3F3F3F\"/></bottom><diagonal/></border><border><left style=\"thin\"><color rgb=\"FFB2B2B2\"/></left><right style=\"thin\"><color rgb=\"FFB2B2B2\"/></right><top style=\"thin\"><color rgb=\"FFB2B2B2\"/></top><bottom style=\"thin\"><color rgb=\"FFB2B2B2\"/></bottom><diagonal/></border><border><left/><right/><top style=\"thin\"><color theme=\"4\"/></top><bottom style=\"double\"><color theme=\"4\"/></bottom><diagonal/></border></borders>"
  "<cellStyleXfs count=\"42\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"><alignment vertical=\"top\"/></xf><xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"0\" applyNumberFormat=\"0\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"3\" fillId=\"0\" borderId=\"1\" applyNumberFormat=\"0\" applyFill=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"4\" fillId=\"0\" borderId=\"2\" applyNumberFormat=\"0\" applyFill=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"5\" fillId=\"0\" borderId=\"3\" applyNumberFormat=\"0\" applyFill=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"5\" fillId=\"0\" borderId=\"0\" applyNumberFormat=\"0\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"6\" fillId=\"2\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"7\" fillId=\"3\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"8\" fillId=\"4\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"9\" fillId=\"5\" borderId=\"4\" applyNumberFormat=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"10\" fillId=\"6\" borderId=\"5\" applyNumberFormat=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"11\" fillId=\"6\" borderId=\"4\" applyNumberFormat=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"12\" fillId=\"0\" borderId=\"6\" applyNumberFormat=\"0\" applyFill=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"13\" fillId=\"7\" borderId=\"7\" applyNumberFormat=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"14\" fillId=\"0\" borderId=\"0\" applyNumberFormat=\"0\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"8\" borderId=\"8\" applyNumberFormat=\"0\" applyFont=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"15\" fillId=\"0\" borderId=\"0\" applyNumberFormat=\"0\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"16\" fillId=\"0\" borderId=\"9\" applyNumberFormat=\"0\" applyFill=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"17\" fillId=\"9\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"10\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"11\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"17\" fillId=\"12\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"17\" fillId=\"13\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"14\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"15\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"17\" fillId=\"16\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"17\" fillId=\"17\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"18\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"19\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"17\" fillId=\"20\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"17\" fillId=\"21\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"22\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"23\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"17\" fillId=\"24\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"17\" fillId=\"25\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"26\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"27\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"17\" fillId=\"28\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"17\" fillId=\"29\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"30\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"31\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/><xf numFmtId=\"0\" fontId=\"17\" fillId=\"32\" borderId=\"0\" applyNumberFormat=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/></cellStyleXfs>"
  "<cellXfs count=\"2\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"><alignment vertical=\"top\"/></xf><xf numFmtId=\"0\" fontId=\"18\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyAlignment=\"1\"/></cellXfs>"
  "<cellStyles count=\"42\"><cellStyle name=\"20% - Accent1\" xfId=\"19\" builtinId=\"30\" customBuiltin=\"1\"/><cellStyle name=\"20% - Accent2\" xfId=\"23\" builtinId=\"34\" customBuiltin=\"1\"/><cellStyle name=\"20% - Accent3\" xfId=\"27\" builtinId=\"38\" customBuiltin=\"1\"/><cellStyle name=\"20% - Accent4\" xfId=\"31\" builtinId=\"42\" customBuiltin=\"1\"/><cellStyle name=\"20% - Accent5\" xfId=\"35\" builtinId=\"46\" customBuiltin=\"1\"/><cellStyle name=\"20% - Accent6\" xfId=\"39\" builtinId=\"50\" customBuiltin=\"1\"/><cellStyle name=\"40% - Accent1\" xfId=\"20\" builtinId=\"31\" customBuiltin=\"1\"/><cellStyle name=\"40% - Accent2\" xfId=\"24\" builtinId=\"35\" customBuiltin=\"1\"/><cellStyle name=\"40% - Accent3\" xfId=\"28\" builtinId=\"39\" customBuiltin=\"1\"/><cellStyle name=\"40% - Accent4\" xfId=\"32\" builtinId=\"43\" customBuiltin=\"1\"/><cellStyle name=\"40% - Accent5\" xfId=\"36\" builtinId=\"47\" customBuiltin=\"1\"/><cellStyle name=\"40% - Accent6\" xfId=\"40\" builtinId=\"51\" customBuiltin=\"1\"/><cellStyle name=\"60% - Accent1\" xfId=\"21\" builtinId=\"32\" customBuiltin=\"1\"/><cellStyle name=\"60% - Accent2\" xfId=\"25\" builtinId=\"36\" customBuiltin=\"1\"/><cellStyle name=\"60% - Accent3\" xfId=\"29\" builtinId=\"40\" customBuiltin=\"1\"/><cellStyle name=\"60% - Accent4\" xfId=\"33\" builtinId=\"44\" customBuiltin=\"1\"/><cellStyle name=\"60% - Accent5\" xfId=\"37\" builtinId=\"48\" customBuiltin=\"1\"/><cellStyle name=\"60% - Accent6\" xfId=\"41\" builtinId=\"52\" customBuiltin=\"1\"/><cellStyle name=\"Accent1\" xfId=\"18\" builtinId=\"29\" customBuiltin=\"1\"/><cellStyle name=\"Accent2\" xfId=\"22\" builtinId=\"33\" customBuiltin=\"1\"/><cellStyle name=\"Accent3\" xfId=\"26\" builtinId=\"37\" customBuiltin=\"1\"/><cellStyle name=\"Accent4\" xfId=\"30\" builtinId=\"41\" customBuiltin=\"1\"/><cellStyle name=\"Accent5\" xfId=\"34\" builtinId=\"45\" customBuiltin=\"1\"/><cellStyle name=\"Accent6\" xfId=\"38\" builtinId=\"49\" customBuiltin=\"1\"/><cellStyle name=\"Bad\" xfId=\"7\" builtinId=\"27\" customBuiltin=\"1\"/><cellStyle name=\"Calculation\" xfId=\"11\" builtinId=\"22\" customBuiltin=\"1\"/><cellStyle name=\"Check Cell\" xfId=\"13\" builtinId=\"23\" customBuiltin=\"1\"/><cellStyle name=\"Explanatory Text\" xfId=\"16\" builtinId=\"53\" customBuiltin=\"1\"/><cellStyle name=\"Good\" xfId=\"6\" builtinId=\"26\" customBuiltin=\"1\"/><cellStyle name=\"Heading 1\" xfId=\"2\" builtinId=\"16\" customBuiltin=\"1\"/><cellStyle name=\"Heading 2\" xfId=\"3\" builtinId=\"17\" customBuiltin=\"1\"/><cellStyle name=\"Heading 3\" xfId=\"4\" builtinId=\"18\" customBuiltin=\"1\"/><cellStyle name=\"Heading 4\" xfId=\"5\" builtinId=\"19\" customBuiltin=\"1\"/><cellStyle name=\"Input\" xfId=\"9\" builtinId=\"20\" customBuiltin=\"1\"/><cellStyle name=\"Linked Cell\" xfId=\"12\" builtinId=\"24\" customBuiltin=\"1\"/><cellStyle name=\"Neutral\" xfId=\"8\" builtinId=\"28\" customBuiltin=\"1\"/><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\" customBuiltin=\"1\"/><cellStyle name=\"Note\" xfId=\"15\" builtinId=\"10\" customBuiltin=\"1\"/><cellStyle name=\"Output\" xfId=\"10\" builtinId=\"21\" customBuiltin=\"1\"/><cellStyle name=\"Title\" xfId=\"1\" builtinId=\"15\" customBuiltin=\"1\"/><cellStyle name=\"Total\" xfId=\"17\" builtinId=\"25\" customBuiltin=\"1\"/><cellStyle name=\"Warning Text\" xfId=\"14\" builtinId=\"11\" customBuiltin=\"1\"/></cellStyles>"
  "<dxfs count=\"0\"/><tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium9\" defaultPivotStyle=\"PivotStyleLight16\"/>\r\n"
  "</styleSheet>\r\n";

const char* theme_xml =
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
  "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\">\r\n"
  "<a:themeElements><a:clrScheme name=\"Office\"><a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1><a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"1F497D\"/></a:dk2><a:lt2><a:srgbClr val=\"EEECE1\"/></a:lt2><a:accent1><a:srgbClr val=\"4F81BD\"/></a:accent1><a:accent2><a:srgbClr val=\"C0504D\"/></a:accent2><a:accent3><a:srgbClr val=\"9BBB59\"/></a:accent3><a:accent4><a:srgbClr val=\"8064A2\"/></a:accent4><a:accent5><a:srgbClr val=\"4BACC6\"/></a:accent5><a:accent6><a:srgbClr val=\"F79646\"/></a:accent6><a:hlink><a:srgbClr val=\"0000FF\"/></a:hlink><a:folHlink><a:srgbClr val=\"800080\"/></a:folHlink></a:clrScheme><a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"Cambria\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"ＭＳ Ｐゴシック\"/><a:font script=\"Hang\" typeface=\"맑은 고딕\"/><a:font script=\"Hans\" typeface=\"宋体\"/><a:font script=\"Hant\" typeface=\"新細明體\"/><a:font script=\"Arab\" typeface=\"Times New Roman\"/><a:font script=\"Hebr\" typeface=\"Times New Roman\"/><a:font script=\"Thai\" typeface=\"Tahoma\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"MoolBoran\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Times New Roman\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/></a:majorFont><a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"ＭＳ Ｐゴシック\"/><a:font script=\"Hang\" typeface=\"맑은 고딕\"/><a:font script=\"Hans\" typeface=\"宋体\"/><a:font script=\"Hant\" typeface=\"新細明體\"/><a:font script=\"Arab\" typeface=\"Arial\"/><a:font script=\"Hebr\" typeface=\"Arial\"/><a:font script=\"Thai\" typeface=\"Tahoma\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"DaunPenh\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Arial\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/></a:minorFont></a:fontScheme><a:fmtScheme name=\"Office\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"50000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs><a:gs pos=\"35000\"><a:schemeClr val=\"phClr\"><a:tint val=\"37000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:tint val=\"15000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"16200000\" scaled=\"1\"/></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:shade val=\"51000\"/><a:satMod val=\"130000\"/></a:schemeClr></a:gs><a:gs pos=\"80000\"><a:schemeClr val=\"phClr\"><a:shade val=\"93000\"/><a:satMod val=\"130000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"94000\"/><a:satMod val=\"135000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"16200000\" scaled=\"0\"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"><a:shade val=\"95000\"/><a:satMod val=\"105000\"/></a:schemeClr></a:solidFill><a:prstDash val=\"solid\"/></a:ln><a:ln w=\"25400\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln><a:ln w=\"38100\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"20000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"38000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"23000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"35000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"23000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"35000\"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst=\"orthographicFront\"><a:rot lat=\"0\" lon=\"0\" rev=\"0\"/></a:camera><a:lightRig rig=\"threePt\" dir=\"t\"><a:rot lat=\"0\" lon=\"0\" rev=\"1200000\"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w=\"63500\" h=\"25400\"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"40000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs><a:gs pos=\"40000\"><a:schemeClr val=\"phClr\"><a:tint val=\"45000\"/><a:shade val=\"99000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"20000\"/><a:satMod val=\"255000\"/></a:schemeClr></a:gs></a:gsLst><a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"-80000\" r=\"50000\" b=\"180000\"/></a:path></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"80000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"30000\"/><a:satMod val=\"200000\"/></a:schemeClr></a:gs></a:gsLst><a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"50000\" r=\"50000\" b=\"50000\"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/>\r\n"
  "</a:theme>\r\n";
*/

const char* workbook_rels_xml =
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
  "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
  //"<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
  //"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>"
  "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
  //"<Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>"
  "</Relationships>";

const char* workbook_xml =
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
  "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
  "<sheets>"
  "<sheet name=\"SheetName1\" sheetId=\"1\" r:id=\"rId3\"/>"
  "</sheets>"
  "</workbook>";

const char* worksheet_xml_begin =
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
  "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
  //"<dimension ref=\"A1:R2\"/>"
  "<sheetViews>"
  "<sheetView tabSelected=\"1\" workbookViewId=\"0\">"
  "<pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/>"
  "<selection pane=\"bottomLeft\"/>"
  "</sheetView>"
  "</sheetViews>"
  //"<sheetFormatPr defaultRowHeight=\"13.5\"/>"
  //"<cols><col min=\"1\" max=\"1\" width=\"40.625\" customWidth=\"1\"/><col min=\"2\" max=\"6\" width=\"19.625\" customWidth=\"1\"/><col min=\"7\" max=\"7\" width=\"23.625\" customWidth=\"1\"/><col min=\"8\" max=\"9\" width=\"19.625\" customWidth=\"1\"/><col min=\"10\" max=\"10\" width=\"18.625\" customWidth=\"1\"/><col min=\"11\" max=\"11\" width=\"8.625\" customWidth=\"1\"/><col min=\"12\" max=\"12\" width=\"128.625\" customWidth=\"1\"/><col min=\"13\" max=\"14\" width=\"20.625\" customWidth=\"1\"/><col min=\"15\" max=\"17\" width=\"48.625\" customWidth=\"1\"/><col min=\"18\" max=\"18\" width=\"15.625\" customWidth=\"1\"/></cols>"
  "<sheetData>";

/*
"<row><c t=\"inlineStr\"><is><t>Column1</t></is></c><c t=\"inlineStr\"><is><t>Column2</t></is></c></row>\r\n"
"<row><c t=\"inlineStr\"><is><t>Test</t></is></c><c t=\"inlineStr\"><is><t>123</t></is></c></row>\r\n"
//"<row r=\"1\"><c r=\"A1\"><is><t>Test String</t></is></c></row>\r\n"
  //"<row r=\"1\" spans=\"1:18\" s=\"1\" customFormat=\"1\"><c r=\"A1\" s=\"1\" t=\"s\"><v>0</v></c><c r=\"B1\" s=\"1\" t=\"s\"><v>1</v></c><c r=\"C1\" s=\"1\" t=\"s\"><v>2</v></c><c r=\"D1\" s=\"1\" t=\"s\"><v>3</v></c><c r=\"E1\" s=\"1\" t=\"s\"><v>4</v></c><c r=\"F1\" s=\"1\" t=\"s\"><v>5</v></c><c r=\"G1\" s=\"1\" t=\"s\"><v>6</v></c><c r=\"H1\" s=\"1\" t=\"s\"><v>7</v></c><c r=\"I1\" s=\"1\" t=\"s\"><v>8</v></c><c r=\"J1\" s=\"1\" t=\"s\"><v>9</v></c><c r=\"K1\" s=\"1\" t=\"s\"><v>10</v></c><c r=\"L1\" s=\"1\" t=\"s\"><v>11</v></c><c r=\"M1\" s=\"1\" t=\"s\"><v>12</v></c><c r=\"N1\" s=\"1\" t=\"s\"><v>13</v></c><c r=\"O1\" s=\"1\" t=\"s\"><v>14</v></c><c r=\"P1\" s=\"1\" t=\"s\"><v>15</v></c><c r=\"Q1\" s=\"1\" t=\"s\"><v>16</v></c><c r=\"R1\" s=\"1\" t=\"s\"><v>17</v></c></row>\r\n"
  //"<row r=\"2\" spans=\"1:18\"><c r=\"A2\" t=\"s\"><v>20</v></c><c r=\"B2\" t=\"s\"><v>21</v></c><c r=\"C2\" t=\"s\"><v>22</v></c><c r=\"D2\" t=\"s\"><v>23</v></c><c r=\"E2\" t=\"s\"><v>24</v></c><c r=\"F2\" t=\"s\"><v>18</v></c><c r=\"G2\" t=\"s\"><v>25</v></c><c r=\"H2\" t=\"s\"><v>26</v></c><c r=\"I2\" t=\"s\"><v>27</v></c><c r=\"J2\" t=\"s\"><v>28</v></c><c r=\"K2\"><v>1548</v></c><c r=\"L2\" t=\"s\"><v>29</v></c><c r=\"M2\" t=\"s\"><v>30</v></c><c r=\"N2\" t=\"s\"><v>31</v></c><c r=\"O2\" t=\"s\"><v>32</v></c><c r=\"P2\" t=\"s\"><v>18</v></c><c r=\"Q2\" t=\"s\"><v>33</v></c><c r=\"R2\" t=\"s\"><v>19</v></c></row>\r\n"
*/
const char* worksheet_xml_end =
  "</sheetData>"
  //"<pageMargins left=\"0.75\" right=\"0.75\" top=\"1\" bottom=\"1\" header=\"0.5\" footer=\"0.5\"/>"
  "</worksheet>";

////////////////////////////////////////////////////////////////////////

#define PIPEFD_READ   0
#define PIPEFD_WRITE  1

struct xlsxio_write_struct {
  char* filename;
  zip_t* zip;
#ifdef USE_WINTHREADS
  HANDLE thread;
#else
  pthread_t thread;
#endif
  int pipefd[2];
  int rowopen;
};

//thread function used for creating .xlsx file from pipe
#ifdef USE_WINTHREADS
DWORD WINAPI thread_proc (LPVOID arg)
#else
void* thread_proc (void* arg)
#endif
{
  xlsxiowriter handle = (xlsxiowriter)arg;
  //initialize zip file object
  if ((handle->zip = zip_open(handle->filename, ZIP_CREATE, NULL)) == NULL) {
    free(handle);
    free(handle->filename);
#ifdef USE_WINTHREADS
    return 0;
#else
    return NULL;
#endif
  }
  //generate required files
  zip_add_static_content_string(handle->zip, "[Content_Types].xml", content_types_xml);
  zip_add_static_content_string(handle->zip, "docProps/core.xml", docprops_core_xml);
  zip_add_static_content_string(handle->zip, "docProps/app.xml", docprops_app_xml);
  zip_add_static_content_string(handle->zip, "_rels/.rels", rels_xml);
  //zip_add_static_content_string(handle->zip, "xl/sharedStrings.xml", sharedstrings_xml);
  //zip_add_static_content_string(handle->zip, "xl/styles.xml", styles_xml);
  //zip_add_static_content_string(handle->zip, "xl/theme/theme1.xml", theme_xml);
  zip_add_static_content_string(handle->zip, "xl/_rels/workbook.xml.rels", workbook_rels_xml);
  zip_add_static_content_string(handle->zip, "xl/workbook.xml", workbook_xml);
//  zip_add_static_content_string(handle->zip, "xl/worksheets/sheet1.xml", worksheet_xml);

  //add sheet content with pipe data as source
  zip_source_t* zipsrc = zip_source_filep(handle->zip, fdopen(handle->pipefd[PIPEFD_READ], "rb"), 0, -1);
  if (zip_file_add(handle->zip, "xl/worksheets/sheet1.xml", zipsrc, ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8) < 0) {
    zip_source_free(zipsrc);
    fprintf(stdout, "Error adding file");
  }
#ifdef ZIP_RDONLY
  zip_file_set_mtime(handle->zip, zip_get_num_entries(handle->zip, 0) - 1, time(NULL), 0);
#endif

  //close zip file (processes all data, will block until pipe is closed)
  if (zip_close(handle->zip) != 0) {
    int ze, se;
#ifdef ZIP_RDONLY
    zip_error_t* error = zip_get_error(handle->zip);
    ze = zip_error_code_zip(error);
    se = zip_error_code_system(error);
#else
    zip_error_get(handle->zip, &ze, &se);
#endif
    fprintf(stderr, "zip_close failed (%i,%i)\n", ze, se);/////
    fprintf(stderr, "can't close zip archive : %s\n", zip_strerror(handle->zip));
  }
  handle->zip = NULL;
#ifdef USE_WINTHREADS
  return 0;
#else
  return NULL;
#endif
}

DLL_EXPORT_XLSXIO xlsxiowriter xlsxiowrite_open (const char* filename)
{
  xlsxiowriter handle;
  if (!filename)
    return NULL;
  if ((handle = (xlsxiowriter)malloc(sizeof(struct xlsxio_write_struct))) != NULL) {
    //initialize
    handle->filename = strdup(filename);
    handle->zip = NULL;
    handle->pipefd[PIPEFD_READ] = -1;
    handle->pipefd[PIPEFD_WRITE] = -1;
    handle->rowopen = 0;
    //create pipe
    if (pipe(handle->pipefd) != 0) {
      fprintf(stderr, "Error creating pipe\n");/////
    }
    //create and start thread that will receive data via pipe
#ifdef USE_WINTHREADS
    if ((handle->thread = CreateThread(NULL, 0, thread_proc, handle, 0, NULL)) == NULL) {
#else
    if (pthread_create(&handle->thread, NULL, thread_proc, handle) != 0) {
#endif
      fprintf(stderr, "Error creating thread\n");/////
    }
    //write worksheet data
    write(handle->pipefd[PIPEFD_WRITE], worksheet_xml_begin, strlen(worksheet_xml_begin));

  }
  return handle;
}

DLL_EXPORT_XLSXIO int xlsxiowrite_close (xlsxiowriter handle)
{
  if (!handle)
    return -1;

  if (handle->pipefd[PIPEFD_WRITE] == -1)
    return 1;
  //close row if needed
  if (handle->rowopen)
    write(handle->pipefd[PIPEFD_WRITE], "</row>", 6);
  //write worksheet data
  write(handle->pipefd[PIPEFD_WRITE], worksheet_xml_end, strlen(worksheet_xml_end));
  //close pipe
  close(handle->pipefd[PIPEFD_WRITE]);
  //wait for thread to finish
#ifdef USE_WINTHREADS
  WaitForSingleObject(handle->thread, INFINITE);
#else
  pthread_join(handle->thread, NULL);
#endif
  //clean up
  free(handle->filename);
  if (handle->zip)
    zip_close(handle->zip);
  if (handle->pipefd[PIPEFD_READ] != -1)
    close(handle->pipefd[PIPEFD_READ]);
  free(handle);
  return 0;
}

DLL_EXPORT_XLSXIO void xlsxiowrite_add_cell_string (xlsxiowriter handle, const char* value)
{
  if (!handle->rowopen) {
    write(handle->pipefd[PIPEFD_WRITE], "<row>", 5);
    handle->rowopen = 1;
  }
  if (value) {
    write(handle->pipefd[PIPEFD_WRITE], "<c t=\"inlineStr\"><is><t>", 24);
    write(handle->pipefd[PIPEFD_WRITE], value, strlen(value));
    write(handle->pipefd[PIPEFD_WRITE], "</t></is></c>", 13);
  } else {
    write(handle->pipefd[PIPEFD_WRITE], "<c/>", 4);
  }
}

DLL_EXPORT_XLSXIO void xlsxiowrite_next_row (xlsxiowriter handle)
{
  if (handle->rowopen)
    write(handle->pipefd[PIPEFD_WRITE], "</row>", 6);
  handle->rowopen = 0;

}

