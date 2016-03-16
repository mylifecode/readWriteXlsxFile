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

#define STR_HELPER(x) #x
#define STR(x) STR_HELPER(x)

DLL_EXPORT_XLSXIO void xlsxiowrite_get_version (int* pmajor, int* pminor, int* pmicro)
{
  if (pmajor)
    *pmajor = XLSXIO_VERSION_MAJOR;
  if (pminor)
    *pminor = XLSXIO_VERSION_MINOR;
  if (pmicro)
    *pmicro = XLSXIO_VERSION_MICRO;
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

#define WITH_XLSX_STYLES 1

#ifdef USE_EXCEL_FOLDERS
#define XML_FOLDER_DOCPROPS             "docProps/"
#define XML_FOLDER_XL                   "xl/"
#define XML_FOLDER_WORKSHEETS           "worksheets/"
#else
#define XML_FOLDER_DOCPROPS             ""
#define XML_FOLDER_XL                   ""
#define XML_FOLDER_WORKSHEETS           ""
#endif
#define XML_FILENAME_CONTENTTYPES       "[Content_Types].xml"
#define XML_FILENAME_RELS               "_rels/.rels"
#define XML_FILENAME_DOCPROPS_CORE      "core.xml"
#define XML_FILENAME_DOCPROPS_APP       "app.xml"
#define XML_FILENAME_XL_WORKBOOK_RELS   "_rels/workbook.xml.rels"
#define XML_FILENAME_XL_WORKBOOK        "workbook.xml"
#define XML_FILENAME_XL_STYLES          "styles.xml"
#define XML_FILENAME_XL_WORKSHEET1      "sheet1.xml"

const char* content_types_xml =
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
  "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
  "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
  "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
  "<Override PartName=\"/" XML_FOLDER_XL XML_FILENAME_XL_WORKBOOK "\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
  "<Override PartName=\"/" XML_FOLDER_DOCPROPS XML_FILENAME_DOCPROPS_CORE "\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>"
  "<Override PartName=\"/" XML_FOLDER_DOCPROPS XML_FILENAME_DOCPROPS_APP "\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>"
#ifdef WITH_XLSX_STYLES
  "<Override PartName=\"/" XML_FOLDER_XL XML_FILENAME_XL_STYLES "\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
#endif
  //"<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>"
  //"<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>"
  "<Override PartName=\"/" XML_FOLDER_XL XML_FOLDER_WORKSHEETS XML_FILENAME_XL_WORKSHEET1 "\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
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
  "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"" XML_FOLDER_DOCPROPS XML_FILENAME_DOCPROPS_CORE "\"/>"
  "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"" XML_FOLDER_DOCPROPS XML_FILENAME_DOCPROPS_APP "\"/>"
  "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"" XML_FOLDER_XL XML_FILENAME_XL_WORKBOOK "\"/>"
  "</Relationships>";

#ifdef WITH_XLSX_STYLES
const char* styles_xml =
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
  "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\r\n"
  "<fonts count=\"2\">\r\n"
  "<font>\r\n"
  "<sz val=\"10\"/>\r\n"
  //"<color theme=\"1\"/>\r\n"
  "<name val=\"Consolas\"/>\r\n"
  "<family val=\"2\"/>\r\n"
  //"<scheme val=\"minor\"/>\r\n"
  "</font>\r\n"
"<font>\r\n"
"<b/><u/>"
"<sz val=\"10\"/>\r\n"
//"<color theme=\"1\"/>\r\n"
"<name val=\"Consolas\"/>\r\n"
"<family val=\"2\"/>\r\n"
//"<scheme val=\"minor\"/>\r\n"
"</font>\r\n"
  "</fonts>\r\n"
  "<fills count=\"1\">\r\n"
  "<fill/>\r\n"
  //"<fill><patternFill patternType=\"none\"/></fill>\r\n"
  "</fills>\r\n"
  "<borders count=\"2\">\r\n"
  "<border>\r\n"
  //"<left/>\r\n"
  //"<right/>\r\n"
  //"<top/>\r\n"
  //"<bottom/>\r\n"
  //"<diagonal/>\r\n"
  "</border>\r\n"
"<border><bottom style=\"thin\"><color indexed=\"64\"/></bottom></border>\r\n"
  "</borders>\r\n"
  //"<cellStyleXfs count=\"1\">\r\n"
  //"<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\r\n"
  //"</cellStyleXfs>\r\n"
  "<cellXfs count=\"6\">\r\n"
  "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>\r\n"
#define STYLE_HEADER 1
  "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment vertical=\"top\"/></xf>\r\n"
#define STYLE_GENERAL 2
  "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment vertical=\"top\"/></xf>\r\n"
#define STYLE_TEXT 3
  "<xf numFmtId=\"49\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyAlignment=\"1\"><alignment vertical=\"top\" wrapText=\"1\"/></xf>\r\n"
#define STYLE_INTEGER 4
  "<xf numFmtId=\"1\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyAlignment=\"1\"><alignment vertical=\"top\"/></xf>\r\n"
#define STYLE_DATETIME 5
  "<xf numFmtId=\"22\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"top\"/></xf>\r\n"
  "</cellXfs>\r\n"
  //"<cellStyles count=\"2\">\r\n"
  //"<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>\r\n"
  //"</cellStyles>\r\n"
  "<dxfs count=\"0\"/>\r\n"
  //"<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium9\" defaultPivotStyle=\"PivotStyleLight16\"/>\r\n"
  "</styleSheet>\r\n";
#endif

/*
const char* theme_xml =
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
  "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\">\r\n"
  "<a:themeElements><a:clrScheme name=\"Office\"><a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1><a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"1F497D\"/></a:dk2><a:lt2><a:srgbClr val=\"EEECE1\"/></a:lt2><a:accent1><a:srgbClr val=\"4F81BD\"/></a:accent1><a:accent2><a:srgbClr val=\"C0504D\"/></a:accent2><a:accent3><a:srgbClr val=\"9BBB59\"/></a:accent3><a:accent4><a:srgbClr val=\"8064A2\"/></a:accent4><a:accent5><a:srgbClr val=\"4BACC6\"/></a:accent5><a:accent6><a:srgbClr val=\"F79646\"/></a:accent6><a:hlink><a:srgbClr val=\"0000FF\"/></a:hlink><a:folHlink><a:srgbClr val=\"800080\"/></a:folHlink></a:clrScheme><a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"Cambria\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"ＭＳ Ｐゴシック\"/><a:font script=\"Hang\" typeface=\"맑은 고딕\"/><a:font script=\"Hans\" typeface=\"宋体\"/><a:font script=\"Hant\" typeface=\"新細明體\"/><a:font script=\"Arab\" typeface=\"Times New Roman\"/><a:font script=\"Hebr\" typeface=\"Times New Roman\"/><a:font script=\"Thai\" typeface=\"Tahoma\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"MoolBoran\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Times New Roman\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/></a:majorFont><a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"ＭＳ Ｐゴシック\"/><a:font script=\"Hang\" typeface=\"맑은 고딕\"/><a:font script=\"Hans\" typeface=\"宋体\"/><a:font script=\"Hant\" typeface=\"新細明體\"/><a:font script=\"Arab\" typeface=\"Arial\"/><a:font script=\"Hebr\" typeface=\"Arial\"/><a:font script=\"Thai\" typeface=\"Tahoma\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"DaunPenh\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Arial\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/></a:minorFont></a:fontScheme><a:fmtScheme name=\"Office\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"50000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs><a:gs pos=\"35000\"><a:schemeClr val=\"phClr\"><a:tint val=\"37000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:tint val=\"15000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"16200000\" scaled=\"1\"/></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:shade val=\"51000\"/><a:satMod val=\"130000\"/></a:schemeClr></a:gs><a:gs pos=\"80000\"><a:schemeClr val=\"phClr\"><a:shade val=\"93000\"/><a:satMod val=\"130000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"94000\"/><a:satMod val=\"135000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"16200000\" scaled=\"0\"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"><a:shade val=\"95000\"/><a:satMod val=\"105000\"/></a:schemeClr></a:solidFill><a:prstDash val=\"solid\"/></a:ln><a:ln w=\"25400\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln><a:ln w=\"38100\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"20000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"38000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"23000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"35000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"23000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"35000\"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst=\"orthographicFront\"><a:rot lat=\"0\" lon=\"0\" rev=\"0\"/></a:camera><a:lightRig rig=\"threePt\" dir=\"t\"><a:rot lat=\"0\" lon=\"0\" rev=\"1200000\"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w=\"63500\" h=\"25400\"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"40000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs><a:gs pos=\"40000\"><a:schemeClr val=\"phClr\"><a:tint val=\"45000\"/><a:shade val=\"99000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"20000\"/><a:satMod val=\"255000\"/></a:schemeClr></a:gs></a:gsLst><a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"-80000\" r=\"50000\" b=\"180000\"/></a:path></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"80000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"30000\"/><a:satMod val=\"200000\"/></a:schemeClr></a:gs></a:gsLst><a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"50000\" r=\"50000\" b=\"50000\"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/>\r\n"
  "</a:theme>\r\n";

const char* sharedstrings_xml =
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
  "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"35\" uniqueCount=\"34\">\r\n"
  "<si><t>Name</t></si><si><t>Expires</t></si><si><t>LastLogon</t></si><si><t>LastSetPassword</t></si><si><t>LastBadPassword</t></si><si><t>LockOut</t></si><si><t>Active</t></si><si><t>Created</t></si><si><t>LastChanged</t></si><si><t>Login</t></si><si><t>Logons</t></si><si><t>ADName</t></si><si><t>FirstName</t></si><si><t>LastName</t></si><si><t>PrimaryEmail</t></si><si><t>HomeDirectory</t></si><si><t>Notes</t></si><si><t>PasswordExpires</t></si><si><t/></si><si><t>Never</t></si><si><t>Firefighter (Realdolmen)</t></si><si><t>2016-02-28 00:00:00</t></si><si><t>2016-02-17 12:00:00</t></si><si><t>2015-12-24 09:41:29</t></si><si><t>2016-02-01 11:50:11</t></si><si><t>Enabled+Expired</t></si><si><t>2013-02-27 10:30:55</t></si><si><t>2016-02-26 08:51:27</t></si><si><t>zxBeweRD008_f</t></si><si><t>CN=Firefighter (Realdolmen),OU=Admin Accounts,OU=Domain Security,DC=ISIS,DC=LOCAL</t></si><si><t>Realdolmen</t></si><si><t>Firefighter</t></si><si><t>Realdolmen.Firefighter@alpro.com</t></si><si><t>consultant RealDolmen - managed services - no often used</t></si>\r\n"
  "</sst>\r\n";
*/

const char* workbook_rels_xml =
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
  "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
  //"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>"
#ifdef WITH_XLSX_STYLES
  "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"" XML_FILENAME_XL_STYLES "\"/>"
#endif
  "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"" XML_FOLDER_WORKSHEETS XML_FILENAME_XL_WORKSHEET1 "\"/>"
  //"<Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>"
  "</Relationships>";

const char* workbook_xml =
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
  "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
  //"<workbookPr/>"
  "<bookViews>"
  "<workbookView/>"
  "</bookViews>"
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

const char* worksheet_xml_end =
  "</sheetData>"
  //"<pageMargins left=\"0.75\" right=\"0.75\" top=\"1\" bottom=\"1\" header=\"0.5\" footer=\"0.5\"/>"
  "</worksheet>";

////////////////////////////////////////////////////////////////////////

//replace part of a string
char* str_replace (char** s, size_t pos, size_t len, char* replacement)
{
  if (!s || !*s)
    return NULL;
  size_t totallen = strlen(*s);
  size_t replacementlen = strlen(replacement);
  if (pos > totallen)
    pos = totallen;
  if (pos + len > totallen)
    len = totallen - pos;
  if (replacementlen > len)
    if ((*s = (char*)realloc(*s, totallen - len + replacementlen + 1)) == NULL)
      return NULL;
  memmove(*s + pos + replacementlen, *s + pos + len, totallen - pos - len + 1);
  memcpy(*s + pos, replacement, replacementlen);
  return *s;
}

//fix string for use as XML data
char* fix_xml_special_chars (char** s)
{
	int pos = 0;
	while (*s && (*s)[pos]) {
		switch ((*s)[pos]) {
			case '&' :
        str_replace(s, pos, 1, "&amp;");
				pos += 5;
				break;
			case '\"' :
				str_replace(s, pos, 1, "&quot;");
				pos += 6;
				break;
			case '<' :
				str_replace(s, pos, 1, "&lt;");
				pos += 4;
				break;
			case '>' :
				str_replace(s, pos, 1, "&gt;");
				pos += 4;
				break;
			case '\r' :
				str_replace(s, pos, 1, "");
				break;
			default:
				pos++;
				break;
		}
	}
	return *s;
}

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
  zip_add_static_content_string(handle->zip, XML_FILENAME_CONTENTTYPES, content_types_xml);
  zip_add_static_content_string(handle->zip, XML_FOLDER_DOCPROPS XML_FILENAME_DOCPROPS_CORE, docprops_core_xml);
  zip_add_static_content_string(handle->zip, XML_FOLDER_DOCPROPS XML_FILENAME_DOCPROPS_APP, docprops_app_xml);
  zip_add_static_content_string(handle->zip, XML_FILENAME_RELS, rels_xml);
#ifdef WITH_XLSX_STYLES
  zip_add_static_content_string(handle->zip, XML_FOLDER_XL XML_FILENAME_XL_STYLES, styles_xml);
#endif
  //zip_add_static_content_string(handle->zip, "xl/theme/theme1.xml", theme_xml);
  zip_add_static_content_string(handle->zip, XML_FOLDER_XL XML_FILENAME_XL_WORKBOOK_RELS, workbook_rels_xml);
  zip_add_static_content_string(handle->zip, XML_FOLDER_XL XML_FILENAME_XL_WORKBOOK, workbook_xml);
  //zip_add_static_content_string(handle->zip, "xl/sharedStrings.xml", sharedstrings_xml);

  //add sheet content with pipe data as source
  zip_source_t* zipsrc = zip_source_filep(handle->zip, fdopen(handle->pipefd[PIPEFD_READ], "rb"), 0, -1);
  if (zip_file_add(handle->zip, XML_FOLDER_XL XML_FOLDER_WORKSHEETS XML_FILENAME_XL_WORKSHEET1, zipsrc, ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8) < 0) {
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

DLL_EXPORT_XLSXIO void xlsxiowrite_add_column (xlsxiowriter handle, const char* value)
{
  if (!handle)
    return;
  if (!handle->rowopen) {
    write(handle->pipefd[PIPEFD_WRITE], "<row s=\"" STR(STYLE_HEADER) "\">", 11);
    handle->rowopen = 1;
  }
  if (value) {
    char* xmlvalue = strdup(value);
    fix_xml_special_chars(&xmlvalue);
#ifdef WITH_XLSX_STYLES
    write(handle->pipefd[PIPEFD_WRITE], "<c t=\"inlineStr\" s=\"" STR(STYLE_HEADER) "\"><is><t>", 30);
#else
    write(handle->pipefd[PIPEFD_WRITE], "<c t=\"inlineStr\"><is><t>", 24);
#endif
    write(handle->pipefd[PIPEFD_WRITE], xmlvalue, strlen(xmlvalue));
    write(handle->pipefd[PIPEFD_WRITE], "</t></is></c>", 13);
    free(xmlvalue);
  } else {
    write(handle->pipefd[PIPEFD_WRITE], "<c/>", 4);
  }
}

DLL_EXPORT_XLSXIO void xlsxiowrite_add_cell_string (xlsxiowriter handle, const char* value)
{
  if (!handle)
    return;
  if (!handle->rowopen) {
    write(handle->pipefd[PIPEFD_WRITE], "<row>", 5);
    handle->rowopen = 1;
  }
  if (value) {
    char* xmlvalue = strdup(value);
    fix_xml_special_chars(&xmlvalue);
#ifdef WITH_XLSX_STYLES
    write(handle->pipefd[PIPEFD_WRITE], "<c t=\"inlineStr\" s=\"" STR(STYLE_TEXT) "\"><is><t>", 30);
#else
    write(handle->pipefd[PIPEFD_WRITE], "<c t=\"inlineStr\"><is><t>", 24);
#endif
    write(handle->pipefd[PIPEFD_WRITE], xmlvalue, strlen(xmlvalue));
    write(handle->pipefd[PIPEFD_WRITE], "</t></is></c>", 13);
    free(xmlvalue);
  } else {
    write(handle->pipefd[PIPEFD_WRITE], "<c/>", 4);
  }
}

DLL_EXPORT_XLSXIO void xlsxiowrite_add_cell_int (xlsxiowriter handle, long value)
{
  if (!handle)
    return;
	char* buf;
	int buflen = snprintf(NULL, 0, "%li", value);
	if (buflen <= 0 || (buf = (char*)malloc(buflen + 1)) == NULL) {
    xlsxiowrite_add_cell_string(handle, NULL);
    return;
	}
	snprintf(buf, buflen + 1, "%li", value);
#ifdef WITH_XLSX_STYLES
  write(handle->pipefd[PIPEFD_WRITE], "<c s=\"" STR(STYLE_INTEGER) "\"><v>", 12);
#else
  write(handle->pipefd[PIPEFD_WRITE], "<c><v>", 6);
#endif
  write(handle->pipefd[PIPEFD_WRITE], buf, strlen(buf));
  write(handle->pipefd[PIPEFD_WRITE], "</v></c>", 8);
  free(buf);
}

DLL_EXPORT_XLSXIO void xlsxiowrite_add_cell_float (xlsxiowriter handle, double value)
{
  if (!handle)
    return;
	char* buf;
	int buflen = snprintf(NULL, 0, "%.32G", value);
	if (buflen <= 0 || (buf = (char*)malloc(buflen + 1)) == NULL) {
    xlsxiowrite_add_cell_string(handle, NULL);
    return;
	}
	snprintf(buf, buflen + 1, "%.32G", value);
#ifdef WITH_XLSX_STYLES
  write(handle->pipefd[PIPEFD_WRITE], "<c s=\"" STR(STYLE_GENERAL) "\"><v>", 12);
#else
  write(handle->pipefd[PIPEFD_WRITE], "<c><v>", 6);
#endif
  write(handle->pipefd[PIPEFD_WRITE], buf, strlen(buf));
  write(handle->pipefd[PIPEFD_WRITE], "</v></c>", 8);
  free(buf);
}

DLL_EXPORT_XLSXIO void xlsxiowrite_add_cell_datetime (xlsxiowriter handle, time_t value)
{
/*
	char buf[20];
	strftime(buf, sizeof(buf), "%Y-%m-%d %H:%M:%S", gmtime(&value));
  write(handle->pipefd[PIPEFD_WRITE], "<c s=\"1\"><v>", 12);
  write(handle->pipefd[PIPEFD_WRITE], buf, strlen(buf));
  write(handle->pipefd[PIPEFD_WRITE], "</v></c>", 8);
*/
  double timestamp = ((double)(value) + .499) / 86400 + 25569; //converstion from Unix to Excel timestamp
	char* buf;
	int buflen = snprintf(NULL, 0, "%.16G", timestamp);
	if (buflen <= 0 || (buf = (char*)malloc(buflen + 1)) == NULL) {
    xlsxiowrite_add_cell_string(handle, NULL);
    return;
	}
	snprintf(buf, buflen + 1, "%.16G", timestamp);
#ifdef WITH_XLSX_STYLES
  write(handle->pipefd[PIPEFD_WRITE], "<c s=\"" STR(STYLE_DATETIME) "\"><v>", 12);
#else
  write(handle->pipefd[PIPEFD_WRITE], "<c><v>", 6);
#endif
  write(handle->pipefd[PIPEFD_WRITE], buf, strlen(buf));
  write(handle->pipefd[PIPEFD_WRITE], "</v></c>", 8);
  free(buf);
}
/*
Windows (And Mac Office 2011+):

    Unix Timestamp = (Excel Timestamp - 25569) * 86400
    Excel Timestamp = (Unix Timestamp / 86400) + 25569

MAC OS X (pre Office 2011):

    Unix Timestamp = (Excel Timestamp - 24107) * 86400
    Excel Timestamp = (Unix Timestamp / 86400) + 24107
*/

DLL_EXPORT_XLSXIO void xlsxiowrite_next_row (xlsxiowriter handle)
{
  if (!handle)
    return;
  if (handle->rowopen)
    write(handle->pipefd[PIPEFD_WRITE], "</row>", 6);
  handle->rowopen = 0;

}

