CMAKE_MINIMUM_REQUIRED(VERSION 2.6)
PROJECT(xlsxio)

# dependancies
SET(CMAKE_MODULE_PATH "${CMAKE_CURRENT_SOURCE_DIR}/CMake" ${CMAKE_MODULE_PATH})
FIND_PACKAGE(libzip REQUIRED)
FIND_PACKAGE(expat REQUIRED)

#SET(CMAKE_CXX_FLAGS "-g -Wall")
#ADD_SUBDIRECTORY(w01-cpp)

#IF(NOT WIN32)
#ENDIF

INCLUDE_DIRECTORIES(include)
#ADD_LIBRARY(xlsxio_read STATIC lib/xlsxio_read.c ${libzip_LIBRARIES} ${expat_LIBRARIES})
#ADD_LIBRARY(xlsxio_write STATIC lib/xlsxio_write.c ${libzip_LIBRARIES})


ADD_LIBRARY(xlsxio_read SHARED lib/xlsxio_read.c ${libzip_LIBRARIES} ${expat_LIBRARIES})
ADD_LIBRARY(xlsxio_read_static STATIC lib/xlsxio_read.c ${libzip_LIBRARIES} ${expat_LIBRARIES})
ADD_LIBRARY(xlsxio_write SHARED lib/xlsxio_write.c ${libzip_LIBRARIES})
ADD_LIBRARY(xlsxio_write_static STATIC lib/xlsxio_write.c ${libzip_LIBRARIES})
SET_TARGET_PROPERTIES (xlsxio_read PROPERTIES DEFINE_SYMBOL "BUILD_XLSXIO_DLL")
SET_TARGET_PROPERTIES (xlsxio_read_static PROPERTIES DEFINE_SYMBOL "BUILD_XLSXIO_STATIC")
SET_TARGET_PROPERTIES (xlsxio_write PROPERTIES DEFINE_SYMBOL "BUILD_XLSXIO_DLL")
SET_TARGET_PROPERTIES (xlsxio_write_static PROPERTIES DEFINE_SYMBOL "BUILD_XLSXIO_STATIC")

TARGET_LINK_LIBRARIES(xlsxio_read xlsxio_write)