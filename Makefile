ifeq ($(OS),)
OS = $(shell uname -s)
endif
PREFIX = /usr/local
CC   = gcc
CPP  = g++
AR   = ar
LIBPREFIX = lib
LIBEXT = .a
ifeq ($(OS),Windows_NT)
BINEXT = .exe
SOEXT = .dll
else ifeq ($(OS),Darwin)
BINEXT =
SOEXT = .dylib
else
BINEXT =
SOEXT = .so
endif
INCS = -Iinclude -Ilib
CFLAGS = $(INCS) -Os
CPPFLAGS = $(INCS) -Os
STATIC_CFLAGS = -DBUILD_XLSXIO_STATIC
SHARED_CFLAGS = -DBUILD_XLSXIO_DLL
LIBS =
LDFLAGS =
ifeq ($(OS),Darwin)
CFLAGS += -I/opt/local/include -I/opt/local/lib/libzip/include
LDFLAGS += -L/opt/local/lib
#CFLAGS += -arch i386 -arch x86_64
#LDFLAGS += -arch i386 -arch x86_64
STRIPFLAG =
else
STRIPFLAG = -s
endif
MKDIR = mkdir -p
RM = rm -f
RMDIR = rm -rf
CP = cp -f
CPDIR = cp -rf
DOXYGEN := $(shell which doxygen)

XLSXIOREAD_OBJ = lib/xlsxio_read.o
XLSXIOREAD_LDFLAGS = -lzip -lexpat
XLSXIOREAD_SHARED_LDFLAGS =
XLSXIOWRITE_OBJ = lib/xlsxio_write.o
XLSXIOWRITE_LDFLAGS = -lzip
XLSXIOWRITE_SHARED_LDFLAGS =
ifneq ($(OS),Windows_NT)
SHARED_CFLAGS += -fPIC
endif
ifeq ($(OS),Windows_NT)
XLSXIOREAD_SHARED_LDFLAGS += -Wl,--out-implib,$@$(LIBEXT)
XLSXIOWRITE_SHARED_LDFLAGS += -Wl,--out-implib,$@$(LIBEXT)
else ifeq ($(OS),Darwin)
else
XLSXIOWRITE_LDFLAGS += -pthread
endif
ifeq ($(OS),Darwin)
OS_LINK_FLAGS = -dynamiclib -o $@
else
OS_LINK_FLAGS = -shared -Wl,-soname,$@ $(STRIPFLAG)
endif

TOOLS_BIN = xlsxio_xlsx2csv$(BINEXT) xlsxio_csv2xlsx$(BINEXT)
EXAMPLES_BIN = example_xlsxio_write_getversion$(BINEXT) example_xlsxio_write$(BINEXT) example_xlsxio_read$(BINEXT) example_xlsxio_read_advanced$(BINEXT)

COMMON_PACKAGE_FILES = README.md LICENSE.txt Changelog.txt
SOURCE_PACKAGE_FILES = $(COMMON_PACKAGE_FILES) Makefile CMakeLists.txt CMake/ doc/Doxyfile include/*.h lib/*.c src/*.c examples/*.c build/*.cbp

default: all

all: static-lib shared-lib tools

%.o: %.c
	$(CC) -c -o $@ $< $(CFLAGS) 

%.static.o: %.c
	$(CC) -c -o $@ $< $(STATIC_CFLAGS) $(CFLAGS) 

%.shared.o: %.c
	$(CC) -c -o $@ $< $(SHARED_CFLAGS) $(CFLAGS)

static-lib: $(LIBPREFIX)xlsxio_read$(LIBEXT) $(LIBPREFIX)xlsxio_write$(LIBEXT)

shared-lib: $(LIBPREFIX)xlsxio_read$(SOEXT) $(LIBPREFIX)xlsxio_write$(SOEXT)

$(LIBPREFIX)xlsxio_read$(LIBEXT): $(XLSXIOREAD_OBJ:%.o=%.static.o)
	$(AR) cru $@ $^

$(LIBPREFIX)xlsxio_read$(SOEXT): $(XLSXIOREAD_OBJ:%.o=%.shared.o)
	$(CC) -o $@ $(OS_LINK_FLAGS) $^ $(XLSXIOREAD_SHARED_LDFLAGS) $(XLSXIOREAD_LDFLAGS) $(LDFLAGS) $(LIBS)

$(LIBPREFIX)xlsxio_write$(LIBEXT): $(XLSXIOWRITE_OBJ:%.o=%.static.o)
	$(AR) cru $@ $^

$(LIBPREFIX)xlsxio_write$(SOEXT): $(XLSXIOWRITE_OBJ:%.o=%.shared.o)
	$(CC) -o $@ $(OS_LINK_FLAGS) $^ $(XLSXIOWRITE_SHARED_LDFLAGS) $(XLSXIOWRITE_LDFLAGS) $(LDFLAGS) $(LIBS)

examples: $(EXAMPLES_BIN)

example_xlsxio_write_getversion$(BINEXT): examples/example_xlsxio_write_getversion.static.o $(LIBPREFIX)xlsxio_write$(LIBEXT)
	$(CC) -o $@ examples/$(@:%$(BINEXT)=%.static.o) $(LIBPREFIX)xlsxio_write$(LIBEXT) $(XLSXIOWRITE_LDFLAGS) $(LDFLAGS)

example_xlsxio_write$(BINEXT): examples/example_xlsxio_write.static.o $(LIBPREFIX)xlsxio_write$(LIBEXT)
	$(CC) -o $@ examples/$(@:%$(BINEXT)=%.static.o) $(LIBPREFIX)xlsxio_write$(LIBEXT) $(XLSXIOWRITE_LDFLAGS) $(LDFLAGS)

example_xlsxio_read$(BINEXT): examples/example_xlsxio_read.static.o $(LIBPREFIX)xlsxio_read$(LIBEXT)
	$(CC) -o $@ examples/$(@:%$(BINEXT)=%.static.o) $(LIBPREFIX)xlsxio_read$(LIBEXT) $(XLSXIOREAD_LDFLAGS) $(LDFLAGS)

example_xlsxio_read_advanced$(BINEXT): examples/example_xlsxio_read_advanced.static.o $(LIBPREFIX)xlsxio_read$(LIBEXT)
	$(CC) -o $@ examples/$(@:%$(BINEXT)=%.static.o) $(LIBPREFIX)xlsxio_read$(LIBEXT) $(XLSXIOREAD_LDFLAGS) $(LDFLAGS)

tools: $(TOOLS_BIN)

xlsxio_xlsx2csv$(BINEXT): src/xlsxio_xlsx2csv.static.o $(LIBPREFIX)xlsxio_read$(LIBEXT)
	$(CC) -o $@ $< $(LIBPREFIX)xlsxio_read$(LIBEXT) $(XLSXIOREAD_LDFLAGS) $(LDFLAGS)

xlsxio_csv2xlsx$(BINEXT): src/xlsxio_csv2xlsx.static.o $(LIBPREFIX)xlsxio_write$(LIBEXT)
	$(CC) -o $@ $< $(LIBPREFIX)xlsxio_write$(LIBEXT) $(XLSXIOWRITE_LDFLAGS) $(LDFLAGS)

.PHONY: doc
doc:
ifdef DOXYGEN
	$(DOXYGEN) doc/Doxyfile
endif

install: all doc
	$(MKDIR) $(PREFIX)/include $(PREFIX)/lib $(PREFIX)/bin
	$(CP) include/*.h $(PREFIX)/include/
	$(CP) *$(LIBEXT) $(PREFIX)/lib/
ifeq ($(OS),Windows_NT)
	$(CP) *$(SOEXT) $(PREFIX)/bin/
else
	$(CP) *$(SOEXT) $(PREFIX)/lib/
endif
	$(CP) $(TOOLS_BIN) $(PREFIX)/bin/
ifdef DOXYGEN
	$(CPDIR) doc/man $(PREFIX)/
endif

.PHONY: version
version:
	sed -ne "s/^#define\s*XLSXIO_VERSION_[A-Z]*\s*\([0-9]*\)\s*$$/\1./p" include/xlsxio_version.h | tr -d "\n" | sed -e "s/\.$$//" > version

.PHONY: package
package: version
	tar cfJ xlsxio-$(shell cat version).tar.xz --transform="s?^?xlsxio-$(shell cat version)/?" $(SOURCE_PACKAGE_FILES)

.PHONY: package
binarypackage: version
	$(MAKE) PREFIX=binarypackage_temp install
	tar cfJ "xlsxio-$(shell cat version)-$(OS).tar.xz" --transform="s?^binarypackage_temp/??" $(COMMON_PACKAGE_FILES) binarypackage_temp/*
	rm -rf binarypackage_temp

.PHONY: clean
clean:
	$(RM) lib/*.o examples/*.o src/*.o *$(LIBEXT) *$(SOEXT) $(TOOLS_BIN) $(EXAMPLES_BIN) version xlsxio-*.tar.xz doc/doxygen_sqlite3.db
	$(RMDIR) doc/html doc/man

