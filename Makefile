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
INCS =  -Iinclude
CFLAGS = $(INCS) -fexpensive-optimizations -Os
CPPFLAGS = $(INCS) -fexpensive-optimizations -Os
STATIC_CFLAGS = -DBUILD_XLSXIO_STATIC
SHARED_CFLAGS = -DBUILD_XLSXIO_DLL
LIBS =
LDFLAGS =
STRIPFLAG = -s
RM = rm -f
CP = cp -f

XLSXIOREAD_OBJ = lib/xlsxio_read.o
XLSXIOREAD_LDFLAGS = -lzip -lexpat
XLSXIOWRITE_OBJ = lib/xlsxio_write.o
XLSXIOWRITE_LDFLAGS = -lzip
ifneq ($(OS),Windows_NT)
SHARED_CFLAGS += -fPIC
endif
ifeq ($(OS),Windows_NT)
XLSXIOREAD_LDFLAGS += -Wl,--out-implib,$@$(LIBEXT)
XLSXIOWRITE_LDFLAGS += -Wl,--out-implib,$@$(LIBEXT)
else
XLSXIOWRITE_LDFLAGS += -pthread
endif

EXAMPLES_BIN = example_xlsxio_write$(BINEXT) example_xlsxio_read$(BINEXT) example_xlsxio_read_advanced$(BINEXT)

default: all

all: static-lib shared-lib

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
	$(CC) -o $@ -shared -Wl,-soname,$@ $^ $(XLSXIOREAD_LDFLAGS) $(LDFLAGS) $(LIBS) $(STRIPFLAG)

$(LIBPREFIX)xlsxio_write$(LIBEXT): $(XLSXIOWRITE_OBJ:%.o=%.static.o)
	$(AR) cru $@ $^

$(LIBPREFIX)xlsxio_write$(SOEXT): $(XLSXIOWRITE_OBJ:%.o=%.shared.o)
	$(CC) -o $@ -shared -Wl,-soname,$@ $^ $(XLSXIOWRITE_LDFLAGS) $(LDFLAGS) $(LIBS) $(STRIPFLAG)

example_xlsxio_write$(BINEXT): $(LIBPREFIX)xlsxio_write$(LIBEXT) examples/example_xlsxio_write.static.o
	$(CC) -o $@ examples/$(@:%$(BINEXT)=%.static.o) $(LIBPREFIX)xlsxio_write$(LIBEXT) $(XLSXIOWRITE_LDFLAGS) $(STRIPFLAG)

example_xlsxio_read$(BINEXT): $(LIBPREFIX)xlsxio_read$(LIBEXT) examples/example_xlsxio_read.static.o
	$(CC) -o $@ examples/$(@:%$(BINEXT)=%.static.o) $(LIBPREFIX)xlsxio_read$(LIBEXT) $(XLSXIOREAD_LDFLAGS) $(STRIPFLAG)

example_xlsxio_read_advanced$(BINEXT): $(LIBPREFIX)xlsxio_read$(LIBEXT) examples/example_xlsxio_read_advanced.static.o
	$(CC) -o $@ examples/$(@:%$(BINEXT)=%.static.o) $(LIBPREFIX)xlsxio_read$(LIBEXT) $(XLSXIOREAD_LDFLAGS) $(STRIPFLAG)

examples: $(EXAMPLES_BIN)

doc:
	doxygen Doxyfile

.PHONY: clean
clean:
	$(RM) lib/*.o examples/*.o src/*.o *$(LIBEXT) *$(SOEXT) $(EXAMPLES_BIN)

install: all
	mkdir -p $(PREFIX)/include $(PREFIX)/lib
	cp -f *.h $(PREFIX)/include
	cp -f *$(LIBEXT) $(PREFIX)/lib
ifeq ($(OS),Windows_NT)
	mkdir -p $(PREFIX)/bin
	cp -f *$(SOEXT) $(PREFIX)/bin
else
	cp -f *$(SOEXT) $(PREFIX)/lib
endif
