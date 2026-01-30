# APIMonitor Makefile
# Cross-compile for Windows using MinGW-w64

CC = x86_64-w64-mingw32-gcc
WINDRES = x86_64-w64-mingw32-windres

TARGET = APIMonitor.exe
SOURCES = main.c
RESOURCES = resources.rc

OBJ = main.o resources.o

# Icon sizes for system tray (16=100% DPI, 24=150%, 32=200%, 48=large, 256=hi-res)
ICON_SIZES = 16 24 32 48 256

CFLAGS = -O2 -mwindows -I.
LDFLAGS = -mwindows
LIBS = -lwinhttp -lshell32 -luser32 -lgdi32 -ladvapi32

.PHONY: all clean icons

all: $(TARGET)

$(TARGET): $(OBJ)
	@echo "Linking executable..."
	$(CC) -o $@ $(OBJ) $(LDFLAGS) $(LIBS)
	@rm -f $(OBJ)
	@echo "Build complete: $(TARGET)"

main.o: $(SOURCES) resource.h
	@echo "Compiling $(SOURCES)..."
	$(CC) -c $< -o $@ $(CFLAGS)

resources.o: $(RESOURCES) resource.h empty.ico success.ico fail.ico blank.ico
	@echo "Compiling resources..."
	$(WINDRES) $< -o $@

# Generate multi-sized .ico files from .svg sources using ImageMagick
icons:
	@for svg in *.svg; do \
		if [ -f "$$svg" ]; then \
			ico=$$(basename "$$svg" .svg).ico; \
			echo "Generating $$ico from $$svg..."; \
			convert -background transparent "$$svg" \
				$(foreach size,$(ICON_SIZES),\( -clone 0 -resize $(size)x$(size) \)) \
				-delete 0 -alpha on "$$ico"; \
			echo "Created $$ico with sizes: $(ICON_SIZES)"; \
		fi \
	done

clean:
	rm -f $(OBJ) $(TARGET)
