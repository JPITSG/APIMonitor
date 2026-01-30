#!/bin/bash

x86_64-w64-mingw32-windres resources.rc -o resources.o
x86_64-w64-mingw32-gcc -O2 main.c resources.o -o APIMonitor.exe -lwinhttp -lshell32 -luser32 -lgdi32 -ladvapi32 -mwindows
