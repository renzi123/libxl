CC = g++
LIBPATH=./lib

CFLAGS = -m64  -I ./include -L $(LIBPATH) -lxl -Wl,-rpath,$(LIBPATH)

all: demo 

demo: demo.cpp
	$(CC) -o demo demo.cpp $(CFLAGS)

clean:
	rm -f demo *.xls *.xlsx

