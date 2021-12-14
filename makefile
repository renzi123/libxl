CC = g++
LIBPATH=./lib

CFLAGS = -m64  -I ./include -L $(LIBPATH) -lxl -Wl,-rpath,$(LIBPATH)

all: invoice 

invoice: invoice.cpp
	$(CC) -o invoice invoice.cpp $(CFLAGS)

clean:
	rm -f invoice *.xls *.xlsx

