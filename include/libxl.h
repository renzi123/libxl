///////////////////////////////////////////////////////////////////////////////
//                                                                           //
//                    LibXL C++ headers version 4.0.0                        //
//                                                                           //
//       Copyright (c) 2008 - 2021 Dmytro Skrypnyk and XLware s.r.o.         //
//                                                                           //
//   THIS FILE AND THE SOFTWARE CONTAINED HEREIN IS PROVIDED 'AS IS' AND     //
//                COMES WITH NO WARRANTIES OF ANY KIND.                      //
//                                                                           //
//           Please define LIBXL_STATIC variable for static linking.         //
//                                                                           //
///////////////////////////////////////////////////////////////////////////////

#ifndef LIBXL_CPP_H
#define LIBXL_CPP_H

#define LIBXL_VERSION 0x04000100

#include "libxl_IBookT.h"
#include "libxl_ISheetT.h"
#include "libxl_IFormatT.h"
#include "libxl_IFontT.h"
#include "libxl_IAutoFilterT.h"
#include "libxl_IFilterColumnT.h"
#include "libxl_IRichStringT.h"
#include "libxl_IFormControlT.h"

#define KEY "libxl"
#if (defined WIN32)|| (defined _MSC_VER)
#define KEY_SN "windows-28232b0208c4ee0369ba6e68abv6v5i3"
#else
#define KEY_SN   "linux-68636b6268646e63696a6e686bp6p5i3"
#endif

namespace libxl {

    #ifdef _UNICODE
        typedef IBookT<wchar_t> Book;
        typedef ISheetT<wchar_t> Sheet;
        typedef IFormatT<wchar_t> Format;
        typedef IFontT<wchar_t> Font;
        typedef IAutoFilterT<wchar_t> AutoFilter;
        typedef IFilterColumnT<wchar_t> FilterColumn;
        typedef IRichStringT<wchar_t> RichString;
        typedef IFormControlT<wchar_t> FormControl;
        #define _xlCreateBook xlCreateBookW
        #define _xlCreateXMLBook xlCreateXMLBookW
    #else
        typedef IBookT<char> Book;
        typedef ISheetT<char> Sheet;
        typedef IFormatT<char> Format;
        typedef IFontT<char> Font;
        typedef IAutoFilterT<char> AutoFilter;
        typedef IFilterColumnT<char> FilterColumn;
        typedef IRichStringT<char> RichString;
        typedef IFormControlT<char> FormControl;
        #define _xlCreateBook xlCreateBookA
        #define _xlCreateXMLBook xlCreateXMLBookA
    #endif


static inline Book* xlCreateBook() {Book* bk=_xlCreateBook();bk->setKey(KEY, KEY_SN);return bk;}
static inline Book* xlCreateXMLBook() {Book* bk=_xlCreateXMLBook();bk->setKey(KEY, KEY_SN);return bk;}
}

#endif
