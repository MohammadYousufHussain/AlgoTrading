//
// This source code resides at www.jaeckel.org/LetsBeRational.7z .
//
// ======================================================================================
// Copyright © 2013-2014 Peter Jäckel.
// 
// Permission to use, copy, modify, and distribute this software is freely granted,
// provided that this notice is preserved.
//
// WARRANTY DISCLAIMER
// The Software is provided "as is" without warranty of any kind, either express or implied,
// including without limitation any implied warranties of condition, uninterrupted use,
// merchantability, fitness for a particular purpose, or non-infringement.
// ======================================================================================
//
#include "xlcall.h"
#include <stdlib.h>
#include <limits.h>

#ifdef _MSC_VER
# define strncpy(dest,source,length) strncpy_s(dest,length,source,length)
#endif

namespace {
   template <class T>
   inline void deleteArray( T *&p ){
      if (p) {
         delete[] p;
         p = 0;
      }
   }

   template <class T>
   inline void deleteOne( T *&p ){
      if (p) {
         delete p;
         p = 0;
      }
   }
}

///////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////
// Constructors and Destructor:
XLOper::XLOper(const char *s) {
   val.str = 0;
   xltype = xltypeStr;
   unsigned long l = strlen(s)+1;
   val.str = new char[l+1];
   strncpy(val.str+1,s,l);
   if (l>UCHAR_MAX)
      l=UCHAR_MAX; // Truncate to maximum Pascal string length.
   val.str[0] = (char)(l-1);
}

XLOper::~XLOper(){
   xltype  |= xlbitDLLFree;
   xLLFree(*this);
}

XLOper::XLOper(int i) {
   xltype = xltypeInt;
   val.w = i;
}

XLOper::XLOper(double x) {
   xltype = xltypeNum;
   val.num = x;
}

XLOper::XLOper(bool b) {
   xltype = xltypeBool;
   val.boolean = b;
}

XLOper::XLOper(XLType xlOperType) {
   xltype = xlOperType;
   val.err = xlerrNum;    //      Default value.
}

///////////////////////////////////////////////////////////////////////////////
// Safe recursive delete of an XLOper
void XLOper::xLLFree(XLOPER &p) {
   if (p.xltype & xlbitDLLFree){
      switch (p.xltype ^ xlbitDLLFree) {
    case xltypeStr:
       deleteArray(p.val.str);
       break;
    case xltypeMulti:
       {
          unsigned short i, j, k = p.val.array.columns;
          for (i=0; i<p.val.array.rows; ++i) {
             for (j=0; j<p.val.array.columns; ++j) {
                xLLFree(p.val.array.lparray[i*k + j]);
             }
          }
          deleteArray(p.val.array.lparray);
       }
       break;
    default:
       break;
      }
      // The following statement is a safety precaution.
      p.xltype = xltypeNil;
   }
   return;
}

void XLOper::xLLFree(LPXLOPER p) {
   if (p) {
      xLLFree(*p);
      deleteOne(p);
   }
   return;
}

XLOper XLOper::XLError(int error_type){
   XLOper err(XLTypeErr);
   err.val.err = error_type;
   return err;
}

XLOper XLOper::xlErrValue = XLOper::XLError(xlerrValue);
