//
// This source code resides at www.jaeckel.org/LetsBeRational.7z .
//
// ======================================================================================
// Copyright © 2013-2017 Peter Jäckel.
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
#include "lets_be_rational.h"
#include "version.h"

#if defined(_WIN32) || defined(_WIN64)
# define sleep(length) Sleep(length*1000);
#else
# include <unistd.h> // For sleep().
#endif

#include <string>

static const char * const theNameOfThisXLL = "LetsBeRational";

#ifndef COPYRIGHT
# define COPYRIGHT "© 2013 Peter Jäckel"
#endif
static const char * const copyRight = COPYRIGHT;

#define STRINGIFY(MACRO) STRINGIFY_AUX(MACRO)
#define STRINGIFY_AUX(MACRO) #MACRO

#ifndef REVISION
# define REVISION_STRING " "
#else
# define REVISION_STRING " (revision "STRINGIFY(REVISION)") "
#endif

static XLOper xlAddinManagerInfoText((theNameOfThisXLL+std::string(REVISION_STRING)+copyRight).c_str());

#define XLL_NAME_STRING theNameOfThisXLL

namespace {
   std::string toString(const XLOPER &x){
      if (xltypeStr==(x.xltype&0x0FFF))
         return std::string(x.val.str+1, (size_t)x.val.str[0]);
      return std::string();
   }
}

void registerAllFunctions(){
   // XLOPER used here since XLOper would have its destructor called at time of exit, which would be an attempt to free string memory that was actually allocated by Excel.
   XLOPER xDLL, xlRet;
   Excel4(xlGetName, &xDLL, 0);
   std::string from=toString(xDLL);
   if (from.size()>0)
      from = " from location "+from;
   Excel4(xlcMessage, &xlRet, 2, &XLOper(true), &XLOper(("Registering library '"+(theNameOfThisXLL+("'"+std::string(REVISION_STRING))+copyRight)+from+" ...").c_str()));
   Excel4(xlfRegister, &xlRet, 10 + 1, &xDLL, &XLOper("norm_cdf"),                 &XLOper("BB"),  &XLOper("CumNorm"),             &XLOper("x"),       &XLOper("1"), &XLOper(XLL_NAME_STRING), &XLOper(""), &XLOper(""), &XLOper("The cumulative normal distribution for the given argument x."),            &XLOper("the argument "));
   Excel4(xlfRegister, &xlRet, 10 + 1, &xDLL, &XLOper("inverse_norm_cdf"),         &XLOper("BB"),  &XLOper("InvCumNorm"),          &XLOper("p"),       &XLOper("1"), &XLOper(XLL_NAME_STRING), &XLOper(""), &XLOper(""), &XLOper("The inverse cumulative normal distribution for the given probability p."), &XLOper("the probability "));
   Excel4(xlfRegister, &xlRet, 10 + 2, &xDLL, &XLOper("normalised_black_call"),    &XLOper("BBB"), &XLOper("NormalisedBlackCall"), &XLOper("x,sigma"), &XLOper("1"), &XLOper(XLL_NAME_STRING), &XLOper(""), &XLOper(""), &XLOper("The normalised Black call option value [exp(x/2)·Phi(x/s+s/2)-exp(-x/2)·Phi(x/s-s/2)] with x=ln(F/K) and s=sigma·sqrt(T)."), &XLOper("x=ln(F/K)"), &XLOper("s=sigma·sqrt(T) "));
   Excel4(xlfRegister, &xlRet, 10 + 3, &xDLL, &XLOper("normalised_black"),         &XLOper("BBBB"), &XLOper("NormalisedBlack"), &XLOper("x,sigma,q"), &XLOper("1"), &XLOper(XLL_NAME_STRING), &XLOper(""), &XLOper(""), &XLOper("The normalised Black option value q·[exp(x/2)·Phi(q·(x/s+s/2))-exp(-x/2)·Phi(q·(x/s-s/2))] with x=ln(F/K) and s=sigma·sqrt(T)."), &XLOper("x=ln(F/K)"), &XLOper("s=sigma·sqrt(T)"), &XLOper("q=±1 for calls/puts "));
   Excel4(xlfRegister, &xlRet, 10 + 5, &xDLL, &XLOper("black"), &XLOper("BBBBBB"), &XLOper("Black"), &XLOper("F,K,sigma,T,q"), &XLOper("1"), &XLOper(XLL_NAME_STRING), &XLOper(""), &XLOper(""), &XLOper("The Black option value Black(F,K,sigma,T,q)."), &XLOper("F"), &XLOper("K"), &XLOper("sigma"), &XLOper("T"), &XLOper("q=±1 for calls/puts "));
   Excel4(xlfRegister, &xlRet, 10 + 5, &xDLL, &XLOper("implied_volatility_from_a_transformed_rational_guess"), &XLOper("BBBBBB"), &XLOper("ImpliedVolatility"), &XLOper("price,F,K,T,q"), &XLOper("1"), &XLOper(XLL_NAME_STRING), &XLOper(""), &XLOper(""), &XLOper("The implied volatility s such that the given price equals the normalised Black option value [F·Phi(q·(x/s+s/2))-K·Phi(q·(x/s-s/2))] with x=ln(F/K) and s=sigma·sqrt(T)."), &XLOper("price=Black(F,K,s=sigma·sqrt(T),q)"), &XLOper("F"), &XLOper("K"), &XLOper("T"), &XLOper("q=±1 for calls/puts "));
   Excel4(xlfRegister, &xlRet, 10 + 3, &xDLL, &XLOper("normalised_implied_volatility_from_a_transformed_rational_guess"), &XLOper("BBBB"), &XLOper("NormalisedImpliedVolatility"), &XLOper("beta,x,q"), &XLOper("1"), &XLOper(XLL_NAME_STRING), &XLOper(""), &XLOper(""), &XLOper("The implied volatility s such that the given value beta equals the normalised Black option value q·[exp(x/2)·Phi(q·(x/s+s/2))-exp(-x/2)·Phi(q·(x/s-s/2))] with x=ln(F/K) and s=sigma·sqrt(T)."), &XLOper("beta=Black/sqrt(F*K)"), &XLOper("x=ln(F/K)"), &XLOper("q=±1 for calls/puts "));
#ifdef ENABLE_SWITCHING_THE_OUTPUT_TO_ITERATION_COUNT
   Excel4(xlfRegister, &xlRet, 10 + 1, &xDLL, &XLOper("set_implied_volatility_output_type"), &XLOper("BB"), &XLOper("SetImpliedVolatilityOutputType"), &XLOper("output_type"), &XLOper("1"), &XLOper(XLL_NAME_STRING), &XLOper(""), &XLOper(""), &XLOper("Configures what the implied volatility function(s) returns. Set it to 0 for the implied volatility. Set it to -1 to receive the iteration count."), &XLOper("The new output type (0: volatility, 1: iteration count) "));
#endif
#ifdef ENABLE_CHANGING_THE_HOUSEHOLDER_METHOD_ORDER
   Excel4(xlfRegister, &xlRet, 10 + 1, &xDLL, &XLOper("set_implied_volatility_householder_method_order"), &XLOper("BB"), &XLOper("SetImpliedVolatilityHouseholderMethodOrder"), &XLOper("householder_method_order"), &XLOper("1"), &XLOper(XLL_NAME_STRING), &XLOper(""), &XLOper(""), &XLOper("Configures the order of the Houserholder iteration method used by the implied volatility function(s)."), &XLOper("The new Householder order (-1:query current value, 2: Newton, 3: Halley, 4: Householder(3)) "));
#endif
   Excel4(xlfRegister, &xlRet, 10 + 1, &xDLL, &XLOper("set_implied_volatility_maximum_iterations"), &XLOper("BB"), &XLOper("SetImpliedVolatilityMaximumIterations"), &XLOper("maximum_iterations"), &XLOper("1"), &XLOper(XLL_NAME_STRING), &XLOper(""), &XLOper(""), &XLOper("Configures the maximum number of iterations performed by the implied volatility function(s), or returns the current number of iterations if the input is negative."), &XLOper("The new maximum number of iterations, or, if < 0, a signal to return the current maximum iteration count. "));
   sleep(2);
   Excel4(xlcMessage, &xlRet, 1, &XLOper(false));
}

EXPORT_EXTERN_C LPXLOPER xlAddInManagerInfo(LPXLOPER xAction) {
   XLOPER xIntAction;
   Excel4(xlCoerce, &xIntAction, 2, xAction, &XLOper(xltypeInt));
   if (xIntAction.val.w == 1)
      return &xlAddinManagerInfoText;
   return 0;
}

EXPORT_EXTERN_C int xlAutoOpen() {
   registerAllFunctions();
   return 1;
}

EXPORT_EXTERN_C int xlAutoClose() {
   return 1;
}

EXPORT_EXTERN_C void xlAutoFree(LPXLOPER lpx) {
   XLOper::xLLFree(lpx);
   return;
}

EXTERN_C BOOL APIENTRY DllMain( HANDLE /* hDLL */, DWORD /* dwReason */, LPVOID /* lpReserved */ ) {
   return TRUE;
}
