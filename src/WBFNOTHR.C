#define STRICT
#include <windows.h>
#include <string.h>
#include "wupbas.h"
#include "wbfunc.h"

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnVarType( int narg, LPVAR lparg, LPVAR lpResult )
{
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;
      
   lpResult->var.tLong = lparg[0].type;
   lpResult->type = VT_I4;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnInterpret( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpsz;
   USHORT usLen;
   ERR err;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   if ( WBPushContext() == NULL )
      return WBERR_OUTOFMEMORY;

   lpsz = WBDerefHlstrLen( lparg[0].var.tString, &usLen );
   usLen = min( usLen, sizeof( g_npContext->szProgText ) - 1 );
   memcpy( g_npContext->szProgText, lpsz, (size_t)usLen );
   g_npContext->szProgText[usLen] = '\0';
   
   g_npContext->npvtLocal = g_npContext->npNext->npvtLocal;
   g_npContext->npchNextPos = g_npContext->szProgText;
      
   NextToken();
   err = WBExpression( lpResult );

   WBPopContext();
   
   return err;
}
   
   
      
