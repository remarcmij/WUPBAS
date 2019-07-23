#define STRICT
#include <windows.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <limits.h>
#include "wupbas.h"
#include "wbfunc.h"
#include "wbtype.h"
#include "wbmem.h"
#include "wbfnstr.h"

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnAsc( int narg, LPVAR lparg, LPVAR lpResult )
{  
   LPBYTE lpsz;
   USHORT usLen;

   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   lpsz = (LPBYTE)WBDerefHlstrLen( lparg[0].var.tString, &usLen );
   if ( usLen == 0 )
      return WBERR_ARGRANGE;
      
   lpResult->var.tLong = (long)*lpsz;
   lpResult->type = VT_I4;
   
   return 0;
}
     
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnChr( int narg, LPVAR lparg, LPVAR lpResult )
{  
   LPBYTE lpsz;

   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   if ( lparg[0].var.tLong < 0 || lparg[0].var.tLong > 255 )
      return WBERR_ARGRANGE;

   lpResult->var.tString = WBCreateTempHlstr( NULL, 1 );
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
   lpsz = (LPBYTE)WBDerefHlstr( lpResult->var.tString );
   *lpsz = (BYTE)lparg[0].var.tLong;

   lpResult->type = VT_STRING;
   return 0;
}
     
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnHex( int narg, LPVAR lparg, LPVAR lpResult )
{
   char szBuffer[32];
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;
      
   _ltoa( lparg[0].var.tLong, szBuffer, 16 );
   _strupr( szBuffer );
   
   lpResult->var.tString = WBCreateTempHlstr( szBuffer, (USHORT)strlen( szBuffer ) );
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
   lpResult->type = VT_STRING;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnIniCaps( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpsz;
   USHORT usLen;

   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;
   
   lpResult->var.tString = WBCreateTempHlstrFromTemp( lparg[0].var.tString );
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;

   lpsz = WBDerefHlstrLen( lpResult->var.tString, &usLen );
   
   while ( usLen ) {

      if ( isalpha((int)*lpsz) ) {   
         *lpsz = (char)toupper((int)*lpsz);
         lpsz++;
         usLen--;
         while ( usLen > 0 && isalpha((int)*lpsz) ) {
            if ( isupper((int)*lpsz) )
               *lpsz = (char)tolower((int)*lpsz);
            lpsz++;
            usLen--;
         }
      }
      
      while ( usLen > 0 && isalpha( (int)*lpsz ) == 0 ) {
         lpsz++;
         usLen--;
      }
   }
   
   lpResult->type = VT_STRING;
   return 0;
}
     
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnInStr( int narg, LPVAR lparg, LPVAR lpResult )
{
   HLSTR hlstr1, hlstr2;
   LPSTR lpsz1, lpsz2;
   USHORT usLen1, usLen2;
   USHORT usStart, usPos;
   BOOL fCompareText;
   ERR err;
   
   switch ( narg ) {
      
      case 2:
         if ( (err = WBVariantToString( &lparg[0] )) != 0 )
            return err;
         hlstr1 = lparg[0].var.tString;
         
         if ( (err = WBVariantToString( &lparg[1] )) != 0 )
            return err;
         hlstr2 = lparg[1].var.tString;
         
         usStart = 0;
         fCompareText = g_npTask->fCompareText;
         break;
         
      case 3:
      case 4:
         if ( (err = WBMakeLongVariant( &lparg[0] )) != 0 )
            return err;
         if ( lparg[0].var.tLong < 1 || lparg[0].var.tLong > MAXSTRING-1 )
            return WBERR_ARGRANGE;
         usStart = (USHORT)lparg[0].var.tLong - 1;
            
         if ( (err = WBVariantToString( &lparg[1] )) != 0 )
            return err;
         hlstr1 = lparg[1].var.tString;
            
         if ( (err = WBVariantToString( &lparg[2] )) != 0 )
            return err;
         hlstr2 = lparg[2].var.tString;
         
         if ( narg == 4 ) {
            if ( (err = WBMakeLongVariant( &lparg[3] )) != 0 )
               return err;
            if ( lparg[3].var.tLong != 0L && lparg[3].var.tLong != 1L )
               return WBERR_ARGRANGE;
            fCompareText = (BOOL)lparg[3].var.tLong;
         }
         else {
            fCompareText = g_npTask->fCompareText;
         }
         break;
         
      default:
         if ( narg < 2 )
            return WBERR_ARGTOOFEW;
         return WBERR_ARGTOOMANY;
   }

   lpsz1 = WBDerefHlstrLen( hlstr1, &usLen1 );
   lpsz2 = WBDerefHlstrLen( hlstr2, &usLen2 );
      
   if ( usLen2 == 0 ) {
      // by definition
      lpResult->var.tLong = (long)(usStart + 1);
   }
   else if ( usLen1 == 0 || (long)usLen2 > (long)usLen1 - (long)usStart ) {
      // no need to search
      lpResult->var.tLong = 0L;                          
   }
   else {
      usPos = WBStrPos( usStart, lpsz1, usLen1, lpsz2, usLen2, fCompareText );
      if ( usPos == (USHORT)-1 )
         lpResult->var.tLong = 0;
      else
         lpResult->var.tLong = (long)usPos+1;
   }

   lpResult->type = VT_I4;
   return 0;
}     

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnIsNumeric( int narg, LPVAR lparg, LPVAR lpResult )
{
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   if ( lparg[0].type == VT_I4 )
      lpResult->var.tLong = WB_TRUE;
   else if ( WBVariantStringToOther( &lparg[0], VT_I4 ) == 0 )
      lpResult->var.tLong = WB_TRUE;
   else
      lpResult->var.tLong = WB_FALSE;
   
   lpResult->type = VT_I4;
   return 0;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*  Token( source, delim )         Current form                             */
/*  TokenCount( source [,delim] )  Supersed form retained for compatiblity  */
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnTokenCount( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszSource;
   LPSTR lpszDelim;
   LPSTR lpch, lpszStart;
   USHORT usLen;
   int nCount;
   char chDelim;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 2 )
      return WBERR_ARGTOOMANY;

   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   
   if ( narg == 2 ) {
      lpszDelim = WBDerefHlstrLen( lparg[1].var.tString, &usLen );
      if ( usLen == 0 )
         return WBERR_ARGRANGE;
      chDelim = *lpszDelim;
   }
   else 
      chDelim = ',';
      
   lpszSource = WBDerefHlstrLen( lparg[0].var.tString, &usLen );
   lpch = lpszStart = lpszSource;
   nCount = 0;

   while ( usLen > 0 ) {
   
      nCount++;

      // skip white space         
      while ( usLen > 0 && isspace( *lpch ) ) {
         lpch++;
         usLen--;
      }
      
      lpszStart = lpch;
      
      // don't break up quoted strings
      if ( usLen > 0 && *lpch == '"' ) {
      
         lpch++;
         usLen--;
         
         while ( usLen > 0 && *lpch != '"' ) {
            lpch++; 
            usLen--;
         }
            
         if ( usLen > 0 && *lpch == '"' ) {
            lpch++; 
            usLen--;               
         }
      }
            
            
      while ( usLen > 0 && *lpch != chDelim ) {
         lpch++; 
         usLen--;
      }
         
      if ( usLen > 0 ) {
         lpch++; 
         usLen--;
      }
   }

   lpResult->var.tLong = nCount;
   lpResult->type = VT_I4;
   return 0;
}      

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*  Token( source, index, [,delim] )                                        */
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnToken( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszSource;
   LPSTR lpszDelim;
   LPSTR lpch, lpszStart;
   USHORT usLen;
   int i, nIndex;
   char chDelim;
   char chQuote;
   
   if ( narg < 2 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 3 )
      return WBERR_ARGTOOMANY;

   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;
         
   if ( lparg[1].var.tLong < 1 || lparg[1].var.tLong > INT_MAX )
      return WBERR_ARGRANGE;
   nIndex = (int)lparg[1].var.tLong;
      
   if ( narg == 3 ) {
      assert( lparg[2].type == VT_STRING );
      lpszDelim = WBDerefHlstrLen( lparg[2].var.tString, &usLen );
      if ( usLen == 0 )
         return WBERR_ARGRANGE;
      chDelim = *lpszDelim;
   }
   else 
      chDelim = ',';
      
   lpszSource = WBLockHlstrLen( lparg[0].var.tString, &usLen );
   lpch = lpszStart = lpszSource;

   for ( i = 0; i < nIndex; i++ ) {

      // skip white space         
      while ( usLen > 0 && isspace( *lpch ) ) {
         lpch++;
         usLen--;
      }
      
      lpszStart = lpch;
      
      // don't break up quoted strings
      if ( usLen > 0 && *lpch == '"' ) {
      
         lpch++;
         usLen--;
         
         while ( usLen > 0 && *lpch != '"' ) {
            lpch++; 
            usLen--;
         }
            
         if ( usLen > 0 && *lpch == '"' ) {
            lpch++; 
            usLen--;               
         }
      }
            
            
      while ( usLen > 0 && *lpch != chDelim ) {
         lpch++; 
         usLen--;
      }
         
      if ( usLen > 0 && i < nIndex - 1 ) {
         lpch++; 
         usLen--;
      }
         
   }

   if ( lpszSource != NULL )
      usLen = (USHORT)(lpch - lpszStart);
   else
      assert( usLen == 0 );
      
   // Remove leading spaces
   while ( usLen > 0 && isspace( *lpszStart ) ) {
      usLen--;
      lpszStart++;
   }
   // Remove trailing spaces
   while ( usLen > 0 && isspace( *(lpszStart+usLen-1) ) )
      usLen--;
      
   // If both first and last characters of result are quotes,
   // remove them both
   if ( usLen >= 2 && (*lpszStart == '"' || *lpszStart == '\'') ) {
      chQuote = *lpszStart;
      if ( *lpszStart == chQuote && *(lpszStart + usLen - 1) == chQuote ) {
         lpszStart++;
         usLen -= 2;
      }
   }
   
   lpResult->var.tString = WBCreateTempHlstr( lpszStart, usLen );
   lpResult->type = VT_STRING;
   WBUnlockHlstr( lparg[0].var.tString );
   
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
   return 0;
}      

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*  TokenExtract( index, source, [,delim] )                                 */
/*                                                                          */
/*  Superseded by Token() but retained for compatibility                    */
/*  Calls WBFnToken with first and second argument swapped.                 */
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnTokenExtract( int narg, LPVAR lparg, LPVAR lpResult )
{
   VARIANT vTemp;
   
   vTemp = lparg[0];
   lparg[0] = lparg[1];
   lparg[1] = vTemp;
   
   return WBFnToken( narg, lparg, lpResult );
}
      
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnLcase( int narg, LPVAR lparg, LPVAR lpResult )
{  
   LPSTR lpsz;
   USHORT usLen;
   USHORT k;

   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;
   
   lpResult->var.tString = WBCreateTempHlstrFromTemp( lparg[0].var.tString );
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
   lpsz = WBDerefHlstrLen( lpResult->var.tString, &usLen );
   
   for ( k = 0; k < usLen; k++, lpsz++ ) {
      if ( isupper( *lpsz ) )
         *lpsz = (char)tolower( (int)*lpsz );
   }
   
   lpResult->type = VT_STRING;
   return 0;
}
     
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnLeft( int narg, LPVAR lparg, LPVAR lpResult )
{
   USHORT usLen;
   ERR err;

   if ( narg < 2 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 2 )
      return WBERR_ARGTOOMANY;

   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;

   if ( lparg[1].var.tLong < 0 || lparg[1].var.tLong > MAXSTRING-1 )
      return WBERR_ARGRANGE;      

   lpResult->var.tString = WBCreateTempHlstrFromTemp( lparg[0].var.tString );
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
   
   usLen = WBGetHlstrLen( lpResult->var.tString );
   usLen = min( (USHORT)lparg[1].var.tLong, usLen );
   if ( (err = WBResizeHlstr( lpResult->var.tString, usLen )) != 0 )
      return err;
      
   lpResult->type = VT_STRING;
   return 0;
}
     
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnLen( int narg, LPVAR lparg, LPVAR lpResult )
{
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   switch ( lparg[0].type ) {
   
      case VT_STRING:
         lpResult->var.tLong = (long)WBGetHlstrLen( lparg[0].var.tString );
         lpResult->type = VT_I4;
         break;
         
      case VT_USERTYPE:
         {
            LPWBTYPEMEM lpTypeMem;
            
            lpTypeMem = (LPWBTYPEMEM)lparg[0].var.tlpVoid;
            lpResult->var.tLong = (long)lpTypeMem->npUserType->uSize;
            lpResult->type = VT_I4;
         }
         break;
         
      default:
         return WBERR_TYPEMISMATCH;
   }
   
   return 0; 
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnLTrim( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpsz;
   int nLeadingBlanks;
   USHORT usLen;

   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   lpsz = WBDerefHlstrLen( lparg[0].var.tString, &usLen );
   
   // Trim leading blanks
   nLeadingBlanks = 0;
   while ( usLen > 0 && isspace( *(lpsz + nLeadingBlanks) ) ) {
      usLen--;
      nLeadingBlanks++;
   }
   
   lpsz = WBLockHlstr( lparg[0].var.tString );
   lpResult->var.tString = WBCreateTempHlstr( lpsz + nLeadingBlanks, usLen );
   WBUnlockHlstr( lparg[0].var.tString );
   
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
   
   lpResult->type = VT_STRING;
   return 0;
}
     
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnMid( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpsz;
   USHORT usStart, usLen, usSize;
   
   if ( narg < 2 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 3 )
      return WBERR_ARGTOOMANY;

   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   usSize = WBGetHlstrLen( lparg[0].var.tString );
   
   if ( lparg[1].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   if ( lparg[1].var.tLong < 1 || lparg[1].var.tLong >= MAXSTRING-1 )
      return WBERR_ARGRANGE;
   usStart = (USHORT)lparg[1].var.tLong;
   
   if ( narg == 3 ) {
      if ( lparg[2].var.tLong < 1 || lparg[2].var.tLong >= MAXSTRING-1 )
         return WBERR_ARGRANGE;
         
      usLen = min( (USHORT)lparg[2].var.tLong, usSize - usStart + 1 );
   }
   else {
      usLen = usSize - usStart + 1;
   }
   
   lpsz = WBLockHlstr( lparg[0].var.tString );
   if ( usStart > usSize ) 
      lpResult->var.tString = WBCreateTempHlstr( NULL, 0 );
   else
      lpResult->var.tString = WBCreateTempHlstr( lpsz + usStart - 1, usLen );
   WBUnlockHlstr( lparg[0].var.tString );
   
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
   lpResult->type = VT_STRING;
   return 0;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/* Overlay( target, position, source, start, len                            */
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnOverlay( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszSource, lpszTarget, lpszResult;
   USHORT cbSource, cbTarget;
   USHORT uPos, uStart, cbLen;

   if ( narg < 3 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 5 )
      return WBERR_ARGTOOMANY;

   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGREQUIRED;

   lpszTarget = WBDerefHlstrLen( lparg[0].var.tString, &cbTarget );

   if ( lparg[1].type == VT_EMPTY )
      return WBERR_ARGREQUIRED;

   if ( lparg[1].var.tLong < 1 || lparg[1].var.tLong > (long)cbTarget )
      return WBERR_ARGRANGE;

   uPos = (USHORT)lparg[1].var.tLong - 1;

   if ( lparg[2].type == VT_EMPTY )
      return WBERR_ARGREQUIRED;

   lpszSource = WBDerefHlstrLen( lparg[2].var.tString, &cbSource );

   if ( narg < 4 || lparg[3].type == VT_EMPTY ) {
      uStart = 0;
   }
   else {
      if ( lparg[3].var.tLong < 1 || lparg[3].var.tLong > (long)cbSource )
         return WBERR_ARGRANGE;
      uStart = (USHORT)lparg[3].var.tLong - 1;
   }

   if ( narg < 5 || lparg[4].type == VT_EMPTY ) {
      cbLen = cbSource - uStart;
   }
   else {
      if ( lparg[4].var.tLong < 0 || 
           (long)uStart + lparg[4].var.tLong > (long)cbSource )
         return WBERR_ARGRANGE;
      cbLen = (USHORT)lparg[4].var.tLong;

      if ( uPos + cbLen > cbTarget )
         return WBERR_ARGRANGE;
   }

   lpResult->var.tString = WBCreateTempHlstr( lpszTarget, cbTarget );
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
   lpResult->type = VT_STRING;

   lpszResult = WBDerefHlstr( lpResult->var.tString );

   if ( cbLen != 0 )
      memcpy( lpszResult + uPos, lpszSource + uStart, cbLen );

   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnRight( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpsz;
   USHORT usSize, usLen;

   if ( narg < 2 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 2 )
      return WBERR_ARGTOOMANY;

   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;

   if ( lparg[1].var.tLong < 0 || lparg[1].var.tLong > MAXSTRING-1 )
      return WBERR_ARGRANGE;
   usSize = (USHORT)lparg[1].var.tLong;      
      
   lpsz = WBLockHlstrLen( lparg[0].var.tString, &usLen );
   usSize = min( (USHORT)lparg[1].var.tLong, usLen );
   lpResult->var.tString = WBCreateTempHlstr( lpsz + usLen - usSize, usSize );
   WBUnlockHlstr( lparg[0].var.tString );
   
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
   lpResult->type = VT_STRING;
   return 0;
}
     
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnRTrim( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpsz;
   USHORT usLen; 
   ERR err;

   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   lpsz = WBDerefHlstrLen( lparg[0].var.tString, &usLen );
   
   // Trim trailing blanks
   while ( usLen > 0 && isspace(*(lpsz + usLen - 1)) )
      usLen--;

   lpResult->var.tString = WBCreateTempHlstrFromTemp( lparg[0].var.tString );
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
   if ( (err = WBResizeHlstr( lpResult->var.tString, usLen )) != 0 )
      return err;
   
   lpResult->type = VT_STRING;
   return 0;
}
     
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnSpace( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpsz;
   HLSTR hlstr;
   USHORT usLen;

   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   if ( lparg[0].var.tLong < 0 || lparg[0].var.tLong > MAXSTRING )
      return WBERR_ARGRANGE;
   usLen = (USHORT)lparg[0].var.tLong;
   
   hlstr = WBCreateTempHlstr( NULL, usLen );
   if ( hlstr == NULL )
      return WBERR_STRINGSPACE;
      
   if ( usLen > 0 ) {
      lpsz = WBDerefHlstr( hlstr );
      assert( lpsz != NULL );
      memset( lpsz, ' ', (size_t)usLen );
   }
   
   lpResult->var.tString = hlstr;
   lpResult->type = VT_STRING;
   return 0;
}
     
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnStr( int narg, LPVAR lparg, LPVAR lpResult )
{
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   switch ( lparg[0].type ) {
   
      case VT_STRING:
         lpResult->var.tString = WBCreateTempHlstrFromTemp( lparg[0].var.tString );
         if ( lpResult->var.tString == NULL )
            return WBERR_STRINGSPACE;
            
         lpResult->type = VT_STRING;
         return 0;
         
      default:
         *lpResult = lparg[0];
         return WBVariantToString( lpResult );
   }
}
     
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnString( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPBYTE lpsz;
   USHORT usLen;
   long lTemp;

   if ( narg < 2 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 2 )
      return WBERR_ARGTOOMANY;

   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   if ( lparg[0].var.tLong < 0 || lparg[0].var.tLong > MAXSTRING )
      return WBERR_ARGRANGE;
   usLen = (USHORT)lparg[0].var.tLong;
   
   switch ( lparg[1].type ) {
   
      case VT_I4:
         lTemp = lparg[1].var.tLong;
         break;
         
      case VT_STRING:
         lpsz = (LPBYTE)WBDerefHlstr( lparg[1].var.tString );
         if ( lpsz == NULL )
            return WBERR_ARGRANGE;
         lTemp = (long)(*lpsz);
         break;
         
      default:
         return WBERR_TYPEMISMATCH;
   }

   lpResult->var.tString = WBCreateTempHlstr( NULL, usLen );
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
   if ( usLen > 0 ) {
      lpsz = (LPBYTE)WBDerefHlstr( lpResult->var.tString );
      assert( lpsz != NULL );
      memset( lpsz, (int)(lTemp %= 256L), (size_t)usLen );
   }
   
   lpResult->type = VT_STRING;
   return 0;
}
     
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnTrim( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpsz;
   int nLeadingBlanks;
   USHORT usLen;

   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;
      
   lpsz = WBDerefHlstrLen( lparg[0].var.tString, &usLen );
   
   // Trim leading blanks
   nLeadingBlanks = 0;
   while ( usLen > 0 && isspace( *(lpsz + nLeadingBlanks) ) ) {
      usLen--;
      nLeadingBlanks++;
   }
   
   // Trim trailing blanks
   while ( usLen > 0 && isspace(*(lpsz + nLeadingBlanks + usLen - 1)) )
      usLen--;
      
   lpsz = WBLockHlstr( lparg[0].var.tString );
   lpResult->var.tString = WBCreateTempHlstr( lpsz + nLeadingBlanks, usLen );
   WBUnlockHlstr( lparg[0].var.tString );
   
   if (  lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
   lpResult->type = VT_STRING;
   return 0;
}
     
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnUcase( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpsz;
   USHORT usLen;
   USHORT k;

   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   lpResult->var.tString = WBCreateTempHlstrFromTemp( lparg[0].var.tString );
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
   lpsz = WBDerefHlstrLen( lpResult->var.tString, &usLen );
   
   for ( k = 0; k < usLen; k++, lpsz++ ) {
      if ( islower( *lpsz ) )
         *lpsz = (char)toupper( (int)*lpsz );
   }
   
   lpResult->type = VT_STRING;
   return 0;
}
     
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnWord( int narg, LPVAR lparg, LPVAR lpResult )
{
   USHORT i, usIndex, usLen;
   LPSTR lpsz, lpszStart;

   if ( narg < 2 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 2 )
      return WBERR_ARGTOOMANY;

   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;
            
   if ( lparg[1].var.tLong <= 0L || lparg[1].var.tLong >= (long)USHRT_MAX )
      return WBERR_ARGRANGE;
      
   usIndex = (USHORT)lparg[1].var.tLong;

   lpsz = WBLockHlstrLen( lparg[0].var.tString, &usLen );

   // Skip the first usIndex-1 words   
   for ( i = 0; i < usIndex-1; i++ ) {
       
      // skip white space
      while ( usLen > 0 && isspace( *lpsz ) ) {
         usLen--;
         lpsz++;
      }
      
      while( usLen > 0 && !isspace( *lpsz ) ) {
         usLen--;
         lpsz++;
      }
      
      if ( usLen == 0 ) 
         break;
   }
   
   // skip white space 
   while ( usLen > 0 && isspace( *lpsz ) ) {
      usLen--;
      lpsz++;
   }
   
   if ( usLen == 0 ) {
      WBUnlockHlstr( lparg[0].var.tString );
      lpResult->var.tString = WBCreateTempHlstr( NULL, 0 );
      if ( lpResult->var.tString == NULL )
         return WBERR_STRINGSPACE;
      lpResult->type = VT_STRING;
      return 0;
   }
   

   lpszStart = lpsz;
   
   while( usLen > 0 && !isspace( *lpsz ) ) {
      usLen--;
      lpsz++;
   }
   
   usLen = (USHORT)(lpsz - lpszStart);
   
   lpResult->var.tString = WBCreateTempHlstr( lpszStart, usLen );
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
   lpResult->type = VT_STRING;
   WBUnlockHlstr( lparg[0].var.tString );
   return 0;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnWords( int narg, LPVAR lparg, LPVAR lpResult )
{
   USHORT usCount, usLen;
   LPSTR lpsz;

   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;
      
   lpsz = WBDerefHlstrLen( lparg[0].var.tString, &usLen );
   usCount = 0;
   
   for ( usCount = 0; ; usCount++ ) {
   
      // skip white space
      while ( usLen > 0 && isspace( *lpsz ) ) {
         usLen--;
         lpsz++;
      }
      
      if ( usLen == 0 ) 
         break;
         
      while( usLen > 0 && !isspace( *lpsz ) ) {
         usLen--;
         lpsz++;
      }
      
   }
   
   lpResult->var.tLong = (long)usCount;
   lpResult->type = VT_I4;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnLeftPart( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszText, lpszDelim, lpszPos, lpszNew;
   char chDelim;
   USHORT usLen, usNew;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 2 )
      return WBERR_ARGTOOMANY;

   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;
      
   if ( narg == 2 ) {
      lpszDelim = WBDerefZeroTermHlstr( lparg[1].var.tString );
      if ( *lpszDelim == '\0' )
         return WBERR_ARGRANGE;
      chDelim = *lpszDelim;
   }
   else {
      chDelim = (char)'=';
   }

   lpszText = WBLockHlstrLen( lparg[0].var.tString, &usLen );
   
   // skip leading white space
   while ( usLen > 0 && isspace( *lpszText ) ) {
      usLen--;
      lpszText++;
   }
    
   lpszNew = NULL;
   usNew = 0;
      
   if ( usLen > 0 ) {
      lpszPos = memchr( lpszText, chDelim, (size_t)usLen );
      if ( lpszPos != NULL ) {
         lpszPos--;
         while ( lpszPos >= lpszText && isspace( *lpszPos ) )
            *lpszPos--;
         lpszNew = lpszText;
         usNew = (USHORT)(lpszPos - lpszText + 1);
      }
   }
   
   lpResult->var.tString = WBCreateTempHlstr( lpszNew, usNew );
   WBUnlockHlstr( lparg[0].var.tString );
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
   lpResult->type = VT_STRING;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnRightPart( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszText, lpszDelim, lpszPos, lpszNew;
   char chDelim;
   USHORT usLen, usNew;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 2 )
      return WBERR_ARGTOOMANY;
      
   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;

   if ( narg == 2 ) {
      lpszDelim = WBDerefZeroTermHlstr( lparg[1].var.tString );
      if ( *lpszDelim == '\0' )
         return WBERR_ARGRANGE;
      chDelim = *lpszDelim;
   }
   else {
      chDelim = (char)'=';
   }

   lpszText = WBLockHlstrLen( lparg[0].var.tString, &usLen );
   
   // skip leading white space
   while ( usLen > 0 && isspace( *lpszText ) ) {
      usLen--;
      lpszText++;
   }
    
   lpszNew = NULL;
   usNew = 0;
      
   if ( usLen > 0 ) {
      lpszPos = memchr( lpszText, chDelim, (size_t)usLen );
      if ( lpszPos != NULL ) {
         usNew = usLen - (USHORT)(lpszPos-lpszText+1);
         lpszNew = lpszPos+1;

         // Trim leading spaces         
         while ( usNew > 0 && isspace( *lpszNew ) ) {
            usNew--;
            lpszNew++;
         }

         // Trim trailing spaces         
         if ( usNew > 0 ) {
            lpszPos = lpszNew + usNew - 1;
            while ( lpszPos >= lpszNew && isspace( *lpszPos ) ) {
               lpszPos--;
               usNew--;
            }
         }

         if ( usNew == 0 )
            lpszNew = NULL;
      }
   }
   
   lpResult->var.tString = WBCreateTempHlstr( lpszNew, usNew );
   WBUnlockHlstr( lparg[0].var.tString );
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
   lpResult->type = VT_STRING;
   return 0;
}
