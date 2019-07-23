#define STRICT
#include <windows.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <dos.h>
#include <direct.h>
#include <time.h>
#include <errno.h>
#include <limits.h>
#include <sys\types.h>
#include <sys\stat.h>
#include "wupbas.h"
#include "wbmem.h"
#include "wbfunc.h"
#include "wbslist.h"
#include "wbfileio.h"
#include "wbfnfile.h"

#define IOBUFFERSIZE    4096

extern NPWBTASK g_npTask;


#define COPYBUFFERSIZE     65500UL

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnFileCopy( int narg, LPVAR lparg, LPVAR lpResult )
{
   HFILE hfileSource, hfileDest;
   OFSTRUCT ofSource, ofDest;
   UINT uDate, uTime;
   UINT uBytes;
   HGLOBAL hglbMem;
   LPSTR lpMem;
   unsigned attrib;
   ERR err;

   if ( narg < 2 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 2 )
      return WBERR_ARGTOOMANY;
      
   assert( lparg[0].type == VT_STRING );
   lpszSource = WBDerefZeroTermHlstr( lparg[0].var.tString );
   
   assert( lparg[1].type == VT_STRING );
   lpszDestFile = WBDerefZeroTermHlstr( lparg[1].var.tString );

   hfileSource = OpenFile( lpszSourceFile, &ofSource, OF_READ );
   if ( hfileSource == HFILE_ERROR )
      return WBERR_DOS00 + ofSource.nErrCode;
      
   _dos_getfileattr( lpszSourceFile, &attrib );
      
   hfileDest = OpenFile( lpszDestFile, &ofDest, OF_CREATE );
   if ( hfileDest == HFILE_ERROR ) {
      _lclose( hfileSource );
      return WBERR_DOS00 + ofDest.nErrCode;
   }

   _dos_getfileattr( lpszSourceFile, &attrib );
   
   if ( (hglbMem = GlobalAlloc( GMEM_MOVEABLE, COPYBUFFERSIZE )) == NULL ) {
      _lclose( hfileSource );
      _lclose( hfileDest );
      return WBERR_OUTOFMEMORY;
   }

   if ( (lpMem = (LPSTR)GlobalLock( hglbMem )) == NULL ) {
      _lclose( hfileSource );
      _lclose( hfileDest );
      GlobalFree( hglbMem );
      return WBERR_OUTOFMEMORY;
   }

   // read and write blocks until complete file copied

   err = 0;
   while ( uBytes = _lread( hfileSource, lpMem, COPYBUFFERSIZE )) {

      if ( _lwrite( hfileDest, lpMem, uBytes ) != uBytes ) {
         err = WBERR_DISKFULL;
         break;
      }
   }

   dwEnd = GetMilliSecs();

   // set the date & time of the destination file equal to that of the source

   _dos_setftime( hfileDest, uDate, uTime );
   _dos_setfileattr( lpszDestFile, attrib );

   _lclose( hfileSource );
   _lclose( hfileDest );

   GlobalUnlock( hglbMem );
   GlobalFree( hglbMem );
   
   lpResult->var.tLong = WB_TRUE;
   lpResult->type = VT_I4;
   return 0;
}

