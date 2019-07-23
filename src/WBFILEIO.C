#define STRICT
#include <windows.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h> 
#include <io.h>
#include <limits.h>
#include "wupbas.h"
#include "wbfunc.h"
#include "wbfileio.h"

extern NPWBTASK g_npTask;

#define DEFAULT_BUFSIZE       128

static ERR WBFlushFileBuffer( OPENFILESTRUCT FAR* lpof );
   
typedef struct tagFILEOBJECT {
// Property data
   HLSTR hlstrFileName;
   BOOL  fEof;
   BOOL  fOpen;
   LONG  lLineCount;
// Private data
   OPENFILESTRUCT FAR* lpof;
} FILEDATA;
typedef FILEDATA FAR* LPFILE;
   
enum property_indices {
   IPROP_SLIST_FILENAME,
   IPROP_FILE_EOF,
   IPROP_FILE_OPEN,
   IPROP_FILE_LINECOUNT
};

#define OFFSETIN(struc, field) ((USHORT)LOWORD(&(((struc *)0)->field)))

static PROPINFO Property_FileName = {
   "FileName",
   DT_HLSTR | PF_fGetData | PF_fNoRuntimeW,
   OFFSETIN(FILEDATA, hlstrFileName)
};

static PROPINFO Property_Eof = {
   "Eof",
   DT_BOOL | PF_fGetData | PF_fNoRuntimeW,
   OFFSETIN(FILEDATA, fEof)
};

static PROPINFO Property_Open = {
   "IsOpen",
   DT_BOOL | PF_fGetData | PF_fNoRuntimeW,
   OFFSETIN(FILEDATA, fOpen)
};

static PROPINFO Property_LineCount = {
   "LineCount",
   DT_LONG | PF_fGetData | PF_fNoRuntimeW,
   OFFSETIN(FILEDATA, lLineCount)
};

// Count is an alias for LineCount
static PROPINFO Property_Count = {
   "Count",
   DT_LONG | PF_fGetData | PF_fNoRuntimeW,
   OFFSETIN(FILEDATA, lLineCount)
};

static PPROPINFO File_Properties[] = {
   &Property_FileName,
   &Property_Eof,
   &Property_Open,
   &Property_LineCount,
   &Property_Count,
   NULL
};

ERR WBFile_MethodOpen( HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult );
ERR WBFile_MethodClose( HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult );
ERR WBFile_MethodRead( HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult );
ERR WBFile_MethodWrite( HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult );
ERR WBFile_MethodSeekGet( HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult );
ERR WBFile_MethodSeekSet( HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult );

static METHODINFO Method_Open    = { "OPEN",  "SII", WBFile_MethodOpen  };
static METHODINFO Method_Close   = { "CLOSE", "",    WBFile_MethodClose };
static METHODINFO Method_Read    = { "READ",  "I",   WBFile_MethodRead  };
static METHODINFO Method_Write   = { "WRITE", "S",   WBFile_MethodWrite };
static METHODINFO Method_SeekGet = { "SEEKGET", "",  WBFile_MethodSeekGet };
static METHODINFO Method_SeekSet = { "SEEKSET", "I", WBFile_MethodSeekSet };

// Keep this list ordered from most frequently to least frequently
// used method for best performance
static PMETHODINFO File_Methods[] = {
   &Method_Read,
   &Method_Write,
   &Method_Open,
   &Method_Close,
   &Method_SeekGet,
   &Method_SeekSet,
   NULL
};

ERR WBFileCtlProc ( HCTL hctl, USHORT msg, USHORT wp, LONG lp );

static MODEL modelFILE = {
   (PCTLPROC)WBFileCtlProc,         // Control procedure
   sizeof( FILEDATA ),              // Size of FILEDATE structure
   File_Properties,                 // Property information table
   File_Methods,                    // Method information table
   "FILE",                          // Object class name
   VT_FILE,                         // Object variant type
   FALSE,
};

#define lpFileDEREF(hctl)  ((LPFILE)WBDerefControl(hctl))

ERR WBFile_OnCreate( HCTL hctl );
ERR WBFile_OnDestroy( HCTL hctl );

/*--------------------------------------------------------------------------*/
/* Note: called by LibMain                                                  */
/*--------------------------------------------------------------------------*/
BOOL WBRegisterModel_FileObject( void )
{
   return WBRegisterModel( NULL, &modelFILE );
}

//---------------------------------------------------------------------------
// File Control Procedure
//---------------------------------------------------------------------------
ERR WBFileCtlProc ( HCTL hctl, USHORT msg, USHORT wp, LONG lp )
{
   switch ( msg ) {
   
      case WM_CREATE:
         WBFile_OnCreate( hctl );
         break;
      
      case WM_DESTROY:
         WBFile_OnDestroy( hctl );
         break;
   }
               
   return WBDefControlProc(hctl, msg, wp, lp);
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBFile_OnCreate( HCTL hctl )
{
   LPFILE lpfile;
             
   lpfile = lpFileDEREF( hctl );

   // set property defaults
   lpfile->hlstrFileName = WBCreateHlstr( NULL, 0 );
   lpfile->fEof          = TRUE;
   lpfile->fOpen         = FALSE;
   lpfile->lLineCount    = 0;
   lpfile->lpof          = NULL;
   
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBFile_OnDestroy( HCTL hctl )
{
   LPFILE lpfile;
   OPENFILESTRUCT FAR* lpof;
   
   lpfile = lpFileDEREF( hctl );
   lpof = lpfile->lpof;
   
   if ( lpof != NULL && lpof->fOpen ) {
      WBCloseFile( lpof );
      WBLocalFree( (HLOCAL)LOWORD(lpof) );
   }
   
   WBDestroyHlstr( lpfile->hlstrFileName );
   
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBFile_MethodOpen( HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult )
{
   LPFILE lpfile;
   LPSTR lpsz = NULL;
   USHORT usLen;
   char szFileName[_MAX_PATH];
   OPENFILESTRUCT FAR* lpof;
   USHORT fuMode;
   USHORT usBufSize;
   ERR err;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 3 )
      return WBERR_ARGTOOMANY;
      
   lpfile = lpFileDEREF( hctl );

   // 1st arg: file name
   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;
      
   lpsz = WBDerefHlstrLen( lparg[0].var.tString, &usLen );
   assert( lpsz != NULL );
   usLen = min( usLen, sizeof( szFileName ) - 1 );
   memcpy( szFileName, lpsz, usLen );
   szFileName[usLen] = '\0';
   
   // 2nd arg: open mode
   if ( narg >= 2 ) {
      if ( lparg[1].type == VT_EMPTY ) {
         fuMode = OM_INPUT;
      }
      else {
         if ( (err = WBMakeLongVariant( &lparg[1] )) != 0 )
            return err;
         fuMode = (USHORT)lparg[1].var.tLong;
      }
   }
   else {
      fuMode = OM_INPUT;
   }
   
   if ( fuMode != OM_INPUT  && fuMode != OM_OUTPUT && 
        fuMode != OM_APPEND && fuMode != OM_BINARY ) {
      return WBERR_ARGRANGE;
   }        
   
   // 3rd arg: buffer size
   if ( narg == 3 ) {
      if ( lparg[2].var.tLong < 1 || lparg[2].var.tLong > MAXSTRING )
         return WBERR_ARGRANGE;
      usBufSize = (USHORT)lparg[2].var.tLong;
   }
   else {
      if ( fuMode == OM_BINARY )
         usBufSize = 1;
      else
         usBufSize = DEFAULT_BUFSIZE;
   }

   lpof = (OPENFILESTRUCT FAR*)WBLocalAlloc( LPTR, sizeof( OPENFILESTRUCT ));
   if ( lpof == NULL )
      return WBERR_OUTOFMEMORY;

   strcpy( lpof->szFileName, szFileName );      
   lpof->enumOpenMode = fuMode;
   lpof->nRecLen = (int)usBufSize;
   
   err = WBOpenFile( lpof );

   if ( err == 0 ) {
      usLen = (USHORT)strlen( lpof->szFileName);
      err = WBResizeHlstr( lpfile->hlstrFileName, usLen );
      if ( err == 0 ) {
         lpsz = WBDerefHlstr( lpfile->hlstrFileName );
         memcpy( lpsz, lpof->szFileName, usLen );
      }
   }
   
   lpfile->lpof = lpof;
   lpfile->fOpen = TRUE;
   lpfile->fEof = lpof->fEof;
      
   lpResult->var.tLong = WB_TRUE;
   lpResult->type = VT_I4;
   return err;
}
  
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBFile_MethodClose( HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult )
{
   LPFILE lpfile;
   OPENFILESTRUCT FAR* lpof;

   UNREFERENCED_PARAM( lparg );   
   
   if ( narg != 0 )
      return WBERR_EXPECT_NOARG;
      
   lpfile = lpFileDEREF( hctl );
   lpof = lpfile->lpof;
   
   if ( lpof == NULL || !lpof->fOpen )
      return WBERR_FILENOTOPEN;

   WBCloseFile( lpof );
   WBLocalFree( (HLOCAL)LOWORD(lpof) );
   lpfile->lpof = NULL;
   
   lpfile->fOpen      = FALSE;
   lpfile->fEof       = TRUE;
   lpfile->lLineCount = 0;
         
   lpResult->var.tLong = WB_TRUE;
   lpResult->type = VT_I4;
   return 0;
}
         
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBFile_MethodRead( HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult )
{
   LPFILE lpfile;
   OPENFILESTRUCT FAR* lpof;
   HLSTR hlstr;
   ERR err;

   if ( narg > 1 )
      return WBERR_ARGTOOMANY;
      
   lpfile = lpFileDEREF( hctl );
   lpof = lpfile->lpof;
   
   if ( lpof->enumOpenMode != OM_INPUT && lpof->enumOpenMode != OM_BINARY )
      return WBERR_INCOMPATIBLEFILEMODE;
      
   if ( narg == 1 ) {
   
      if ( lpof->enumOpenMode != OM_BINARY )
         return WBERR_INCOMPATIBLEFILEMODE;

      if ( lparg[0].var.tLong <= 0 || lparg[0].var.tLong > (long)INT_MAX )
         return WBERR_ARGRANGE;
         
      lpof->nRecLen = (int)lparg[0].var.tLong;
   }
      
   if ( (err = WBReadString( lpof, &hlstr )) != 0 )
      return err;
      
   lpfile->fEof = lpof->fEof;
   lpfile->lLineCount++;      
      
   lpResult->var.tString = WBCreateTempHlstrFromTemp( hlstr );
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;

   WBDestroyHlstr( hlstr );
   lpResult->type = VT_STRING;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBFile_MethodWrite( HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult )
{
   LPFILE lpfile;
   OPENFILESTRUCT FAR* lpof;
   ERR err;

   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if (narg > 1 )
      return WBERR_ARGTOOMANY;
      
   lpfile = lpFileDEREF( hctl );
   lpof = lpfile->lpof;
   
   if ( lpof->enumOpenMode == OM_INPUT )
      return WBERR_INCOMPATIBLEFILEMODE;

   if ( (err = WBWriteString( lpof, lparg[0].var.tString )) != 0 )
      return err;

   lpfile->lLineCount++;      
   
   lpResult->var.tLong = WB_TRUE;      
   lpResult->type = VT_I4;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBFile_MethodSeekGet( HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult )
{
   LPFILE lpfile;
   OPENFILESTRUCT FAR* lpof;
   
   UNREFERENCED_PARAM( lparg );   
   
   if ( narg > 0 )
      return WBERR_EXPECT_NOARG;
   
   lpfile = lpFileDEREF( hctl );
   lpof = lpfile->lpof;
   
   if ( lpof->enumOpenMode != OM_BINARY )
      return WBERR_INCOMPATIBLEFILEMODE;
      
   lpResult->var.tLong = _llseek( lpof->hFile, 0L, 1 );
   if ( lpResult->var.tLong == HFILE_ERROR )
      return WBERR_SEEKERROR;
      
   lpResult->type = VT_I4;
   return 0;
}
      
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBFile_MethodSeekSet( HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult )
{
   LPFILE lpfile;
   OPENFILESTRUCT FAR* lpof;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if (narg > 1 )
      return WBERR_ARGTOOMANY;
   
   lpfile = lpFileDEREF( hctl );
   lpof = lpfile->lpof;
   
   if ( lpof->enumOpenMode != OM_BINARY )
      return WBERR_INCOMPATIBLEFILEMODE;
      
   lpResult->var.tLong = _llseek( lpof->hFile, lparg[0].var.tLong, 0 );
   if ( lpResult->var.tLong == HFILE_ERROR )
      return WBERR_SEEKERROR;
      
   lpResult->type = VT_I4;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBOpenFile( OPENFILESTRUCT FAR* lpof )
{
   UINT uPrevErrorMode;
   OFSTRUCT of;

   if ( lpof->enumOpenMode != OM_BINARY ) {   
   
      lpof->hglbBuffer = GlobalAlloc( GMEM_MOVEABLE, (UINT)lpof->nRecLen );
      if ( lpof->hglbBuffer == NULL )
         WBRuntimeError( WBERR_OUTOFMEMORY );
         
      lpof->lpBuffer = (LPBYTE)GlobalLock( lpof->hglbBuffer );
      if ( lpof->lpBuffer == NULL)
      assert( lpof->lpBuffer != NULL );
   }      
   
   uPrevErrorMode = SetErrorMode( SEM_NOOPENFILEERRORBOX );   
   
   switch ( lpof->enumOpenMode ) {
   
      case OM_INPUT:
         lpof->hFile = OpenFile( lpof->szFileName, &of, OF_READ );
         break;
         
      case OM_OUTPUT:
         lpof->hFile = OpenFile( lpof->szFileName, &of, OF_CREATE );
         break;
         
      case OM_APPEND:
         lpof->hFile = OpenFile( lpof->szFileName, &of, OF_WRITE );
         if ( lpof->hFile != HFILE_ERROR )
            _lseek( lpof->hFile, 0, SEEK_END );
         else
            lpof->hFile = OpenFile( lpof->szFileName, &of, OF_CREATE );
         break;
         
      case OM_BINARY:
         lpof->hFile = OpenFile( lpof->szFileName, &of, OF_READWRITE );
         if ( lpof->hFile == HFILE_ERROR )
            lpof->hFile = OpenFile( lpof->szFileName, &of, OF_WRITE );
         if ( lpof->hFile == HFILE_ERROR )
            lpof->hFile = OpenFile( lpof->szFileName, &of, OF_READ );
         break;
         
      default:
         assert( FALSE );
         return WBERR_INTERNAL_ERROR;
   }
   
   SetErrorMode( uPrevErrorMode );
         
   if ( lpof->hFile == HFILE_ERROR ) {
      if ( lpof->enumOpenMode != OM_BINARY ) {
         GlobalUnlock( lpof->hglbBuffer );
         GlobalFree( lpof->hglbBuffer );
      }         
      return WBERR_DOS00 + of.nErrCode;
   }
   
   // update file name with name returned by OpenFile
   strcpy( lpof->szFileName, of.szPathName );
   
   lpof->fOpen = TRUE;
   lpof->uBufIndex = 0;

   if ( lpof->enumOpenMode == OM_INPUT ) {
   
      lpof->uByteCount = _lread( lpof->hFile, lpof->lpBuffer, (UINT)lpof->nRecLen );
      if ( lpof->uByteCount == HFILE_ERROR )
         return WBERR_READERROR;
         
      if ( lpof->uByteCount == 0 )
         lpof->fEof = TRUE;
      else
         lpof->fEof = FALSE;
   }
   
   return 0;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBCloseFile( OPENFILESTRUCT FAR* lpof )
{
   if ( lpof->fOpen ) {
      
      if ( lpof->enumOpenMode == OM_OUTPUT || lpof->enumOpenMode == OM_APPEND )
         WBFlushFileBuffer( lpof );
            
      _lclose( lpof->hFile );
      lpof->fOpen = FALSE;
      
      if ( lpof->enumOpenMode != OM_BINARY ) {         
         GlobalUnlock( lpof->hglbBuffer );
         GlobalFree( lpof->hglbBuffer );
      }         
   }
}   

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
BOOL WBEof( OPENFILESTRUCT FAR* lpof )
{
   return lpof->fEof;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBReadString( OPENFILESTRUCT FAR* lpof, HLSTR FAR* lphlstr )
{
   LPBYTE lpbyte;
   LPSTR lpszText;
   UINT uBytesLeft;
   USHORT usSizeOld, usSizeNew;
   UINT cb;
   ERR err;

   // Important note: WBReadString returns a non-temp string
   *lphlstr = WBCreateHlstr( NULL, 0 );
   if ( *lphlstr == NULL )
      return WBERR_STRINGSPACE;
      
   if ( lpof->enumOpenMode == OM_BINARY ) {

      err = WBResizeHlstr( *lphlstr, (USHORT)lpof->nRecLen );
      if ( err != 0 )
         return err;
            
      lpszText = WBDerefHlstr( *lphlstr );

      cb = _lread( lpof->hFile, lpszText, (UINT)lpof->nRecLen );
      if ( cb != (UINT)lpof->nRecLen )
         return WBERR_READERROR;
         
      if ( cb == 0 )
         lpof->fEof = TRUE;
   }
   else {
   
      usSizeOld = usSizeNew = 0;
      lpszText = NULL;
   
      for ( ;; ) {
      
         uBytesLeft = lpof->uByteCount - lpof->uBufIndex;
         assert( uBytesLeft != 0 );
         
         lpbyte = memchr( lpof->lpBuffer + lpof->uBufIndex, 
                          '\n', 
                          uBytesLeft );
      
         if ( lpbyte != NULL ) 
            cb = lpbyte - (lpof->lpBuffer + lpof->uBufIndex);
         else 
            cb = uBytesLeft;
            
         usSizeNew = usSizeOld + (USHORT)cb;
   
         if ( cb > 0 ) {         
            err = WBResizeHlstr( *lphlstr, usSizeNew );
            if ( err != 0 )
               return err;
            lpszText = WBDerefHlstr( *lphlstr );
            memcpy( lpszText + usSizeOld, 
                    lpof->lpBuffer + lpof->uBufIndex, 
                    cb ); 
         }
                 
         if ( lpbyte != NULL )
            cb++;
             
         lpof->uBufIndex += cb;
         
         if ( lpof->uBufIndex >= lpof->uByteCount ) {
         
            lpof->uByteCount = _lread( lpof->hFile, lpof->lpBuffer, (UINT)lpof->nRecLen );
            
            if ( lpof->uByteCount == HFILE_ERROR )
               return WBERR_READERROR;
            
            if ( lpof->uByteCount == 0 )
               lpof->fEof = TRUE;
               
            lpof->uBufIndex = 0;
            
         }
         
         if ( lpbyte != NULL || lpof->fEof )
            break;
            
         usSizeOld = usSizeNew;
      }
   
      if ( usSizeNew != 0 ) {
         // strip off trailing control characters such as \r and 
         // ctrl-Z (old-style logical end-of-file)
         lpszText += usSizeNew - 1;
         while ( usSizeNew != 0 && iscntrl( *lpszText ) ) {
            usSizeNew--;
            lpszText--;
         }
         WBResizeHlstr( *lphlstr, usSizeNew );
      }          
       
   }
   
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBWriteString( OPENFILESTRUCT FAR* lpof, HLSTR hlstr )
{
   HLSTR hlstrTemp;
   USHORT usSize;
   LPSTR lpszText, lpsz1, lpsz2;
   UINT uBytesWritten;
   UINT uBytesLeft;
   UINT uIndex;
   UINT cb; 
   
   if ( lpof->enumOpenMode == OM_BINARY ) {
   
      lpszText = WBDerefHlstrLen( hlstr, &usSize );
      if ( usSize != 0 ) {
         uBytesWritten = _lwrite( lpof->hFile, lpszText, (UINT)usSize );
         if ( uBytesWritten == HFILE_ERROR )
            return WBERR_WRITEERROR;
      }
   }
   else {
      // Create a temporary string as we need to add \r\n to the string
      // and we shouldn't modify the original 
      usSize = WBGetHlstrLen( hlstr );
      hlstrTemp = WBCreateTempHlstr( NULL, usSize + 2 );
      if ( hlstrTemp == NULL )
         return WBERR_STRINGSPACE;
         
      lpsz2 = WBDerefHlstr( hlstrTemp );
      lpsz1 = WBDerefHlstr( hlstr );
      memcpy( lpsz2, lpsz1, (size_t)usSize );
      *(lpsz2+usSize) = '\r';
      *(lpsz2+usSize+1) = '\n';
      usSize += 2;
      
      for ( uIndex = 0;; ) {   
   
         uBytesLeft = lpof->nRecLen - lpof->uBufIndex;   
         assert( uBytesLeft != 0 );
         
         cb = min( uBytesLeft, (UINT)usSize - uIndex );
   
         memcpy( lpof->lpBuffer + lpof->uBufIndex,
                 lpsz2 + uIndex,
                 cb );
                 
         lpof->uBufIndex += cb;
         uIndex += cb;
         
         if ( lpof->uBufIndex == (UINT)lpof->nRecLen ) {
            uBytesWritten = _lwrite( lpof->hFile,
                                     lpof->lpBuffer,
                                     (UINT)lpof->nRecLen );
            if ( uBytesWritten == HFILE_ERROR )
               return WBERR_WRITEERROR;
            lpof->uBufIndex = 0;
         }                                           
         
         if ( uIndex == (UINT)usSize )
            break;
      }
   
      // Clean up temporary string   
      WBDestroyHlstrIfTemp( hlstrTemp );
   }
   
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
static ERR WBFlushFileBuffer( OPENFILESTRUCT FAR* lpof )
{
   UINT uBytesWritten;
   
   uBytesWritten = _lwrite( lpof->hFile,
                            lpof->lpBuffer,
                            lpof->uBufIndex );
                            
   if ( uBytesWritten == HFILE_ERROR )
      return WBERR_WRITEERROR;
      
   return 0;
}

      
      
   
   
   
