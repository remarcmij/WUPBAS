#define STRICT
#include <windows.h>
#include <commdlg.h>
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
#include "wbfndos.h"
#include "wbfnwin.h"

#define IOBUFFERSIZE      4096
#define MAXBUFFERSIZE     65500U

static char szDefFilter[] = "All files(*.*),*.*";

UINT CALLBACK  
FileOpenHookProc( HWND hwndDlg, UINT msg, WPARAM wParam, LPARAM lParam );
ERR WBFileCopy( LPSTR lpszSourcePath, LPSTR lpszDestPath );

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnFileCopy( int narg, LPVAR lparg, LPVAR lpResult )
{  
   LPSTR lpch, lpch1;
   char szSourceSpec[_MAX_PATH];
   char szDestSpec[_MAX_PATH];
   char szSourcePath[_MAX_PATH];
   char szDestPath[_MAX_PATH];
   char szTemplate[13];
   int i, ret;
   struct _find_t fileinfo;
   ERR err;
   
   if ( narg < 2 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 2 )
      return WBERR_ARGTOOMANY;
   
   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   lpch = WBDerefZeroTermHlstr( lparg[0].var.tString );
   if ( _fullpath( szSourceSpec, lpch, sizeof(szSourceSpec) ) == NULL )
      return WBERR_DOS03;
   _strupr( szSourceSpec );
   
   lpch = WBDerefZeroTermHlstr( lparg[1].var.tString );
   if ( _fullpath( szDestSpec, lpch, sizeof(szDestSpec) ) == NULL )
      return WBERR_DOS03;
   _strupr( szDestSpec );

   // Check if destination specifies a root directory
   if ( strcmp( &(szDestSpec[1]), ":\\" ) == 0 ) {
      strcat( szDestSpec, "*.*" );
   }
   else {
      // Check if destination specified an existing subdirectory
      ret = _dos_findfirst( szDestSpec, _A_SUBDIR, &fileinfo );
      if ( ret == 0 && (fileinfo.attrib & _A_SUBDIR) == _A_SUBDIR ) {
         strcat( szDestSpec, "\\*.*" );
      }
   }
   
   lpch = strrchr( szDestSpec, '\\' );
   assert( lpch != NULL );
   lpch++;
   i = 0;
   
   // Copy characters until first '*' in file name part
   while ( i < 8 && *lpch && *lpch != '.' && *lpch != '*' )
      szTemplate[i++] = *lpch++;
   
   // Ignore character after the first '*' up until the period or
   // end of string.
   while ( *lpch && *lpch != '.' )
      lpch++;
   
   if ( *lpch )
      lpch++;
   
   // Fill up pattern with '?' to 8 characters in file name part
   while ( i < 8 )
      szTemplate[i++] = '?';
   
   szTemplate[i++] = '.';
   
   // Copy characters until first '*' in file extension part
   if ( *lpch ) {
      while( i < 12 && *lpch && *lpch != '*' )
         szTemplate[i++] = *lpch++;
   }
   
   // Fill up pattern with '?' to 3 characters in file extension part
   while ( i < 12 )
      szTemplate[i++] = '?';
   
   szTemplate[i] = '\0';     

   // Copy files given by source to dest                         
   ret = _dos_findfirst( szSourceSpec, _A_NORMAL|_A_HIDDEN, &fileinfo );
   if ( ret != 0 )
      return WBERR_DOS02;
   
   WBShowWaitCursor();

   while ( ret == 0 ) {
   
      strcpy( szSourcePath, szSourceSpec );
      lpch = strrchr( szSourcePath, '\\' );
      assert( lpch != NULL );
      strcpy( lpch+1, fileinfo.name );

      strcpy( szDestPath, szDestSpec );
      
      lpch1 = strrchr( szDestPath, '\\' );
      assert( lpch1 != NULL );
      lpch1++;

      lpch = fileinfo.name;
         
      i = 0;
   
      while ( *lpch ) {
   
         if ( *lpch == '.' )
         {
            while (szTemplate[i] != '.')
            {
               if (szTemplate[i] != '?')
                  *lpch1++ = szTemplate[i];
               i++;
            }
         }
   
         if ( szTemplate[i] == '?' )
            *lpch1++ = *lpch;
         else
            *lpch1++ = szTemplate[i];
               
         i++;
         lpch++;
      }
   
      *lpch1 = '\0';

      if ( (err = WBFileCopy( szSourcePath, szDestPath )) != 0 ) {
         WBRestoreCursor();
         return err;
      }
               
      ret = _dos_findnext( &fileinfo );
   }
   
   WBRestoreCursor();
   
   lpResult->var.tLong = WB_TRUE;
   lpResult->type = VT_I4;

   return 0;
}   
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBFileCopy( LPSTR lpszSourcePath, LPSTR lpszDestPath )
{
   HFILE hfileSource, hfileDest;
   OFSTRUCT ofSource, ofDest;
   UINT cBytes;
   HGLOBAL hglbMem;
   LPSTR lpMem;
   UINT cbBuffer;
   WORD wDrive;
   struct _find_t findSource, findDest;
   VARIANT vArg, vTemp;
   int ret;
   ERR err;

   // Check if destination file is in use by Windows
   vArg.var.tString = WBCreateTempHlstr( lpszDestPath, 
                           (USHORT)strlen( lpszDestPath ) );
   if ( vArg.var.tString == NULL )
      return WBERR_STRINGSPACE;
   vArg.type = VT_STRING;
      
   if ( (err = WBFnFindModule( 1, &vArg, &vTemp )) != 0 )
      return err;
      
   WBDestroyHlstrIfTemp( vArg.var.tString );
      
   if ( vTemp.var.tLong != 0 ) 
      return WBERR_FILEINUSE;
  
   hfileSource = OpenFile( lpszSourcePath, &ofSource, OF_READ );
   if ( hfileSource == HFILE_ERROR )
      return WBERR_DOS00 + ofSource.nErrCode;
      
   ret = _dos_findfirst( lpszSourcePath, _A_NORMAL|_A_HIDDEN, &findSource );
   assert( ret == 0 );
   
   if ( findSource.size > MAXBUFFERSIZE )
      cbBuffer = MAXBUFFERSIZE;
   else
      cbBuffer = (UINT)findSource.size;

   // Must allocate at least 1 byte otherwise GlobalLock will fail      
   cbBuffer = max( 1, cbBuffer );
      
   if ( (hglbMem = GlobalAlloc( GMEM_MOVEABLE, cbBuffer )) == NULL ) {
      _lclose( hfileSource );
      return WBERR_OUTOFMEMORY;
   }

   if ( (lpMem = (LPSTR)GlobalLock( hglbMem )) == NULL ) {
      _lclose( hfileSource );
      GlobalFree( hglbMem );
      return WBERR_OUTOFMEMORY;
   }

   wDrive = (WORD)(*lpszDestPath - 'A' + 1);
   
   if ( _dos_findfirst( lpszDestPath, _A_NORMAL|_A_HIDDEN, &findDest ) == 0 )
      WBBroadcastEvent( DISKSPACE_CHANGE, wDrive, -findDest.size );
      
   hfileDest = OpenFile( lpszDestPath, &ofDest, OF_CREATE );
   if ( hfileDest == HFILE_ERROR ) {
      _lclose( hfileSource );
      GlobalUnlock( hglbMem );
      GlobalFree( hglbMem );
      return WBERR_DOS00 + ofDest.nErrCode;
   }


   // read and write blocks until complete file copied
   for ( ;; ) {
   
      cBytes = _lread( hfileSource, lpMem, cbBuffer );
      if ( cBytes == HFILE_ERROR ) {
         err = WBERR_READERROR;
         break;
      }
      else if ( cBytes != 0 ) {
         if ( _lwrite( hfileDest, lpMem, cBytes ) != cBytes ) {
            err = WBERR_DISKFULL;
            break;
         }
      }
      else {
         err = 0;
         break;
      }
   }

   // Set date, time and attributes of destination equal to source
   _dos_setftime( hfileDest, findSource.wr_date, findSource.wr_time );
   _lclose( hfileDest );
   _lclose( hfileSource );

   GlobalUnlock( hglbMem );
   GlobalFree( hglbMem );
   
   if ( err == 0 ) {
      _dos_setfileattr( lpszDestPath, findSource.attrib );
      WBBroadcastEvent( DISKSPACE_CHANGE, wDrive, findSource.size );
   }
   else {
      remove( lpszDestPath);
   }
   
   return err;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnFileDate( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszFileName;
   struct stat fstat;
   int ret;   
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   lpszFileName = WBDerefZeroTermHlstr( lparg[0].var.tString );
      
   if ( strpbrk( lpszFileName, "?*" ) != NULL )
      return WBERR_AMBIGUOUSFILENAME;
   
   ret = stat( lpszFileName, &fstat );
   if ( ret != 0 ) 
      return WBERR_FILENOTFOUND;
   else {
      lpResult->var.tLong = (long)fstat.st_mtime;
      lpResult->type = VT_DATE;
      return 0;
   }
}
                               
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnFileDelete( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszFileName;
   struct _find_t fileinfo;
   char szFileName[_MAX_PATH];
   LPSTR lpszBase;
   VARIANT vArg, vTemp;
   int len;
   WORD wDrive;
   ERR err;
   int ret;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   lpszFileName = WBDerefZeroTermHlstr( lparg[0].var.tString );
   
   while ( isspace( *lpszFileName ) )
      lpszFileName++;
   
   len = strlen( lpszFileName );
   if ( len >= 2 && lpszFileName[1] == ':' ) {
      assert( isalpha( lpszFileName[0] ) );
      wDrive = (WORD)(_toupper( lpszFileName[0] ) - 'A' + 1);
   }
   else {
      wDrive = (WORD)_getdrive();
   }
   
   strcpy( szFileName, lpszFileName );
   lpszBase = strrchr( szFileName, '\\' );
   if ( lpszBase == NULL )
      lpszBase = szFileName;
   else
      lpszBase++;

   ret = _dos_findfirst( lpszFileName, _A_NORMAL|_A_HIDDEN, &fileinfo );
   
   while ( ret == 0 ) {
      strcpy( lpszBase, fileinfo.name );
      
      vArg.var.tString = WBCreateTempHlstr( szFileName, (USHORT)strlen( szFileName ) );
      if ( vArg.var.tString == NULL )
         return WBERR_STRINGSPACE;
      vArg.type = VT_STRING;
      
      // Check if destination file is in use by Windows
      if ( (err = WBFnFindModule( 1, &vArg, &vTemp )) != 0 )
         return err;
         
      WBDestroyHlstrIfTemp( vArg.var.tString );
      
      if ( vTemp.var.tLong != 0 ) 
         return WBERR_FILEINUSE;
         
      if ( remove( szFileName ) != 0 )
         return WBERR_DOS00 + errno;

      WBBroadcastEvent( DISKSPACE_CHANGE, wDrive, -fileinfo.size );
      
      ret = _dos_findnext( &fileinfo );
   }   

   lpResult->var.tLong = WB_TRUE;
   lpResult->type = VT_I4;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnFileExist( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszFileName;
   struct _find_t fileinfo;
   UINT fuPrevErrorMode;
   int ret;   
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   lpszFileName = WBDerefZeroTermHlstr( lparg[0].var.tString );

   fuPrevErrorMode = SetErrorMode( SEM_FAILCRITICALERRORS );
   ret = _dos_findfirst( lpszFileName, 
                         _A_NORMAL|_A_HIDDEN|_A_SYSTEM,
                         &fileinfo );
   SetErrorMode( fuPrevErrorMode );                         
                         
   lpResult->var.tLong = ( ret == 0 ) ? WB_TRUE : WB_FALSE;
   lpResult->type = VT_I4;
   return 0;
}
                               
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnFileIsOpen( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszFileName;
   OFSTRUCT of;
   HFILE hfile;   
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   lpszFileName = WBDerefZeroTermHlstr( lparg[0].var.tString );
      
   if ( strpbrk( lpszFileName, "?*" ) != NULL )
      return WBERR_AMBIGUOUSFILENAME;
      
   lpResult->type = VT_I4;
   lpResult->var.tLong = WB_FALSE;
   
   hfile = OpenFile(lpszFileName, &of, OF_EXIST|OF_SHARE_EXCLUSIVE);
   if (hfile == HFILE_ERROR) 
   {
      switch (of.nErrCode) 
      {
         case 0x0002:   // File not found implies file is not open
            break;
            
         case 0x0020:   // Sharing violation: file is open
            lpResult->var.tLong = WB_TRUE;
            break;
            
         default:
            return WBERR_DOS00 + of.nErrCode;
      }
   }
   
   return 0;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnFileLen( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszFileName;
   struct stat fstat;
   int ret;   
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   lpszFileName = WBDerefZeroTermHlstr( lparg[0].var.tString );
      
   if ( strpbrk( lpszFileName, "?*" ) != NULL )
      return WBERR_AMBIGUOUSFILENAME;
   
   ret = stat( lpszFileName, &fstat );
   if ( ret != 0 ) 
      return WBERR_FILENOTFOUND;
   else {
      lpResult->var.tLong = (long)(fstat.st_size);
      lpResult->type = VT_I4;
      return 0;
   }
}
                               
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnFileRename( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszOldName;
   LPSTR lpszNewName;
   
   if ( narg < 2 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 2 )
      return WBERR_ARGTOOMANY;

   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   lpszOldName = WBDerefZeroTermHlstr( lparg[0].var.tString );
   
   assert( lparg[1].type == VT_STRING );
   lpszNewName = WBDerefZeroTermHlstr( lparg[1].var.tString );
      
   if ( rename( lpszOldName, lpszNewName ) != 0 )
      return WBERR_DOS00 + errno;

   lpResult->var.tLong = WB_TRUE;
   lpResult->type = VT_I4;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnGetAttr( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszFileName;
   unsigned attrib;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   lpszFileName = WBDerefZeroTermHlstr( lparg[0].var.tString );
      
   if ( _dos_getfileattr( lpszFileName, &attrib ) != 0 )
      return WBERR_DOS00 + errno;

   lpResult->var.tLong = (long)attrib;
   lpResult->type = VT_I4;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnSetAttr( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszFileName;
   unsigned attrib;
   
   if ( narg < 2 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 2 )
      return WBERR_ARGTOOMANY;

   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   lpszFileName = WBDerefZeroTermHlstr( lparg[0].var.tString );
   
   attrib = (unsigned)lparg[1].var.tLong;
      
   if ( _dos_setfileattr( lpszFileName, attrib ) != 0 )
      return WBERR_DOS00 + errno;

   lpResult->var.tLong = WB_TRUE;
   lpResult->type = VT_I4;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnLoadFile( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpsz;
   OPENFILESTRUCT FAR* lpof;
   VARIANT vTemp;
   USHORT usLen;
   USHORT iProp;
   ERR err;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   lpof = (OPENFILESTRUCT FAR*)WBLocalAlloc( LPTR, sizeof( OPENFILESTRUCT ));
   if ( lpof == NULL )
      return WBERR_OUTOFMEMORY;

   lpsz = WBDerefHlstrLen( lparg[0].var.tString, &usLen );
   usLen = min( usLen, sizeof(lpof->szFileName) - 1 );
   memcpy( lpof->szFileName, lpsz, (size_t)usLen );
   lpof->szFileName[usLen] = '\0';
   
   lpof->enumOpenMode = OM_INPUT;
   lpof->nRecLen = IOBUFFERSIZE;
   
   if ( (err = WBOpenFile( lpof )) != 0 )
      return err;
      
   if ( (err = WBCreateObject( "SLIST", lpResult )) != 0 )
      return err;

   vTemp.var.tString = WBCreateTempHlstr( lpof->szFileName, 
                                          (USHORT)strlen(lpof->szFileName));
   if ( vTemp.var.tString == NULL )
      return WBERR_STRINGSPACE;
                                                
   vTemp.type = VT_STRING;
   if ( (iProp = WBLookupProperty( lpResult->var.tCtl, "FILENAME" )) == (USHORT)-1 ) {
      assert( FALSE );
      return WBERR_INTERNAL_ERROR;
   }
      
   if ( (err = WBSetPropertyVariant( lpResult->var.tCtl, iProp, -1, &vTemp )) != 0 )
      return err;

   WBShowWaitCursor();
   WBEnableLocalCompact( FALSE );   
   
   while ( !WBEof( lpof ) ) {
   
      // Read next line
      if ( (err = WBReadString( lpof, &(vTemp.var.tString) )) != 0 )
         break;

      // Insert line into edit buffer 
      err = WBSListAppendLine( lpResult->var.tCtl, 
                               (LPVOID)(vTemp.var.tString), 
                               (USHORT)-1);
      if ( err != 0 )
         break;                                 
         
   }
   
   WBEnableLocalCompact( TRUE );

   WBCloseFile( lpof );
   WBLocalFree( (HLOCAL)LOWORD( lpof ));
   WBRestoreCursor();

   if ( err == 0 ) 
      WBSListSetModifiedFlag( lpResult->var.tCtl, FALSE );
      
   return err;
}

/*--------------------------------------------------------------------------*/
/* FileSelect( defdir, filter, title, deffile )                             */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnFileSelect( int narg, LPVAR lparg, LPVAR lpResult )
{
   OPENFILENAME ofn;
   char szFileName[_MAX_PATH];
   char szFileTitle[_MAX_PATH];
   char szDefExt[4];
   LPSTR lpsz;
   USHORT usLen;
   USHORT i;
   HLSTR hlstr = NULL;
   
   if ( narg > 4 )
      return WBERR_ARGTOOMANY;

   memset( &ofn, 0, sizeof(ofn) );
   ofn.lStructSize = sizeof(OPENFILENAME);
   ofn.hwndOwner = g_npTask->hwndClient;
   ofn.Flags = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_ENABLEHOOK;
   ofn.lpfnHook = FileOpenHookProc;
   
   // 1st arg: initial directory;
   if ( narg >= 1 ) {
      if ( lparg[0].type == VT_EMPTY ) {
         ofn.lpstrInitialDir = NULL;
      }
      else if ( lparg[0].type == VT_STRING ) {
         ofn.lpstrInitialDir = WBDerefZeroTermHlstr( lparg[0].var.tString );
      }
      else
         return WBERR_TYPEMISMATCH;
   }
   else {
      ofn.lpstrInitialDir = NULL;
   }
   
   // 2nd arg: filter
   if ( narg >= 2 ) {
      if ( lparg[1].type == VT_EMPTY ) {
         lpsz = szDefFilter;
         usLen = (USHORT)strlen( lpsz );
      }
      else if ( lparg[1].type == VT_STRING ) {
         lpsz = WBDerefHlstrLen( lparg[1].var.tString, &usLen );
      }
      else
         return WBERR_TYPEMISMATCH;
   }
   else {
      lpsz = szDefFilter;
      usLen = (USHORT)strlen( lpsz );
   }
      
   if ( *(lpsz + usLen - 1) != '|' ) {
      hlstr = WBCreateTempHlstr( lpsz, usLen + 1 );
      if ( hlstr == NULL )
         return WBERR_STRINGSPACE;
      lpsz = WBDerefHlstrLen( hlstr, &usLen );
      *(lpsz + usLen - 1) = '|';
      lpsz = WBDerefZeroTermHlstr( hlstr );
   }
   
   for ( i = 0; i < usLen; i++ ) {
      if ( *(lpsz+i) == '|' || *(lpsz+i) == ',' )
         *(lpsz+i) = '\0';
   }
   ofn.lpstrFilter = lpsz;
   ofn.nFilterIndex = 1;
   
   ofn.lpstrFileTitle = szFileTitle;
   ofn.nMaxFileTitle = sizeof( szFileTitle );
   
   // 3rd arg: title
   if ( narg >= 3 ) {
      if ( lparg[2].type == VT_EMPTY ) {
         ofn.lpstrTitle = "Select File";
      }
      else if ( lparg[2].type == VT_STRING ) {
         ofn.lpstrTitle = WBDerefZeroTermHlstr( lparg[2].var.tString );
      }
   }
   else {
      ofn.lpstrTitle = "Select File";
   }
   
   // 4th arg: default file name
   if ( narg == 4 ) {
      if( lparg[3].type != VT_STRING )
         return WBERR_TYPEMISMATCH;
      lpsz = WBDerefHlstrLen( lparg[3].var.tString, &usLen );
      usLen = min( usLen, sizeof(szFileName) - 1 );
      memcpy( szFileName, lpsz, (size_t)usLen );
      szFileName[usLen] = '\0';
   }
   else {
      szFileName[0] = '\0';
   }
   
   ofn.lpstrFile = szFileName;
   ofn.nMaxFile = sizeof( szFileName );
   
   lpsz = strchr( szFileName, '\\' );
   if ( lpsz != NULL )
      lpsz = strchr( lpsz+1, '.' );
   else
      lpsz = strchr( szFileName, '.' );
   if ( lpsz != NULL ) {
      usLen = min( (USHORT)strlen( lpsz+1 ), sizeof(szDefExt) - 1 );
      memcpy( szDefExt, lpsz+1, (size_t)usLen );
      szDefExt[usLen] = '\0';
   }
   else {
      szDefExt[0] = '\0';
   }
   
   ofn.lpstrDefExt = szDefExt;

   if ( GetOpenFileName( &ofn ) )
      lpResult->var.tString = WBCreateTempHlstr( szFileName, (USHORT)strlen( szFileName ));
   else
      lpResult->var.tString = WBCreateTempHlstr( NULL, 0 );

   if ( hlstr != NULL )   
      WBDestroyHlstrIfTemp( hlstr );
      
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
   lpResult->type = VT_STRING;
   return 0;
} 

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
UINT CALLBACK __export 
FileOpenHookProc( HWND hwndDlg, UINT msg, WPARAM wParam, LPARAM lParam )
{
   UNREFERENCED_PARAM( wParam );
   UNREFERENCED_PARAM( lParam );
   
   switch ( msg ) {
   
      case WM_INITDIALOG:
         {      
         int nScreenWidth, nScreenHeight;
         int xPos, yPos;
         RECT rcParent, rcDialog;
         
         GetWindowRect( hwndDlg, &rcDialog );
         nScreenWidth = GetSystemMetrics( SM_CXSCREEN );
         nScreenHeight = GetSystemMetrics( SM_CYSCREEN );
         
         xPos = (nScreenWidth - (rcDialog.right - rcDialog.left)) / 2;
         yPos = (nScreenHeight - (rcDialog.bottom - rcDialog.top)) / 3;
         xPos = max( xPos, 0 );
         xPos = min( xPos, nScreenWidth - (rcDialog.right - rcDialog.left));
         yPos = max( yPos, 0 );
         yPos = min( yPos, nScreenHeight - (rcDialog.bottom - rcDialog.top));
         
         GetClientRect( g_npTask->hwndClient, &rcParent );
         xPos -= rcParent.left;
         yPos -= rcParent.top;
         
         MoveWindow( hwndDlg, 
                     xPos, 
                     yPos, 
                     rcDialog.right - rcDialog.left,
                     rcDialog.bottom - rcDialog.top,
                     FALSE );
         return TRUE;
         }
   }
      
   return FALSE;
}
