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
#include <io.h>
#include "wupbas.h"
#include "wbfunc.h"
#include "wbslist.h"
#include "wbmem.h"
#include "wbfnwin.h"
#include "wbfndos.h"

extern WORD NEAR _F000H;

extern HINSTANCE g_hinstDLL;
extern char __near g_szAppTitle[];

static ERR SubTreeList( LPCSTR lpszDir, LPCSTR lpszFileSpec, LPVAR lpResult );
ERR WBDeleteDirectory( LPSTR lpszDirName, WORD wDrive );

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnChDir( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszDirName;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   lpszDirName = WBDerefZeroTermHlstr( lparg[0].var.tString );
   
   if ( _chdir( lpszDirName ) != 0 )
      return WBERR_DOS00 + ENOENT;
      
   lpResult->var.tLong = WB_TRUE;
   lpResult->type = VT_I4;
   return 0;
}      

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnChDrive( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszDrive;
   int nDrive;
   UINT newDrive;
   UINT cuDrives;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   lpszDrive = WBDerefZeroTermHlstr( lparg[0].var.tString );
   
   while ( isspace( *lpszDrive ) )
      lpszDrive++;
      
   if ( *lpszDrive != '\0' ) {
      nDrive = toupper( *lpszDrive ) - 'A' + 1;
      if ( nDrive < 1 || nDrive > 26 )
         return WBERR_ARGRANGE;
      
      _dos_setdrive( (unsigned)nDrive, &cuDrives );
      _dos_getdrive( &newDrive );
      
      if ( nDrive != (int)newDrive )
         return WBERR_INVALID_DRIVE;
   }
      
   lpResult->var.tLong = WB_TRUE;
   lpResult->type = VT_I4;
   return 0;
}      

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnCurDir( int narg, LPVAR lparg, LPVAR lpResult )
{
   char szBuffer[_MAX_PATH];
   
   UNREFERENCED_PARAM( lparg );
   
   if ( narg != 0 )
      return WBERR_EXPECT_NOARG;
      
   _getcwd( szBuffer, sizeof(szBuffer) );
   
   lpResult->var.tString = WBCreateTempHlstr(  szBuffer, (USHORT)strlen( szBuffer ));
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
   lpResult->type = VT_STRING;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnDir( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszFileSpec = NULL;
   unsigned attrib;
   struct _find_t FAR* lpfileinfo;
   int ret;
   
   if ( narg > 2 ) 
      return WBERR_ARGTOOMANY;
      
   if ( narg >= 1 ) {
      if ( lparg[0].type == VT_EMPTY )
         return WBERR_ARGRANGE;
      lpszFileSpec = WBDerefZeroTermHlstr( lparg[0].var.tString );
   }
   
   if ( narg == 2 ) {
      if ( lparg[1].type == VT_EMPTY )
         return WBERR_ARGRANGE;
      if ( lparg[1].var.tLong < 0 || lparg[1].var.tLong > USHRT_MAX )
         return WBERR_ARGRANGE;
      attrib = (unsigned)lparg[1].var.tLong;
   }
   else {
      attrib = _A_NORMAL;
   }
   
   // If volume is include, ignore all other attributes
   if ( (attrib & _A_VOLID) != 0 )
      attrib = _A_VOLID;

   lpfileinfo = &g_npContext->fileinfo;
   
   if ( narg == 0 && lpfileinfo->name[0] == '\0' )
      return WBERR_FUNCTIONCALL;
      
   if ( narg != 0 ) {
      assert( lpszFileSpec != NULL );
      ret = _dos_findfirst( lpszFileSpec, attrib, lpfileinfo );
   }
   else
      ret = _dos_findnext( lpfileinfo );
      
   if ( ret == 0 ) {
      lpResult->var.tString = WBCreateTempHlstr( lpfileinfo->name,
                                 (USHORT)strlen( lpfileinfo->name ) );
   }                                        
   else {
      memset( lpfileinfo, 0, sizeof(struct _find_t) );
      lpResult->var.tString = WBCreateTempHlstr( NULL, 0 );
   }
   
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
  lpResult->type = VT_STRING;
  return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnDirExist( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszDirName;
   char szDirName[_MAX_PATH];
   struct _find_t fileinfo;
   int ret;   
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   lpszDirName = WBDerefZeroTermHlstr( lparg[0].var.tString );
   
   while ( isspace( *lpszDirName ) )
      lpszDirName++;
   
   if ( strpbrk( lpszDirName, "?*" ) != NULL )
      return WBERR_AMBIGUOUSFILENAME;

   if ( _fullpath( szDirName, lpszDirName, sizeof( szDirName ) ) == NULL )
      return WBERR_DOS03;
   
   ret = _dos_findfirst( szDirName, 
                         _A_SUBDIR|_A_HIDDEN|_A_SYSTEM,
                         &fileinfo );
                         
   if ( ret != 0 || ( (fileinfo.attrib) & _A_SUBDIR ) != _A_SUBDIR ) 
      lpResult->var.tLong = WB_FALSE;
   else
      lpResult->var.tLong = WB_TRUE;
   lpResult->type = VT_I4;
   return 0;
}
                               
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnFileList( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszFileSpec;
   HLSTR hlstr;
   LPSTR lpsz;
   USHORT usLen;
   char szFullSpec[_MAX_PATH];
   char szPath[_MAX_PATH];
   USHORT usPathLen;
   struct _find_t fileinfo;
   unsigned attrib;
   int ret;
   VARIANT vTemp;
   ERR err;

   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 2 )
      return WBERR_ARGTOOMANY;

   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;

   if ( narg == 2 ) {      
      if ( lparg[1].var.tLong < 0 )
         return WBERR_ARGRANGE;
      attrib = (int)lparg[1].var.tLong;
   }
   else {
      attrib = _A_NORMAL;
   }
   
   // If volume is include, ignore all other attributes
   if ( (attrib & _A_VOLID) != 0 )
      attrib = _A_VOLID;

   if ( (err = WBCreateObject( "SLIST", lpResult )) != 0 )
      return err;
            
   lpszFileSpec = WBDerefZeroTermHlstr( lparg[0].var.tString );
   while ( isspace( *lpszFileSpec ) )
      lpszFileSpec++;

   _fullpath( szFullSpec, lpszFileSpec, sizeof( szFullSpec ) - 1  );
   _strupr( szFullSpec);
   
   lpsz = strrchr( szFullSpec, '\\' );
   assert( lpsz != NULL );
   usPathLen = (USHORT)(lpsz - szFullSpec + 1);
   strncpy( szPath, szFullSpec, (size_t)usPathLen );
   
   ret = _dos_findfirst( szFullSpec, attrib, &fileinfo );

   while ( ret == 0 ) {

      // File name
      usLen = (USHORT)strlen( fileinfo.name );
      hlstr = WBCreateTempHlstr( NULL, usPathLen + usLen );
      if ( hlstr == NULL )
         return WBERR_STRINGSPACE;
      lpsz = WBDerefHlstr( hlstr );
      memcpy( lpsz, szPath, usPathLen );
      memcpy( lpsz + usPathLen, fileinfo.name, usLen );
         
      // Add line to StrList object
      vTemp.var.tString = hlstr;
      vTemp.type = VT_STRING;
      if ( (err = WBSListAppendItem( lpResult->var.tCtl, &vTemp )) != 0 )
         return err;
      WBDestroyHlstrIfTemp( hlstr );
         
      ret = _dos_findnext( &fileinfo );
   }

   WBSListSetModifiedFlag( lpResult->var.tCtl, FALSE );
   return 0;
}
     
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnFileSearch( int narg, LPVAR lparg, LPVAR lpResult )
{
   char szPattern[MAXTEXT];
   LPSTR lpsz;
   USHORT usLen;
   LPSTR lpszFileSpec;
   ERR err;

   if ( narg < 2 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 2 )
      return WBERR_ARGTOOMANY;

   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;

   if ( (err = WBCreateObject( "SLIST", lpResult )) != 0 )
      return err;

   lpsz = WBDerefHlstrLen( lparg[0].var.tString, &usLen );
   usLen = min( usLen, sizeof( szPattern ) - 1);
   memcpy( szPattern, lpsz, usLen );
   szPattern[usLen] = '\0';
   _strupr( szPattern );
   
   if ( lparg[1].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   lpszFileSpec = WBDerefZeroTermHlstr( lparg[1].var.tString );
   
   while ( isspace( *lpszFileSpec ) )
      lpszFileSpec++;
   
   WBShowWaitCursor();
   WBEnableLocalCompact( FALSE );   
   
   if ( (err = SubTreeList( szPattern, lpszFileSpec, lpResult )) != 0 )
      return err;
      
   WBEnableLocalCompact( TRUE );
   WBRestoreCursor();
   
   WBSListSetModifiedFlag( lpResult->var.tCtl, FALSE );
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
static ERR SubTreeList( LPCSTR lpszPattern, LPCSTR lpszDir, LPVAR lpResult )
{
   char szFileSpec[_MAX_PATH];
   char szThisDir[_MAX_PATH];
   char szSubDir[_MAX_PATH];
   USHORT usPathLen;
   USHORT usLen;
   HLSTR hlstr;
   LPSTR lpsz;
   VARIANT vTemp;
   struct _find_t fileinfo;
   BOOL fExtraSlash;
   int ret;
   ERR err;
   
   strupr( strcpy( szThisDir, lpszDir ) );
   
   usPathLen = (USHORT)strlen( szThisDir );
   strcpy( szFileSpec, szThisDir );
   if ( szFileSpec[strlen(szFileSpec)-1] != '\\' ) {
      fExtraSlash = TRUE;
      strcat( szFileSpec, "\\" );
   }
   else {
      fExtraSlash = FALSE;
   }
   
   strcat( szFileSpec, lpszPattern );
   
   ret = _dos_findfirst( szFileSpec, _A_NORMAL|_A_SYSTEM|_A_HIDDEN, &fileinfo );
   while ( ret == 0 ) {
   
      // File name
      usLen = (USHORT)strlen( fileinfo.name );
      hlstr = WBCreateTempHlstr( NULL, usPathLen + usLen + (fExtraSlash ? 1 : 0) );
      if ( hlstr == NULL )
         return WBERR_STRINGSPACE;
      lpsz = WBDerefHlstr( hlstr );

      memcpy( lpsz, szThisDir, usPathLen );
      lpsz += usPathLen;
      if ( fExtraSlash )
         *lpsz++ = '\\';
      memcpy( lpsz, fileinfo.name, usLen );
         
      // Add line to StrList object
      vTemp.var.tString = hlstr;
      vTemp.type = VT_STRING;
      if ( (err = WBSListAppendItem( lpResult->var.tCtl, &vTemp )) != 0 )
         return err;
         
      WBDestroyHlstrIfTemp( hlstr );
      
      ret = _dos_findnext( &fileinfo );
   }
   
   lpsz = strrchr( szFileSpec, '\\' );
   assert( lpsz != NULL );
   strcpy( lpsz+1, "*.*" );
   
   ret = _dos_findfirst( szFileSpec, _A_SUBDIR|_A_SYSTEM|_A_HIDDEN, &fileinfo );
   while ( ret == 0 ) {
      if ( (fileinfo.attrib & _A_SUBDIR) && fileinfo.name[0] != '.' ) {
         strcpy( szSubDir, szThisDir );
         if ( szSubDir[strlen(szSubDir)-1] != '\\' )
            strcat( szSubDir, "\\" );
         strcat( szSubDir, fileinfo.name );
         if ( (err = SubTreeList( lpszPattern, szSubDir, lpResult )) != 0 )
            return err;
      }
      ret = _dos_findnext( &fileinfo );
   }
   
   return 0;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnEnviron( int narg, LPVAR lparg, LPVAR lpResult )
{
   USHORT usLen;
   LPSTR lpsz1, lpsz2;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   usLen = WBGetHlstrLen( lparg[0].var.tString );
   if ( usLen == 0 )
      return WBERR_ARGRANGE;

   lpsz1 = WBDerefZeroTermHlstr( lparg[0].var.tString );

   // getenv() is case sensitive. The argument to Environ() should be
   // specified in uppercase for normal environment variables. However
   // in order to also get the value of entries such as "windir" (lowercase)
   // it was chosen here not to routinely translate the argument to uppercase
   lpsz2 = getenv( lpsz1 );

   if ( lpsz2 == NULL )   
      lpResult->var.tString = WBCreateTempHlstr( NULL, 0 );
   else
      lpResult->var.tString = WBCreateTempHlstr( lpsz2, (USHORT)strlen(lpsz2));
   
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
   lpResult->type = VT_STRING;
   return 0;
}
     
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnEnvironList( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR FAR* lpszEnvString;
   VARIANT vTemp;
   ERR err;

   UNREFERENCED_PARAM( lparg );
      
   if ( narg != 0 )
      return WBERR_EXPECT_NOARG;

   if ( (err = WBCreateObject( "SLIST", lpResult )) != 0 )
      return err;
   
   lpszEnvString = _environ;
   vTemp.type = VT_STRING;
   
   while ( *lpszEnvString != NULL ) {
   
      vTemp.var.tString = WBCreateTempHlstr( *lpszEnvString, (USHORT)strlen(*lpszEnvString));
      if ( vTemp.var.tString == NULL )
         return WBERR_STRINGSPACE;

      if ( (err = WBSListAppendItem( lpResult->var.tCtl, &vTemp )) != 0 )
         return err;
         
      WBDestroyHlstrIfTemp( vTemp.var.tString );
   
      lpszEnvString++;
   }
   
   WBSListSetModifiedFlag( lpResult->var.tCtl, FALSE );
   
   return 0;
}
     
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnGetDiskFreeSpace( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszDrive;
   int nDrive;
   struct _diskfree_t diskspace;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;
   
   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;
      
   lpszDrive = WBDerefZeroTermHlstr( lparg[0].var.tString );
   while ( isspace( *lpszDrive ) )
      lpszDrive++;
      
   nDrive = toupper(*lpszDrive) - 'A' + 1;
   if (nDrive < 1 || nDrive > 26)
      return WBERR_ARGRANGE;
      
   _dos_getdiskfree(nDrive, &diskspace);
   
   lpResult->var.tLong = (long)diskspace.avail_clusters * 
                        (long)diskspace.sectors_per_cluster * 
                        (long)diskspace.bytes_per_sector;
   lpResult->type = VT_I4;
   
   return 0;
}   
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnGetDiskParms( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszDrive;
   int nDrive;
   struct _diskfree_t diskspace;
   char szBuffer[256];
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;
   
   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;
      
   lpszDrive = WBDerefZeroTermHlstr( lparg[0].var.tString );
   while ( isspace( *lpszDrive ) )
      lpszDrive++;
      
   nDrive = toupper(*lpszDrive) - 'A' + 1;
   if (nDrive < 1 || nDrive > 26)
      return WBERR_ARGRANGE;
      
   _dos_getdiskfree(nDrive, &diskspace);

   wsprintf(szBuffer, "%u,%u,%u,%u",
      diskspace.total_clusters,
      diskspace.avail_clusters,
      diskspace.sectors_per_cluster,
      diskspace.bytes_per_sector);
      
   lpResult->var.tString = WBCreateTempHlstr( szBuffer, (USHORT)strlen(szBuffer) );
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
   lpResult->type = VT_STRING;
   return 0;
}   
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnGetDosVersion( int narg, LPVAR lparg, LPVAR lpResult )
{
   DWORD dwVersion;
   char szBuffer[16];
   
   UNREFERENCED_PARAM( lparg );
   
   if ( narg != 0 )
      return WBERR_EXPECT_NOARG;
      
   dwVersion = GetVersion();
   wsprintf( szBuffer, "%d.%d", HIBYTE(HIWORD(dwVersion)),
                                LOBYTE(HIWORD(dwVersion)) );
                                
   lpResult->var.tString = WBCreateTempHlstr( szBuffer, (USHORT)strlen(szBuffer) );
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
   lpResult->type = VT_STRING;
   return 0;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnGetWinVersion( int narg, LPVAR lparg, LPVAR lpResult )
{
   DWORD dwVersion;
   char szBuffer[16];
   
   UNREFERENCED_PARAM( lparg );
   
   if ( narg != 0 )
      return WBERR_EXPECT_NOARG;
      
   dwVersion = GetVersion();
   wsprintf( szBuffer, "%d.%d", LOBYTE(LOWORD(dwVersion)),
                                HIBYTE(LOWORD(dwVersion)) );
                                
   lpResult->var.tString = WBCreateTempHlstr( szBuffer, (USHORT)strlen(szBuffer) );
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
   lpResult->type = VT_STRING;
   return 0;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnGetDriveType( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpsz;
   int nDrive;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;
      
   lpsz = WBDerefZeroTermHlstr( lparg[0].var.tString );
   nDrive = toupper( *lpsz ) - 'A';
   if ( nDrive < 0 || nDrive >= 26 )
      return WBERR_ARGRANGE;
      
   lpResult->var.tLong = (long)GetDriveType( nDrive );
   lpResult->type = VT_I4;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnMkDir( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszDirName;
   char szDirName[_MAX_PATH];
   WORD wDrive;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   lpszDirName = WBDerefZeroTermHlstr( lparg[0].var.tString );
   if ( _fullpath( szDirName, lpszDirName, sizeof(szDirName) ) == NULL )
      return WBERR_DOS05;
   
   if ( _mkdir( szDirName ) != 0 )
      return WBERR_DOS00 + errno;

   wDrive = (WORD)(szDirName[0] - 'A' + 1 );      
   WBBroadcastEvent( DISKSPACE_CHANGE, wDrive, 1L );
   
   lpResult->var.tLong = WB_TRUE;
   lpResult->type = VT_I4;
   return 0;
}      

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnRmDir( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszDirName;
   char szDirName[_MAX_PATH];
   WORD wDrive;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   lpszDirName = WBDerefZeroTermHlstr( lparg[0].var.tString );
   if ( _fullpath( szDirName, lpszDirName, sizeof(szDirName) ) == NULL )
      return WBERR_DOS05;
   
   if ( _rmdir( lpszDirName ) != 0 )
      return WBERR_DOS00 + errno;
      
   wDrive = (WORD)(szDirName[0] - 'A' + 1 );      
   WBBroadcastEvent( DISKSPACE_CHANGE, wDrive, -1L );
      
   lpResult->var.tLong = WB_TRUE;
   lpResult->type = VT_I4;
   return 0;
}      

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnTreeDelete( int narg, LPVAR lparg, LPVAR lpResult )
{  
   LPSTR lpszDirName;
   char szDirName[_MAX_PATH];
   BOOL fConfirm;
   struct _find_t fileinfo;
   int ret;
   char szMsgBuf1[MAXTEXT], szMsgBuf2[MAXTEXT];
   int nButton;
   WORD wDrive;
   ERR err;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 2 )
      return WBERR_ARGTOOMANY;

   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   lpszDirName = WBDerefZeroTermHlstr( lparg[0].var.tString );
   while ( isspace( *lpszDirName ) )
      lpszDirName++;
      
   if ( _fullpath( szDirName, lpszDirName, sizeof(szDirName) ) == NULL )
      return WBERR_DOS03;
      
   _strupr( szDirName );

   if ( narg == 2 ) {
      assert( lparg[1].type == VT_I4 );
      if ( lparg[1].var.tLong != 0L )
         fConfirm = TRUE;
      else
         fConfirm = FALSE;
   }
   else {
      fConfirm = TRUE;
   }
      
   if ( strcmp( &(szDirName[1]), ":\\" ) == 0 )
      return WBERR_DELETE_ROOTDIR;

   ret = _dos_findfirst( szDirName, _A_SUBDIR, &fileinfo );
   if ( ret != 0 || (fileinfo.attrib & _A_SUBDIR) != _A_SUBDIR )
      return WBERR_DOS03;

   wDrive = (WORD)(szDirName[0] - 'A' + 1);
   
   if ( fConfirm ) {
      LoadString( g_hinstDLL, IDS_TREEDELETE_CONFIRM, 
                  szMsgBuf1, sizeof( szMsgBuf1 ) );
      wsprintf( szMsgBuf2, szMsgBuf1, (LPSTR)szDirName );
      MessageBeep( MB_ICONEXCLAMATION );
      nButton = MessageBox( g_npTask->hwndClient, 
                            szMsgBuf2,
                            g_szAppTitle, 
                            MB_ICONEXCLAMATION+MB_OKCANCEL+MB_DEFBUTTON2 );
      if ( nButton == IDCANCEL ) {
         lpResult->var.tLong = WB_TRUE;                            
         lpResult->type = VT_I4;
         return 0;
      }
   }
   
   WBShowWaitCursor();
   err = WBDeleteDirectory( szDirName, wDrive );
   WBRestoreCursor();
   
   lpResult->var.tLong = WB_TRUE;                            
   lpResult->type = VT_I4;
   
   return err;
}   

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBDeleteDirectory( LPSTR lpszDirName, WORD wDrive )
{
   char szFileSpec[_MAX_PATH];
   char szDirEntry[_MAX_PATH];
   LPSTR lpszTag;
   VARIANT vArg, vTemp;
   struct _find_t fileinfo;
   int ret;
   ERR err;

   _strupr( strcpy( szFileSpec, lpszDirName ) );
   if ( szFileSpec[strlen(szFileSpec)-1] != '\\' )
      strcat( szFileSpec, "\\" );
      
   strcpy( szDirEntry, szFileSpec );
   lpszTag = strchr( szDirEntry, '\0' );
   
   strcat( szFileSpec, "*.*" );
   
   ret = _dos_findfirst( szFileSpec, _A_NORMAL|_A_HIDDEN|_A_SUBDIR, &fileinfo );

   while ( ret == 0 ) {
   
      strcpy( lpszTag, fileinfo.name );
      
      if ( ( fileinfo.attrib & _A_SUBDIR ) == _A_SUBDIR ) {
         if ( fileinfo.name[0] != '.' ) {
            if ( (err = WBDeleteDirectory( szDirEntry, wDrive )) != 0 )
               return err;
         }
      }
      else {
      
         vArg.var.tString = WBCreateTempHlstr( szDirEntry, (USHORT)strlen( szDirEntry ) );
         if  ( vArg.var.tString == NULL )
            return WBERR_STRINGSPACE;
         vArg.type = VT_STRING;
         
         // Check if destination file is in use by Windows
         if ( (err = WBFnFindModule( 1, &vArg, &vTemp )) != 0 )
            return err;
            
         if ( vTemp.var.tLong != 0 ) 
            return WBERR_FILEINUSE;
            
         WBDestroyHlstrIfTemp( vArg.var.tString );
         
         // Remove read-only attribute
         if ( (fileinfo.attrib & _A_RDONLY) == _A_RDONLY )
            _chmod( szDirEntry, _S_IREAD|_S_IWRITE );
            
         ret = remove( szDirEntry );
         if ( ret != 0 )
            return WBERR_DOS05;
            
         WBBroadcastEvent( DISKSPACE_CHANGE, wDrive, -fileinfo.size );
      }
      
      ret = _dos_findnext( &fileinfo );
   }

   ret = _rmdir( lpszDirName );
   if ( ret != 0 )
      return WBERR_DOS05;

   WBBroadcastEvent( DISKSPACE_CHANGE, wDrive, -1L );
   
   return 0;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnReadROMBios( int narg, LPVAR lparg, LPVAR lpResult )
{
   WORD __F000H = (WORD)(&_F000H);
   unsigned char FAR *lpch1;
   unsigned char FAR *lpch2;
   WORD wOffset;
   WORD i, cb;
   
   if ( narg < 2 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 2 )
      return WBERR_ARGTOOMANY;
      
   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   if ( lparg[0].var.tLong < 0L )
      return WBERR_ARGRANGE;
   else if ( lparg[0].var.tLong >= 0x10000L )
      return WBERR_ARGRANGE;
      
   if ( lparg[1].var.tLong < 0L )
      return WBERR_ARGRANGE;
      
   if ( lparg[0].var.tLong + lparg[1].var.tLong > 0x10000L )
      return WBERR_ARGRANGE;
   
   wOffset = (WORD)(lparg[0].var.tLong);
   cb = (WORD)(lparg[1].var.tLong);
   
   lpResult->var.tString = WBCreateTempHlstr( NULL, (USHORT)cb );
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
   lpch1 = (unsigned char FAR*)WBDerefZeroTermHlstr( lpResult->var.tString );
   memset( lpch1, 0, cb );
   
   lpch2 = (unsigned char FAR*)MAKELP(__F000H, wOffset);
   
   for ( i = 0; i < cb; i++ ) 
      *lpch1++ = *lpch2++;
      
   lpResult->type = VT_STRING;   
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnGetExtMemSize( int narg, LPVAR lparg, LPVAR lpResult )
{
   int MemorySize = 0;
   
   UNREFERENCED_PARAM( lparg );
  
   if ( narg != 0 )
      return WBERR_EXPECT_NOARG;

   // The following assemlby code was taken from a Microsoft Knowledge Base 
   // article. See the end of this file for the text of the article.
   _asm  { 
      xor      ax, ax 
      mov      al, 17h 
      out      70h, al 
      nop
      nop
      nop
      in       al, 71h 
      mov      MemorySize, ax 
      mov      al, 18h 
      out      70h, al 
      nop
      nop
      nop
      in       al, 71h 
      xchg     al, ah 
      or       MemorySize, ax 
   }

      
   lpResult->var.tLong = (long)MemorySize;
   lpResult->type = VT_I4;
   
   return 0;
}

/*---------------------------------------------------------------------------
Title: INF: Techniques to Find Extended Memory in a Windows DLL 
Document Number: Q75682           Publ Date: 26-MAY-1994 
Product Name: Microsoft Windows Software Development Kit 
Product Version: 3.00 3.10 
Operating System: WINDOWS



 ---------------------------------------------------------------------- 
 The information in this article applies to:


  - Microsoft Windows Software Development Kit (SDK) for Windows 
    versions 3.0 and 3.1 
 ----------------------------------------------------------------------


 SUMMARY 
 =======


 For a Windows application to determine the aount of extended memory 
 installed in the machine on which it is running, the extended memory word 
 can be retrieved from the machine's AT Real Time Clock/CMOS RAM. Typically, 
 80286, 80386 SX, 80386 DX, and i486 based computers store setup information 
 in CMOS. This article details how to access this data.


 MORE INFORMATION 
 ================


 The AT Real Time Clock/CMOS RAM is accessed via port 70h, and read from and 
 written to via port 71h. The extended memory word is stored low-order byte 
 first in the CMOS. The first (low order) byte is stored at address 17h, and 
 the second (high order) byte is stored at address 18h.


 To read a particular address in the CMOS, write the address to read to port 
 70h, then retrieve the information from port 71h. Figure 1 contains a code 
 fragment, written in Microsoft C version 6.0 with inline assembly, which 
 shows how to check the CMOS for extended memory size. For compatibility 
 with future versions of Windows, the following code should NOT be placed 
 directly into a Windows application. Instead, this code should be placed in 
 a dynamic-link library (DLL), which can be called by applications.


 Figure 1. Sample Extended Memory Size Check 
 -------------------------------------------


    int MemorySize = 0;


     _asm 
         { 
         xor   ax, ax 
         mov   al, 17h 
         out   70h, al 
         ; jmp   $+3               ; delay not needed in most cases 
         in    al, 71h 
         mov   MemorySize, ax 
         mov   al, 18h 
         out   70h, al 
         ; jmp   $+3               ; delay not needed in most cases 
         in    al, 71h 
         xchg  al, ah 
         or    MemorySize, ax 
         }


      printf("\nExtended Memory size = %u KB\n", MemorySize);


 The AT Real Time Clock/CMOS RAM configuration is documented in its entirety 
 in the IBM PC technical reference, on pages 1-56 to 1-68. The information 
 contained in the article was obtained from "The Programmer's PC Source 
 Book," by Thom Hogan, published by Microsoft Press (1988).


 Please note that although the BIOS on most machines automatically 
 configures the CMOS entry for extended memory size with the amount of 
 memory the BIOS power-on self-test (POST) routines find in the machine, 
 some BIOS setup programs allow the user to configure the extended memory 
 setting themselves. If the user has not filled in the correct number for 
 the amount of extended memory installed in the machine, the check described 
 in this article will be useless.


 WEXTMEM is a file in the Software Library that contains source code for a b Windows application that reports the amount of extended memory installed in 
 the machine on which the program is run. Download WEXTMEM.EXE, a self- 
 extracting file, from the Microsoft Software Library (MSL) on the following 
 services:


  - CompuServe 
       GO MSL 
       Search for WEXTMEM.EXE 
       Display results and download


  - Microsoft Download Service (MSDL) 
       Dial (206) 936-6735 to connect to MSDL 
       Download WEXTMEM.EXE


  - Internet (anonymous FTP) 
       ftp ftp.microsoft.com 
       Change to the \softlib\mslfiles directory 
       Get WEXTMEM.EXE 
 Additional reference words: 3.00 3.10 
 KBCategory: Prg 
 KBSubcategory: KrMm


COPYRIGHT Microsoft Corporation, 1994.
----------------------------------------------------------------------------*/