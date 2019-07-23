#pragma warning( disable: 4001 ) /* 4001: Allow single line comments */

#define STRICT
#include <windows.h>
#include <toolhelp.h>
#include <wfext.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <direct.h>
#include <assert.h>
#include "vbapisub.h"
#include "resource.h"
#include "wtlres.h"

#include <nwcaldef.h>
#include <nwserver.h>
#include <nwerror.h>
#include <nwdpath.h>
#include <nwconnec.h>
#include <nwbindry.h>
#include <nwalias.h>

#define UNREFERENCED_PARAM(x)  ((void)(x))

char szAppTitle[] = "Novell Netware Services";
char szLoginServer[49];
char szLoginUserID[128];
char szLoginPassword[128];

HINSTANCE hinstDLL;

typedef ERR (WINAPI FAR* LPFUNCTION)( int, LPVAR, LPVAR );
typedef struct tagFUNCTIONINFO {
   NPSTR npszName;
   NPSTR npszTemplate;
   LPFUNCTION lpfn;
} FUNCTIONINFO;
typedef FUNCTIONINFO FAR* LPFUNCTIONINFO;
typedef int (CALLBACK FAR* ENUMCALLBACK)( LPSTR, LPSTR, LPFUNCTION );

ERR WINAPI WBFnAttach( int narg, LPVAR lparg, LPVAR lpResult );
ERR WINAPI WBFnDetach( int narg, LPVAR lparg, LPVAR lpResult );
ERR WINAPI WBFnGetFullName( int narg, LPVAR lparg, LPVAR lpResult );
ERR WINAPI WBFnGetInternetAddress( int narg, LPVAR lparg, LPVAR lpResult );
ERR WINAPI WBFnGetServerName( int narg, LPVAR lparg, LPVAR lpResult );
ERR WINAPI WBFnGetUserID( int narg, LPVAR lparg, LPVAR lpResult );
ERR WINAPI WBFnMapDel( int narg, LPVAR lparg, LPVAR lpResult );
ERR WINAPI WBFnMapDrive( int narg, LPVAR lparg, LPVAR lpResult );
ERR WINAPI WBFnMemberOf( int narg, LPVAR lparg, LPVAR lpResult );

FUNCTIONINFO FnTable[] = {
   "NETATTACH",            "S",     WBFnAttach,
   "NETDETACH",            "S",     WBFnDetach,
   "GETFULLNAME",          "",      WBFnGetFullName,
   "GETINTERNETADDRESS",   "",      WBFnGetInternetAddress,
   "GETSERVERNAME",        "",      WBFnGetServerName,
   "GETUSERID",            "S",     WBFnGetUserID,
   "MAPDEL",               "SI",    WBFnMapDel,
   "MAPDRIVE",             "SSI",   WBFnMapDrive,
   "MEMBEROF",             "S",     WBFnMemberOf
};

BOOL WBCheckForFilesInUse( int nDrive );
ERR WTSetErrorMessage( ERR err );
BOOL CALLBACK __export 
WTLoginDlgProc( HWND hwndDlg, UINT msg, WPARAM wParam, LPARAM lParam );

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
int FAR PASCAL 
LibMain( HANDLE hModule, WORD wDataSeg, WORD cbHeapSize, LPSTR lpszCmdLine)
{
   UNREFERENCED_PARAM( wDataSeg );
   UNREFERENCED_PARAM( lpszCmdLine );
   
   if ( cbHeapSize != 0 )
      UnlockData( 0 );
      
   hinstDLL = hModule;

   return 1;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
int FAR PASCAL _WEP ( int bSystemExit)
{
   UNREFERENCED_PARAM( bSystemExit );
   
   return 1;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WINAPI __export EnumFunctions( ENUMCALLBACK lpEnumFunc )
{
   int i;
   
   for ( i = 0; i < sizeof(FnTable) / sizeof(FUNCTIONINFO); i++ ) {
      if ( lpEnumFunc( FnTable[i].npszName, 
                       FnTable[i].npszTemplate, 
                       FnTable[i].lpfn ) == 0 )
         break;
   }
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WINAPI __export WBFnAttach( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpsz;
   USHORT usLen;
   NWCONN_HANDLE hconn;
   NWCONN_NUM nconn;
   NWCCODE ccode;
   WORD wObjectType;
   DWORD dwObjectID;
   BYTE chLoginTime[7];
   char szPrompt[256];
   BOOL fDefUserID;
   int ret;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 3 ) 
      return WBERR_ARGTOOMANY;
      
   assert( lparg[0].type == VT_STRING );
   lpsz = WBDerefHlstrLen( lparg[0].var.tString, &usLen );
   usLen = min( usLen, sizeof(szLoginServer) - 1 );
   _fmemcpy( szLoginServer, lpsz, usLen );
   szLoginServer[usLen] = '\0';
   _fstrupr( szLoginServer );
      
   if ( narg >= 2 ) {
      if ( lparg[1].type == VT_EMPTY ) {
         fDefUserID = TRUE;
      }
      else {
         if ( lparg[1].type != VT_STRING )
            return WBERR_TYPEMISMATCH;
         lpsz = WBDerefHlstrLen( lparg[1].var.tString, &usLen );
         usLen = min( usLen, sizeof(szLoginUserID) - 1 );
         _fmemcpy( szLoginUserID, lpsz, usLen );
         szLoginUserID[usLen] = '\0';
         fDefUserID = FALSE;
      }
   }
   else {
      fDefUserID = TRUE;
   }

   if ( fDefUserID ) {   
      ccode = NWGetPrimaryConnectionID( &hconn );
      if ( ccode == 0 )
         ccode = NWGetConnectionNumber( hconn, &nconn );
      if ( ccode == 0 )      
         ccode = NWGetConnectionInformation( hconn, nconn, szLoginUserID, 
                                             &wObjectType, 
                                             &dwObjectID, chLoginTime );
      if ( ccode != 0 )                                             
         szLoginUserID[0] = '\0';
         
      _fstrlwr( szLoginUserID );
   }

   if ( narg == 3 ) {
      if ( lparg[2].type != VT_STRING )
         return WBERR_TYPEMISMATCH;
         
      lpsz = WBDerefHlstrLen( lparg[2].var.tString, &usLen );
      usLen = min( usLen, sizeof(szPrompt) - 1 );
      _fmemcpy( szPrompt, lpsz, usLen );
      szPrompt[usLen] = '\0';
   }
   else {
      if ( LoadString( hinstDLL, IDS_LOGIN_PROMPT, szPrompt, sizeof(szPrompt )) == 0 )
         szPrompt[0] = '\0';
      _fstrcat( szPrompt, szLoginServer );
   }
   
   ret = DialogBoxParam( hinstDLL, 
                         MAKEINTRESOURCE(IDD_LOGIN),
                         WBGetHostWindow(),
                         (DLGPROC)WTLoginDlgProc,
                         (LPARAM)(LPSTR)szPrompt );
                    
   if ( ret == IDCANCEL ) {
      lpResult->var.tLong = WB_FALSE;
      lpResult->type = VT_I4;
      return 0;
   }
   
   _fstrupr( szLoginUserID );
   _fstrupr( szLoginPassword );

   ccode = NWAttachToFileServer( szLoginServer, 0, &hconn );
   switch ( ccode ) {
   
      case 0:
         break;
         
      case ALREADY_ATTACHED:
         {
            char szMessage[128];
            char szBuffer[128];
            int nYesNo;
            LoadString( hinstDLL, WTERR_ALREADY_ATTACHED, szMessage, sizeof(szMessage ));
            wsprintf( szBuffer, szMessage, (LPSTR)szLoginServer );
            MessageBeep( MB_ICONQUESTION );
            nYesNo = MessageBox( WBGetHostWindow(), szBuffer, szAppTitle, 
                                 MB_ICONQUESTION | MB_YESNO );
            if ( nYesNo == IDNO ) {
               lpResult->var.tLong = WB_FALSE;
               lpResult->type = VT_I4;
               return 0;
            }
            // Get connection handle of existing session
            ccode = NWGetConnectionHandle( (LPBYTE)szLoginServer, 0, &hconn, NULL );
            if ( ccode != 0 )
               return WTSetErrorMessage( WTERR_UNEXPECTED_ERR );

            // Logout from server if logged in               
            NWLogoutFromFileServer( hconn );
         }
         break;
         
      default:
         return WTSetErrorMessage( WTERR_UNEXPECTED_ERR );
   }
   
   ccode = NWLoginToFileServer( hconn, szLoginUserID, OT_USER, szLoginPassword );
   _fstrset( szLoginPassword, 0 );                            
      
   if ( ccode == 0 )
      lpResult->var.tLong = WB_TRUE;
   else {
      NWDetachFromFileServer( hconn );
      lpResult->var.tLong = WB_FALSE; 
   }
      
   lpResult->type = VT_I4;
   return 0;
}
      
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WINAPI __export WBFnDetach( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpsz;
   USHORT usLen;
   char szServerName[49];
   char szLocalName[3];
   char szRemoteName[_MAX_PATH];
   UINT cb;
   NWCONN_HANDLE hconn;
   NWCCODE ccode;
   LPSTR lpch;
   int i;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;
      
   assert( lparg[0].type == VT_STRING );
   lpsz = WBDerefHlstrLen( lparg[0].var.tString, &usLen );
   usLen = min( usLen, sizeof(szServerName) - 1 );
   _fmemcpy( szServerName, lpsz, (size_t)usLen );
   szServerName[usLen] = '\0';
   _fstrupr( szServerName );

   // Check for files in use by Windows on drives mapped on specified server   
   _fstrcpy( szLocalName, "A:" );
   for ( i = 0; i < 26; i++ ) {
      if ( GetDriveType( i ) == DRIVE_REMOTE ) {
         szLocalName[0] = (char)('A' + i);
         cb = sizeof( szRemoteName );
         WNetGetConnection( szLocalName, szRemoteName, &cb );
         lpch = szRemoteName;
         while ( *lpch == '\\' )
            lpch++;
         lpch = _fstrtok( lpch, ":\\" );
         if ( lpch && _fstrcmp( lpch, szServerName ) == 0 ) {
            if ( WBCheckForFilesInUse( i ) )
               return WTSetErrorMessage( WTERR_CANNOT_DETACH );
         }
      }
   }

   ccode = NWGetConnectionHandle( (LPBYTE)szServerName, 0, &hconn, NULL );
   if ( ccode == 0 )
      ccode = NWLogoutFromFileServer( hconn );
   if ( ccode == 0 )
      ccode = NWDetachFromFileServer( hconn );
      
   if ( ccode == 0 )
      lpResult->var.tLong = WB_TRUE;
   else
      lpResult->var.tLong = WB_FALSE;
   lpResult->type = VT_I4;
   return 0;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
BOOL WBCheckForFilesInUse( int nDrive )
{
   char chDrive;
   MODULEENTRY me;
   BOOL fMore;
   
   chDrive = (char)('A' + nDrive);
   me.dwSize = sizeof( MODULEENTRY );
   fMore = ModuleFirst( &me );
   
   while ( fMore ) {
      if ( me.szExePath[0] == chDrive )
         return TRUE;
      fMore = ModuleNext( &me );
   }
   
   return FALSE;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WINAPI __export WBFnGetFullName( int narg, LPVAR lparg, LPVAR lpResult )
{
   NWCONN_HANDLE hconn;
   NWCONN_NUM nconn;
   NWCCODE ccode;
   char szUserID[49];
   char szFullName[256];
   WORD wObjectType;
   DWORD dwObjectID;
   char chLoginTime[7];
   NWFLAGS fMore;
   
   UNREFERENCED_PARAM( lparg );
   
   if ( narg != 0 )
      return WBERR_EXPECT_NOARG;
      
   ccode = NWGetPrimaryConnectionID( &hconn );
   if ( ccode == 0 )
      ccode = NWGetConnectionNumber( hconn, &nconn );
   if ( ccode == 0 )      
      ccode = NWGetConnectionInformation( hconn, nconn, szUserID, &wObjectType, 
                                          &dwObjectID, chLoginTime );
   if ( ccode == 0 ) {                                          
      ccode = NWReadPropertyValue( hconn,
                                   szUserID,
                                   OT_USER,
                                   "IDENTIFICATION",
                                   1,
                                   szFullName,
                                   &fMore,
                                   NULL );
   }
   
   if ( ccode == 0 ) {
      lpResult->var.tString = WBCreateHlstr( szFullName, 
                                 (USHORT)_fstrlen(szFullName) );
   }
   else {
      lpResult->var.tString = WBCreateHlstr( NULL, 0 );
   }                                 
                                 
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
   lpResult->type = VT_STRING;
   return 0;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WINAPI __export WBFnGetInternetAddress( int narg, LPVAR lparg, LPVAR lpResult )
{  
   NWCONN_HANDLE hconn;
   NWCONN_NUM nconn;
   NWNET_ADDR bInternetAddress[10];
   NWCCODE ccode;  
   LPSTR lpch;
   BYTE bDigit;
   int i;
   
   UNREFERENCED_PARAM( lparg );
      
   if ( narg != 0 )
      return WBERR_EXPECT_NOARG;
      
   ccode = NWGetPrimaryConnectionID( &hconn );
   if ( ccode == 0 )
      ccode = NWGetConnectionNumber( hconn, &nconn );
   if ( ccode == 0 )
      ccode = NWGetInternetAddress( hconn, nconn, bInternetAddress );
      
   if ( ccode != 0 )
      return WTSetErrorMessage( WTERR_UNEXPECTED_ERR );
      
   lpResult->var.tString = WBCreateTempHlstr( NULL, 21 );
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
   lpch = WBDerefHlstr( lpResult->var.tString );
   
   for ( i = 0; i < 10; i++ ) {
   
      if ( i == 4 )
         *lpch++ = ':';
         
      bDigit = (BYTE)(bInternetAddress[i] >> 4);
      if ( bDigit > 9 ) 
         *lpch++ = (char)(bDigit - 10 + 'A');
      else
         *lpch++ = (char)(bDigit + '0');
         
      bDigit = (BYTE)(bInternetAddress[i] & 0x0F);
      if ( bDigit > 9 ) 
         *lpch++ = (char)(bDigit - 10 + 'A');
      else
         *lpch++ = (char)(bDigit + '0');
   }
 
   lpResult->type = VT_STRING;
   return 0;     
}   
      
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WINAPI __export WBFnGetServerName( int narg, LPVAR lparg, LPVAR lpResult )
{  
   NWCONN_HANDLE hconn;
   NWCCODE ccode;
   char szServerName[49];

   UNREFERENCED_PARAM( lparg );
      
   if ( narg != 0 )
      return WBERR_EXPECT_NOARG;
      
   ccode = NWGetPrimaryConnectionID( &hconn );
   if ( ccode == 0 )
      ccode = NWGetFileServerName( hconn, szServerName );
    
   if ( ccode != 0 )
      return WTSetErrorMessage( WTERR_UNEXPECTED_ERR );
      
   lpResult->var.tString = WBCreateHlstr( szServerName, (USHORT)_fstrlen(szServerName) );
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
   lpResult->type = VT_STRING;
   return 0;
}      
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WINAPI __export WBFnGetUserID( int narg, LPVAR lparg, LPVAR lpResult )
{  
   NWCONN_HANDLE hconn;
   NWCONN_NUM nconn;
   NWCCODE ccode;
   char szUserID[49];
   LPSTR lpsz;
   USHORT usLen;
   char szServerName[49];
   WORD wObjectType;
   DWORD dwObjectID;
   BYTE chLoginTime[7];
   
   UNREFERENCED_PARAM( lparg );
   
   if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   if ( narg == 1 ) {
      assert( lparg[0].type == VT_STRING );
      lpsz = WBDerefHlstrLen( lparg[0].var.tString, &usLen );
      usLen = min( usLen, sizeof(szServerName) - 1 );
      _fmemcpy( szServerName, lpsz, (size_t)usLen );
      szServerName[usLen] = '\0';
      _fstrupr( szServerName );
      ccode = NWGetConnectionHandle( (LPBYTE)szServerName, 0, &hconn, NULL );
      if ( ccode != 0 )
         return WTSetErrorMessage( WTERR_NO_CONNECTION );
   }
   else {
      ccode = NWGetPrimaryConnectionID( &hconn );
   }
   
   if ( ccode == 0 )
      ccode = NWGetConnectionNumber( hconn, &nconn );
   if ( ccode == 0 )      
      ccode = NWGetConnectionInformation( hconn, nconn, szUserID, &wObjectType, 
                                          &dwObjectID, chLoginTime );
                                          
   lpResult->var.tString = WBCreateTempHlstr( szUserID, (USHORT)_fstrlen(szUserID) );
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
   lpResult->type = VT_STRING;
   return 0;
}                                          

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WINAPI __export WBFnMapDel( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpsz;
   NWDRIVE_NUM drivenum;
   NWCCODE ccode;
   HWND hwnd;
   char chDrive;
   MODULEENTRY me;
   BOOL fMore;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 2 )
      return WBERR_ARGTOOMANY;
      
   assert( lparg[0].type == VT_STRING );
   lpsz = WBDerefZeroTermHlstr( lparg[0].var.tString );
   chDrive = (char)toupper(*lpsz);
   drivenum = (NWDRIVE_NUM)(chDrive - 'A' + 1);
   if ( drivenum > (NWDRIVE_NUM)('Z' - 'A' + 1) )
      return WTSetErrorMessage( WTERR_INVALID_DRIVENUM );

   // Check for file on target drive still in use by Windows
   me.dwSize = sizeof( MODULEENTRY );
   fMore = ModuleFirst( &me );
   while ( fMore ) {
      if ( me.szExePath[0] == chDrive ) 
         return WTSetErrorMessage( WTERR_DRIVE_BUSY );
      fMore = ModuleNext( &me );
   }

   if ( drivenum == (NWDRIVE_NUM)_getdrive() )
      _chdrive( (int)('C' - 'A' + 1) );
      
   ccode = NWDeleteDriveBase( drivenum, 0 );  

   if ( ccode == 0 ) {
      if ( narg == 2 ) {
         assert( lparg[1].type == VT_I4 );
         if ( lparg[1].var.tLong != 0 ) {
            if ( (hwnd = FindWindow( "WFS_Frame", NULL )) != NULL )
               SendMessage( hwnd, FM_REFRESH_WINDOWS, TRUE, 0L );
         }
       }
      lpResult->var.tLong = WB_TRUE;
   }
   else
      lpResult->var.tLong = WB_FALSE;
   lpResult->type = VT_I4;
   return 0;
}
      
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WINAPI __export WBFnMapDrive( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpsz, lpch;
   USHORT usLen;
   char szDirectory[_MAX_PATH];
   char szServerName[49];
   NWDRIVE_NUM drivenum;
   NWCONN_HANDLE hconn;
   NWCCODE ccode;
   HWND hwnd;
   
   if ( narg < 2 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 3 )
      return WBERR_ARGTOOMANY;
      
   assert( lparg[0].type == VT_STRING );
   lpsz = WBDerefZeroTermHlstr( lparg[0].var.tString );
   drivenum = (NWDRIVE_NUM)(toupper(*lpsz) - 'A' + 1);
   if ( drivenum > (NWDRIVE_NUM)('Z' - 'A' + 1) )
      return WTSetErrorMessage( WTERR_INVALID_DRIVENUM );
   
   assert( lparg[1].type == VT_STRING );
   lpsz = WBDerefZeroTermHlstr( lparg[1].var.tString );
   
   // Ignore leading spaces
   while ( isspace( *lpsz ) )
      lpsz++;

   // Check for new-style server name
   if ( _fstrcmp( lpsz, "\\\\" ) == 0 ) {
      lpsz += 2;
      if ( (lpch = _fstrchr( lpsz, '\\' )) == 0 )
         return WTSetErrorMessage( WTERR_INVALID_NETDIR );
      usLen = min( (USHORT)(lpch - lpsz), sizeof( szServerName) - 1 );
      _fmemcpy( szServerName, lpsz, (size_t)usLen );
      szServerName[usLen] = '\0';
      _fstrupr( szServerName );
      ccode = NWGetConnectionHandle( (LPBYTE)szServerName, 0, &hconn, NULL );
      lpsz = lpch+1;
   }
   else if ( (lpch = _fstrchr( lpsz, '/' )) != NULL ) {
      usLen = min( (USHORT)(lpch - lpsz), sizeof( szServerName) - 1 );
      _fmemcpy( szServerName, lpsz, (size_t)usLen );
      szServerName[usLen] = '\0';
      _fstrupr( szServerName );
      ccode = NWGetConnectionHandle( (LPBYTE)szServerName, 0, &hconn, NULL );
      lpsz = lpch+1;
   }
   else {
      ccode = NWGetPrimaryConnectionID( &hconn );
   }
   
   if ( ccode != 0 )
      return WTSetErrorMessage( WTERR_UNEXPECTED_ERR );
      
   usLen = min( (USHORT)_fstrlen( lpsz ), sizeof(szDirectory) - 1 );
   _fmemcpy( szDirectory, lpsz, (size_t)usLen );
   szDirectory[usLen] = '\0';
   _fstrupr( szDirectory );
   
   if ( drivenum == (NWDRIVE_NUM)_getdrive() )
      _chdrive( (int)('C' - 'A' + 1) );
      
   NWDeleteDriveBase( drivenum, 0 );  
   
   ccode = NWSetDriveBase( drivenum, hconn, 0, szDirectory, 0 );
   
   if ( ccode == 0 ) {
      if ( narg == 3 ) {
         assert( lparg[2].type == VT_I4 );
         if ( lparg[2].var.tLong != 0 ) {
            if ( (hwnd = FindWindow( "WFS_Frame", NULL )) != NULL )
               SendMessage( hwnd, FM_REFRESH_WINDOWS, TRUE, 0L );
         }
       }
      lpResult->var.tLong = WB_TRUE;
   }
   else
      lpResult->var.tLong = WB_FALSE;
      
   lpResult->type = VT_I4;
   return 0;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WINAPI __export WBFnMemberOf( int narg, LPVAR lparg, LPVAR lpResult )
{  
   LPSTR lpsz;
   USHORT usLen;
   NWCONN_HANDLE hconn;
   NWCONN_NUM nconn;
   NWCCODE ccode;
   char szGroup[49];
   char szUserID[49];
   WORD wObjectType;
   DWORD dwObjectID;
   BYTE chLoginTime[7];
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;
      
   assert( lparg[0].type == VT_STRING );
   lpsz = WBDerefHlstrLen( lparg[0].var.tString, &usLen );
   usLen = min( usLen, sizeof(szGroup) - 1 );
   _fmemcpy( szGroup, lpsz, usLen );
   szGroup[usLen] = '\0';
   _fstrupr( szGroup );
   
   ccode = NWGetPrimaryConnectionID( &hconn );
   if ( ccode == 0 )
      ccode = NWGetConnectionNumber( hconn, &nconn );
   if ( ccode == 0 )      
      ccode = NWGetConnectionInformation( hconn, nconn, szUserID, &wObjectType, 
                                          &dwObjectID, chLoginTime );
                                          
   if ( ccode != 0 )
      return WTSetErrorMessage( WTERR_UNEXPECTED_ERR );
                                                
   ccode = NWIsObjectInSet( hconn, szUserID, OT_USER, 
                            "GROUPS_I'M_IN", szGroup, OT_USER_GROUP );
   switch ( ccode ) {
      case 0:
         lpResult->var.tLong = WB_TRUE;
         break;
                                     
      case NO_SUCH_OBJECT:
      case NO_SUCH_MEMBER:
         lpResult->var.tLong = WB_FALSE;
         break;
         
      default:
         return WTSetErrorMessage( WTERR_UNEXPECTED_ERR );
   }
   
   lpResult->type = VT_I4;
   return 0;
} 

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
BOOL CALLBACK __export 
WTLoginDlgProc( HWND hwndDlg, UINT msg, WPARAM wParam, LPARAM lParam )
{

   switch ( msg ) {
   
      case WM_INITDIALOG:
         {             
         int nScreenWidth, nScreenHeight;
         int xPos, yPos;
         RECT rcParent, rcDialog;
         
         SetDlgItemText( hwndDlg, IDC_LOGINPROMPT, (LPSTR)lParam );
         SetDlgItemText( hwndDlg, IDC_USERID, szLoginUserID );
         SetDlgItemText( hwndDlg, IDC_PASSWORD, "" );
         
         GetWindowRect( hwndDlg, &rcDialog );
         nScreenWidth = GetSystemMetrics( SM_CXSCREEN );
         nScreenHeight = GetSystemMetrics( SM_CYSCREEN );
         
         xPos = (nScreenWidth - (rcDialog.right - rcDialog.left)) / 2;
         yPos = (nScreenHeight - (rcDialog.bottom - rcDialog.top)) / 3;
         
         xPos = max( xPos, 0 );
         xPos = min( xPos, nScreenWidth - (rcDialog.right - rcDialog.left));
         yPos = max( yPos, 0 );
         yPos = min( yPos, nScreenHeight - (rcDialog.bottom - rcDialog.top));
         
         GetClientRect( WBGetHostWindow(), &rcParent );
         xPos -= rcParent.left;
         yPos -= rcParent.top;
         
         MoveWindow( hwndDlg, 
                     xPos, 
                     yPos, 
                     rcDialog.right - rcDialog.left,
                     rcDialog.bottom - rcDialog.top,
                     FALSE );
                     
         if ( szLoginUserID[0] != '\0' )
            SetFocus( GetDlgItem( hwndDlg, IDC_PASSWORD ) );
         else
            SetFocus( GetDlgItem( hwndDlg, IDC_USERID ) );
         return FALSE;
         }
         
      case WM_COMMAND:
         switch ( wParam ) {
         
            case IDOK:
               GetDlgItemText( hwndDlg, IDC_USERID, szLoginUserID, 
                               sizeof(szLoginUserID)-1 );
               GetDlgItemText( hwndDlg, IDC_PASSWORD, szLoginPassword, 
                               sizeof(szLoginPassword)-1 );
               EndDialog( hwndDlg, IDOK );
               return TRUE;
               
            case IDCANCEL:
               EndDialog( hwndDlg, IDCANCEL );
               return TRUE;
         }
         break;
   }
   
   return FALSE;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WTSetErrorMessage( ERR err )
{
   static char szMessage[128];
   
   if ( LoadString( hinstDLL, (UINT)err, szMessage, sizeof(szMessage) ) == 0 )
      strcpy( szMessage, "Undefined error" );
      
   return WBSetErrorMessage( err, szMessage );
}