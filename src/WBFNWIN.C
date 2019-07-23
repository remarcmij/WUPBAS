#define STRICT
#include <windows.h>
#include <windowsx.h>
#include <toolhelp.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include "wupbas.h"
#include "wbfunc.h"
#include "wbslist.h"
#include "wbfnwin.h"

#define MAXOPTIONS      10
#define WM_EXITTASK     WM_USER+1

extern NPWBTASK g_npTaskFirst;
extern HINSTANCE g_hinstDLL;
extern char __near g_szAppTitle[];
extern HBRUSH g_hbrPrompt;

static LPSTR lpszNull = "";
static char szWaitClass[] = "WupBas:WaitShell";

typedef struct tagINPUTBOXSTRUCT {
   LPSTR lpszPrompt;
   LPSTR lpszTitle;
   LPSTR lpszDefault;
   UINT xPos, yPos;
} INPUTBOXSTRUCT;
typedef INPUTBOXSTRUCT FAR* LPINPUTBOXSTRUCT;

typedef struct tagOPTIONBOXSTRUCT {
   LPSTR lpszPrompt;
   LPSTR lpszTitle;
   LPSTR lpszOption[MAXOPTIONS];
   int   ciOption;
   int   nDefault;
   UINT xPos, yPos;
} OPTIONBOXSTRUCT;
typedef OPTIONBOXSTRUCT FAR* LPOPTIONBOXSTRUCT;

typedef struct tagCHECKBOXSTRUCT {
   LPSTR lpszPrompt;
   LPSTR lpszTitle;
   LPSTR lpszOption[MAXOPTIONS];
   int   ciOption;
   DWORD dwCheckBits;
   UINT xPos, yPos;
} CHECKBOXSTRUCT;
typedef CHECKBOXSTRUCT FAR* LPCHECKBOXSTRUCT;

typedef struct tagFINDWINDOWSTRUCT {
   LPSTR lpszTitle;
   HWND hwnd;
} FINDWINDOWSTRUCT;

BOOL CALLBACK __export 
WBInputBoxDlgProc( HWND hwndDlg, UINT msg, WPARAM wParam, LPARAM lParam );
BOOL CALLBACK __export 
WBOptionBoxDlgProc( HWND hwndDlg, UINT msg, WPARAM wParam, LPARAM lParam );
BOOL CALLBACK __export 
WBCheckBoxDlgProc( HWND hwndDlg, UINT msg, WPARAM wParam, LPARAM lParam );
LRESULT CALLBACK __export 
WBShellWaitWndProc( HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam );
BOOL CALLBACK __export WBNotifyCallback( WORD wID, DWORD dwData );
NPWBTASK WBFindTaskParent( HINSTANCE hinst );
BOOL CALLBACK __export EnumWndProc( HWND hwnd, LPARAM lParam );

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnBeep( int narg, LPVAR lparg, LPVAR lpResult )
{
   UINT uAlert;
   
   if ( narg > 1 )
      return WBERR_ARGTOOMANY;
      
   if ( narg == 1 )
      uAlert = (UINT)lparg[0].var.tLong;
   else 
      uAlert = 0;
      
   MessageBeep( uAlert );
   
   lpResult->var.tLong = WB_TRUE;
   lpResult->type = VT_I4;
   return 0;
}
         
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnCPUType( int narg, LPVAR lparg, LPVAR lpResult )
{
   DWORD dwFlags;

   UNREFERENCED_PARAM( lparg );
   
   if ( narg != 0 )
      return WBERR_EXPECT_NOARG;
   
   dwFlags = GetWinFlags();
   lpResult->var.tLong = (long)(dwFlags & (WF_CPU286|WF_CPU386|WF_CPU486));
   lpResult->type = VT_I4;
   return 0;
}
      
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnFindModule( int narg, LPVAR lparg, LPVAR lpResult )
{
   MODULEENTRY me;
   LPSTR lpsz1, lpsz2;
   BOOL fNoPath;
   BOOL fMore;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;
      
   lpsz1 = WBDerefZeroTermHlstr( lparg[0].var.tString );
   
   if ( strpbrk( lpsz1, "?*" ) != NULL )
      return WBERR_AMBIGUOUSFILENAME;
      
   if ( strchr( lpsz1, '\\' ) == NULL )
      fNoPath = TRUE;
   else
      fNoPath = FALSE;
      
   me.dwSize = sizeof( MODULEENTRY );
   
   fMore = ModuleFirst( &me );
   
   while ( fMore ) {
   
      if ( fNoPath ) {
         lpsz2 = strrchr( me.szExePath, '\\' );
         if ( lpsz2 != NULL )
            lpsz2++;
      }
      else {
         lpsz2 = me.szExePath;
      }
      
      if ( lpsz2 != NULL && _stricmp( lpsz1, lpsz2 ) == 0 ) {
         lpResult->var.tLong = (long)me.wcUsage;
         lpResult->type = VT_I4;
         return 0;
      }
      
      fMore = ModuleNext( &me );
   }
   
   lpResult->var.tLong = 0;
   lpResult->type = VT_I4;
   return 0;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnGetWinDir( int narg, LPVAR lparg, LPVAR lpResult )
{
   char szWinDir[_MAX_PATH];

   UNREFERENCED_PARAM( lparg );
      
   if ( narg != 0 )
      return WBERR_EXPECT_NOARG;
      
   GetWindowsDirectory( szWinDir, sizeof(szWinDir) );
   
   lpResult->var.tString = WBCreateTempHlstr( szWinDir, (USHORT)strlen(szWinDir) );
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
   lpResult->type = VT_STRING;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnGetSystemMetrics( int narg, LPVAR lparg, LPVAR lpResult )
{
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;
      
   lpResult->var.tLong = (long)GetSystemMetrics( (int)(lparg[0].var.tLong) );
   lpResult->type = VT_I4;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnIniDelete( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszSection, lpszKey, lpszIniFile;
   BOOL fOk;
   
   if ( narg < 2 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 3 )
      return WBERR_ARGTOOMANY;

   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   lpszSection = WBDerefZeroTermHlstr( lparg[0].var.tString );   
   
   while ( isspace( *lpszSection ) )
      lpszSection++;
      
   if ( *lpszSection == '\0' )
      return WBERR_ARGRANGE;

   if ( lparg[1].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   lpszKey = WBDerefZeroTermHlstr( lparg[1].var.tString );   
      
   while ( isspace( *lpszKey ) )
      lpszKey++;

   if ( *lpszKey == '\0' )
      return WBERR_ARGRANGE;

   if ( narg == 3 ) {
      lpszIniFile = WBDerefZeroTermHlstr( lparg[2].var.tString );   
      
      while ( isspace( *lpszIniFile ) )
         lpszIniFile++;
   
      if ( *lpszIniFile == '\0' )
         return WBERR_ARGRANGE;
      
      fOk = WritePrivateProfileString( lpszSection, lpszKey, NULL, lpszIniFile );
   }
   else {
      fOk = WriteProfileString( lpszSection, lpszKey, NULL );
   }
   
   lpResult->var.tLong = (long)( fOk ? WB_TRUE : WB_FALSE );
   lpResult->type = VT_I4;
   return 0;
}         

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnIniDeletePvt( int narg, LPVAR lparg, LPVAR lpResult )
{
   if ( narg < 3 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 3 )
      return WBERR_ARGTOOMANY;
      
   return WBFnIniDelete( narg, lparg, lpResult );
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnIniFlush( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszIniFile;
   BOOL fOk;
   
   if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   if ( narg == 1 ) {
      lpszIniFile = WBDerefZeroTermHlstr( lparg[0].var.tString );   
      while ( isspace( *lpszIniFile ) )
         lpszIniFile++;
   
      if ( *lpszIniFile == '\0' )
         return WBERR_ARGRANGE;
         
      fOk = WritePrivateProfileString( NULL, NULL, NULL, lpszIniFile );
   }
   else {
      fOk = WriteProfileString( NULL, NULL, NULL );
   }
   
   lpResult->var.tLong = (long)( fOk ? WB_TRUE : WB_FALSE );
   lpResult->type = VT_I4;
   return 0;
}         

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnIniFlushPvt( int narg, LPVAR lparg, LPVAR lpResult )
{  
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;
   
   return WBFnIniFlush( narg, lparg, lpResult );
}         

/*--------------------------------------------------------------------------*/
/*                                                                          */
/* IniList(section [,file] )                                                */
/*                                                                          */
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnIniList( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszSection, lpszIniFile;
   LPSTR lpszKey, lpszBuffer;
   char szText[MAXTEXT];
   char szValue[MAXTEXT];
   VARIANT vTemp;
   size_t len, len1;
   ERR err;

   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 2 )
      return WBERR_ARGTOOMANY;
   
   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   lpszSection = WBDerefZeroTermHlstr( lparg[0].var.tString );
   
   while ( isspace( *lpszSection ) )
      lpszSection++;
      
   if ( *lpszSection == '\0' )   
      return WBERR_ARGRANGE;
      
   if ( narg == 2 ) {
   
      lpszIniFile = WBDerefZeroTermHlstr( lparg[1].var.tString );
      
      while ( isspace( *lpszIniFile ) )
         lpszIniFile++;
         
      if ( *lpszIniFile == '\0' )   
         return WBERR_ARGRANGE;
   }
   else {
      lpszIniFile = NULL;
   }
   
   lpszBuffer = (LPSTR)WBLocalAlloc( LMEM_FIXED, 2048 );
   if ( LOWORD(lpszBuffer) == NULL )
      return WBERR_OUTOFMEMORY;
                           
   if ( (err = WBCreateObject( "SLIST", lpResult )) != 0 )
      return err;
      
   if ( lpszIniFile != NULL ) {
      GetPrivateProfileString( lpszSection, NULL, "", 
                               lpszBuffer, 2048,
                               lpszIniFile );
   }                               
   else {
      GetProfileString( lpszSection, NULL, "", lpszBuffer, 2048 );
   }      

   for ( lpszKey = lpszBuffer;
         *lpszKey != '\0'; lpszKey += strlen( lpszKey ) + 1 ) {

      if ( lpszIniFile != NULL ) {
         GetPrivateProfileString( lpszSection, lpszKey, "", 
                                  szValue, sizeof( szValue ),
                                  lpszIniFile );
      }                                  
      else {
         GetProfileString( lpszSection, lpszKey, "", 
                            szValue, sizeof( szValue ) );
      }

      len = min( strlen( lpszKey ), MAXTEXT-2 );
      memcpy( szText, lpszKey, len );
      szText[len++] = '=';
      szText[len] = '\0';
      
      len1 = min( strlen(szValue), MAXTEXT - len - 1 );
      if ( len1 > 0 ) {
         memcpy( szText + len, szValue, len1 );
         szText[len + len1] = '\0';
      }
         
      vTemp.var.tString = WBCreateTempHlstr( szText, 
                              (USHORT)strlen(szText) );
                              
      if ( vTemp.var.tString == NULL )
         return WBERR_STRINGSPACE;
      vTemp.type = VT_STRING;
      
      if ( (err = WBSListAppendItem( lpResult->var.tCtl, &vTemp )) != 0 )
         return err;
         
      WBDestroyHlstrIfTemp( vTemp.var.tString );
   }

   WBSListSetModifiedFlag( lpResult->var.tCtl, FALSE );
   
   WBLocalFree( (HLOCAL)LOWORD(lpszBuffer) );
   
   return 0;
}
     
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnIniRead( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszSection, lpszKey, lpszDefault, lpszIniFile;
   char szBuffer[256];
   USHORT usLen;

   if ( narg < 3 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 4 )
      return WBERR_ARGTOOMANY;
   
   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;
      
   lpszSection = WBDerefZeroTermHlstr( lparg[0].var.tString );   
   
   while ( isspace( *lpszSection ) )
      lpszSection++;
      
   if ( *lpszSection == '\0' )
      return WBERR_ARGRANGE;
      
   if ( lparg[1].type == VT_EMPTY )
      return WBERR_ARGRANGE;
      
   lpszKey = WBDerefZeroTermHlstr( lparg[1].var.tString );   
      
   while ( isspace( *lpszKey ) )
      lpszKey++;

   if ( *lpszKey == '\0' )
      return WBERR_ARGRANGE;

   if ( lparg[2].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   lpszDefault = WBDerefZeroTermHlstr( lparg[2].var.tString );

   if ( narg == 4 ) {
      if ( lparg[3].type == VT_EMPTY )
         return WBERR_ARGRANGE;
         
      lpszIniFile = WBDerefZeroTermHlstr( lparg[3].var.tString );   
      while ( isspace( *lpszIniFile ) )
         lpszIniFile++;
      if ( *lpszIniFile == '\0' )
         return WBERR_ARGRANGE;
         
      usLen = (USHORT)GetPrivateProfileString( 
                        lpszSection,
                        lpszKey,
                        lpszDefault,
                        szBuffer,
                        sizeof( szBuffer ),
                        lpszIniFile );
   }
   else {                        
      usLen = (USHORT)GetProfileString( 
                        lpszSection,
                        lpszKey,
                        lpszDefault,
                        szBuffer,
                        sizeof( szBuffer ) );
   }                        
         
   lpResult->var.tString = WBCreateTempHlstr( szBuffer, usLen );
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
   lpResult->type = VT_STRING;
   return 0;
}         

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnIniReadPvt( int narg, LPVAR lparg, LPVAR lpResult )
{
   if ( narg < 4 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 4 )
      return WBERR_ARGTOOMANY;

   return WBFnIniRead( narg, lparg, lpResult );
}   
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnIniWrite( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpszSection, lpszKey, lpszValue, lpszIniFile;
   BOOL fOk;
   
   if ( narg < 3 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 4 )
      return WBERR_ARGTOOMANY;
   
   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;
      
   lpszSection = WBDerefZeroTermHlstr( lparg[0].var.tString );   
   
   while ( isspace( *lpszSection ) )
      lpszSection++;
      
   if ( *lpszSection == '\0' )
      return WBERR_ARGRANGE;
      
   if ( lparg[1].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   lpszKey = WBDerefZeroTermHlstr( lparg[1].var.tString );   
      
   while ( isspace( *lpszKey ) )
      lpszKey++;

   if ( *lpszKey == '\0' )
      return WBERR_ARGRANGE;

   if ( lparg[2].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   lpszValue = WBDerefZeroTermHlstr( lparg[2].var.tString );
   
   if ( narg == 4 ) {
      if ( lparg[3].type == VT_EMPTY )
         return WBERR_ARGRANGE;
      lpszIniFile = WBDerefZeroTermHlstr( lparg[3].var.tString );   
      while ( isspace( *lpszIniFile ) )
         lpszIniFile++;
      if ( *lpszIniFile == '\0' )
         return WBERR_ARGRANGE;
         
      fOk = WritePrivateProfileString( lpszSection, lpszKey, 
                                       lpszValue,   lpszIniFile );
   }
   else {                                       
      fOk = WriteProfileString( lpszSection, lpszKey, lpszValue );
   }
   
   lpResult->var.tLong = (long)( fOk ? WB_TRUE : WB_FALSE );
   lpResult->type = VT_I4;
   return 0;
}         

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnIniWritePvt( int narg, LPVAR lparg, LPVAR lpResult )
{
   if ( narg < 4 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 4 )
      return WBERR_ARGTOOMANY;
   
   return WBFnIniWrite( narg, lparg, lpResult );
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnEnableLargeDialogs( int narg, LPVAR lparg, LPVAR lpResult )
{
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;
      
   if (lparg[0].var.tLong == 0)
      g_npTask->fLargeDialogs = FALSE;
   else
      g_npTask->fLargeDialogs = TRUE;
      
   lpResult->type = VT_I4;
   lpResult->var.tLong = WB_TRUE;
   return 0;
}
      
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnInputBox( int narg, LPVAR lparg, LPVAR lpResult )
{
   INPUTBOXSTRUCT InputBoxStruct;
   LPSTR lpszTemp;
   int idDialog;
   int ret;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 5 )
      return WBERR_ARGTOOMANY;

   memset( &InputBoxStruct, 0 , sizeof( INPUTBOXSTRUCT ));
   
   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;
      
   lpszTemp = WBDerefZeroTermHlstr( lparg[0].var.tString );
   InputBoxStruct.lpszPrompt = WBReplaceMetaChars(lpszTemp, TRUE);
   if (InputBoxStruct.lpszPrompt == NULL)
      return WBERR_OUTOFMEMORY;
   
   if ( narg >= 2 ) {
      if ( lparg[1].type == VT_EMPTY ) 
         InputBoxStruct.lpszTitle = g_szAppTitle;
      else 
         InputBoxStruct.lpszTitle = WBDerefZeroTermHlstr( lparg[1].var.tString );
   }
   else {
      InputBoxStruct.lpszTitle = g_szAppTitle;
   }   
      
   if ( narg >= 3 ) {
      if ( lparg[2].type == VT_EMPTY ) 
         InputBoxStruct.lpszTitle = lpszNull;
      else
         InputBoxStruct.lpszDefault = WBDerefZeroTermHlstr( lparg[2].var.tString );
   }
   else {
      InputBoxStruct.lpszDefault = lpszNull;
   }   
         
   if ( narg >= 4 ) {
   
      if ( narg != 5 )
         return WBERR_ARGTOOFEW;
               
      if ( lparg[3].type == VT_EMPTY )
         return WBERR_ARGRANGE;
      InputBoxStruct.xPos = (UINT)lparg[3].var.tLong;
            
      if ( lparg[4].type == VT_EMPTY )
         return WBERR_ARGRANGE;
      InputBoxStruct.yPos = (UINT)lparg[4].var.tLong;
   }

   if (g_npTask->fLargeDialogs)
      idDialog = IDD_INPUTBOX_EX;
   else
      idDialog = IDD_INPUTBOX;
         
   ret = DialogBoxParam( g_hinstDLL, 
                         MAKEINTRESOURCE(idDialog),
                         g_npTask->hwndClient,
                         (DLGPROC)WBInputBoxDlgProc,
          (LPARAM)(LPINPUTBOXSTRUCT)&InputBoxStruct );
                    
   WBRestoreTaskState();                    

   switch ( ret ) {
   
      case IDOK:
         lpResult->var.tString = g_npTask->hlstrInputBox;
         break;
         
      case IDCANCEL:
         lpResult->var.tString = WBCreateTempHlstr( NULL, 0 );
         if ( lpResult->var.tString == NULL )
            return WBERR_STRINGSPACE;
         break;
        
      default:
         assert( FALSE );
         return WBERR_INTERNAL_ERROR;
         
   }
   
   WBLocalFree((HLOCAL)LOWORD(InputBoxStruct.lpszPrompt));
   
   lpResult->type = VT_STRING;
   return 0;
}
      
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
BOOL CALLBACK __export 
WBInputBoxDlgProc( HWND hwndDlg, UINT msg, WPARAM wParam, LPARAM lParam )
{
   LPSTR lpsz;
   USHORT usLen;
   
   WBRestoreTaskState();
   
   switch ( msg ) {
   
      case WM_INITDIALOG:
         {             
         LPINPUTBOXSTRUCT lpInputBoxStruct;
         int nScreenWidth, nScreenHeight;
         int xPos, yPos;
         int nLogPixelsX, nLogPixelsY;
         RECT rcParent, rcDialog;
         HDC hdc;
         
         lpInputBoxStruct = (LPINPUTBOXSTRUCT)lParam;
         SetWindowText( hwndDlg, lpInputBoxStruct->lpszTitle );
         SetDlgItemText( hwndDlg, IDC_EDIT, lpInputBoxStruct->lpszPrompt );
         SetDlgItemText( hwndDlg, IDC_INPUT, lpInputBoxStruct->lpszDefault );
         SendDlgItemMessage( hwndDlg, IDC_INPUT, EM_SETSEL, 0, MAKELPARAM(32767,32767) );
         
         GetWindowRect( hwndDlg, &rcDialog );
         nScreenWidth = GetSystemMetrics( SM_CXSCREEN );
         nScreenHeight = GetSystemMetrics( SM_CYSCREEN );
         
         if ( lpInputBoxStruct->xPos != 0 ) {         
            hdc = GetDC( hwndDlg );
            nLogPixelsX = GetDeviceCaps( hdc, LOGPIXELSX );
            nLogPixelsY = GetDeviceCaps( hdc, LOGPIXELSY );
            ReleaseDC( hwndDlg, hdc );
            xPos = (int)((long)lpInputBoxStruct->xPos * 
                        (long)nLogPixelsX / 1440L);
            yPos = (int)((long)lpInputBoxStruct->yPos * 
                        (long)nLogPixelsY / 1440L);
         }
         else {
            xPos = (nScreenWidth - (rcDialog.right - rcDialog.left)) / 2;
            yPos = (nScreenHeight - (rcDialog.bottom - rcDialog.top)) / 3;
         }
         
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

         SetFocus( GetDlgItem( hwndDlg, IDC_INPUT ) );
         return FALSE;
         }
         
      case WM_COMMAND:
         switch ( wParam ) {
         
            case IDOK:
               usLen = (USHORT)SendDlgItemMessage( hwndDlg, 
                                                   IDC_INPUT, 
                                                   WM_GETTEXTLENGTH, 
                                                   0, 
                                                   0L );
               g_npTask->hlstrInputBox = WBCreateTempHlstr( NULL, usLen );
               if ( usLen > 0 ) {
                  lpsz = WBDerefHlstr( g_npTask->hlstrInputBox );
                  GetDlgItemText( hwndDlg, IDC_INPUT, lpsz, (int)usLen+1 );
               }
               
               EndDialog( hwndDlg, IDOK );
               return TRUE;
               
            case IDCANCEL:
               EndDialog( hwndDlg, IDCANCEL );
               return TRUE;
         }
         break;
         
      case WM_CTLCOLOR:
         if (g_npTask->fLargeDialogs)
         {
            if ((HWND)LOWORD(lParam) == GetDlgItem(hwndDlg, IDC_EDIT)) {
               SetBkColor((HDC)wParam, RGB(192, 192, 192));
               return (BOOL)g_hbrPrompt;
            }
         }
         break;
   }
   
   return FALSE;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/* OptionBox( prompt, title, optionstring [, [default][, xpos, ypos]] )     */
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnOptionBox( int narg, LPVAR lparg, LPVAR lpResult )
{
   OPTIONBOXSTRUCT OptionBoxStruct;
   USHORT usLen;
   NPSTR npszTemp2;
   LPSTR lpszTemp;
   LPSTR lpszOption;
   int idDialog;
   int i;
   int ret;
   
   if ( narg < 3 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 6 )
      return WBERR_ARGTOOMANY;

   memset( &OptionBoxStruct, 0 , sizeof( OPTIONBOXSTRUCT ));

   // 1st arg: Prompt   
   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;
      
   lpszTemp = WBDerefZeroTermHlstr( lparg[0].var.tString );
   OptionBoxStruct.lpszPrompt = WBReplaceMetaChars(lpszTemp, TRUE);
   if (OptionBoxStruct.lpszPrompt == NULL)
      return WBERR_OUTOFMEMORY;
      
   // 2ng arg: Title
   if ( lparg[1].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   OptionBoxStruct.lpszTitle = WBDerefZeroTermHlstr( lparg[1].var.tString );

   // 3rd arg: Options      
   if ( lparg[2].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   usLen = WBGetHlstrLen( lparg[2].var.tString );
   npszTemp2 = (NPSTR)WBLocalAlloc( LPTR, usLen + 1 );
   if ( npszTemp2 == NULL )
      return WBERR_OUTOFMEMORY;
   
   strcpy( npszTemp2, WBDerefZeroTermHlstr( lparg[2].var.tString ));

   lpszOption = strtok( npszTemp2, "|" );
   
   for ( i = 0; i < MAXOPTIONS; i++ ) {
      if ( lpszOption == NULL )
         break;
      OptionBoxStruct.lpszOption[i] = lpszOption;
      lpszOption = strtok( NULL, "|" );
   }
   
   if ( i == 0 ) 
      return WBERR_NOOPTIONS;
      
   OptionBoxStruct.ciOption = i;

   if ( narg >= 4 ) {
      // 4th arg: Default
      if ( lparg[3].type == VT_EMPTY ) {
         OptionBoxStruct.nDefault = 0;
      }
      else {
         if ( lparg[3].var.tLong < 1 || lparg[3].var.tLong > MAXOPTIONS )
            return WBERR_ARGRANGE;
         OptionBoxStruct.nDefault = (int)lparg[3].var.tLong - 1;
      }
   }
   else {
      OptionBoxStruct.nDefault = 0;
   }
         
   if ( narg >= 5 ) {
   
      if ( narg != 6 )
         return WBERR_ARGTOOFEW;
               
      if ( lparg[4].type == VT_EMPTY )
         return WBERR_ARGRANGE;
      OptionBoxStruct.xPos = (UINT)lparg[4].var.tLong;
      OptionBoxStruct.yPos = (UINT)lparg[5].var.tLong;
   }
   
   if (g_npTask->fLargeDialogs)
      idDialog = IDD_OPTIONBOX_EX;
   else
      idDialog = IDD_OPTIONBOX;
         
   ret = DialogBoxParam( g_hinstDLL, 
                         MAKEINTRESOURCE(idDialog),
                         g_npTask->hwndClient,
                         (DLGPROC)WBOptionBoxDlgProc,
                         (LPARAM)(LPOPTIONBOXSTRUCT)&OptionBoxStruct );
                    
   WBRestoreTaskState();                    

   switch ( ret ) {
   
      case IDOK:
         lpResult->var.tLong = g_npTask->lOption;
         break;
         
      case IDCANCEL:
         lpResult->var.tLong = 0;
         break;
        
      default:
         assert( FALSE );
         return WBERR_INTERNAL_ERROR;
         
   }
   
   WBLocalFree( (HLOCAL)npszTemp2 );
   WBLocalFree( (HLOCAL)LOWORD(OptionBoxStruct.lpszPrompt));
   
   lpResult->type = VT_I4;
   return 0;
}
      
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
BOOL CALLBACK __export 
WBOptionBoxDlgProc( HWND hwndDlg, UINT msg, WPARAM wParam, LPARAM lParam )
{
   int i;
   
   WBRestoreTaskState();
   
   switch ( msg ) {
   
      case WM_INITDIALOG:
         {             
         LPOPTIONBOXSTRUCT lpOptionBoxStruct;
         int nScreenWidth, nScreenHeight;
         int xPos, yPos;
         int nLogPixelsX, nLogPixelsY;
         RECT rcParent, rcDialog;
         HWND hwndButton;
         HDC hdc;
         
         lpOptionBoxStruct = (LPOPTIONBOXSTRUCT)lParam;
         SetWindowText( hwndDlg, lpOptionBoxStruct->lpszTitle );
         SetDlgItemText( hwndDlg, IDC_EDIT, lpOptionBoxStruct->lpszPrompt );
         
         for ( i = 0; i < lpOptionBoxStruct->ciOption; i++ ) {
            hwndButton = GetDlgItem( hwndDlg, IDC_RADIO1 + i );
            SetWindowText( hwndButton,
                           lpOptionBoxStruct->lpszOption[i] );
         }
          
         // Disable and hide unused options
         for ( ; i < MAXOPTIONS; i++ ) {
            hwndButton = GetDlgItem( hwndDlg, IDC_RADIO1 + i );
            EnableWindow( hwndButton, FALSE );
            ShowWindow( hwndButton, SW_HIDE );
         }
         
         g_npTask->lOption = (long)(lpOptionBoxStruct->nDefault + 1);
         SendDlgItemMessage( hwndDlg,
                             IDC_RADIO1 + (WORD)g_npTask->lOption - 1,
                             BM_SETCHECK,
                             TRUE,
                             0L );
                             
         GetWindowRect( hwndDlg, &rcDialog );
         nScreenWidth = GetSystemMetrics( SM_CXSCREEN );
         nScreenHeight = GetSystemMetrics( SM_CYSCREEN );
         
         if ( lpOptionBoxStruct->xPos != 0 ) {         
            hdc = GetDC( hwndDlg );
            nLogPixelsX = GetDeviceCaps( hdc, LOGPIXELSX );
            nLogPixelsY = GetDeviceCaps( hdc, LOGPIXELSY );
            ReleaseDC( hwndDlg, hdc );
            xPos = (int)((long)lpOptionBoxStruct->xPos * 
                        (long)nLogPixelsX / 1440L);
            yPos = (int)((long)lpOptionBoxStruct->yPos * 
                        (long)nLogPixelsY / 1440L);
         }
         else {
            xPos = (nScreenWidth - (rcDialog.right - rcDialog.left)) / 2;
            yPos = (nScreenHeight - (rcDialog.bottom - rcDialog.top)) / 3;
         }
         
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

         SetFocus( GetDlgItem( hwndDlg, IDOK ));
         return FALSE;
         }
         
      case WM_COMMAND:
         switch ( wParam ) {
         
            case IDC_RADIO1:
            case IDC_RADIO2:
            case IDC_RADIO3:
            case IDC_RADIO4:
            case IDC_RADIO5:
            case IDC_RADIO6:
            case IDC_RADIO7: 
            case IDC_RADIO8: 
            case IDC_RADIO9: 
            case IDC_RADIO10: 
               g_npTask->lOption = (long)(wParam - IDC_RADIO1 + 1);
               return TRUE;           
               
            case IDOK:
               EndDialog( hwndDlg, IDOK );
               return TRUE;
               
            case IDCANCEL:
               g_npTask->lOption = 0;
               EndDialog( hwndDlg, IDCANCEL );
               return TRUE;
               
         }
         break;
         
      case WM_CTLCOLOR:
         if (g_npTask->fLargeDialogs)
         {
            if ((HWND)LOWORD(lParam) == GetDlgItem(hwndDlg, IDC_EDIT)) {
               SetBkColor((HDC)wParam, RGB(192, 192, 192));
               return (BOOL)g_hbrPrompt;
            }
         }
         break;
         
   }
   
   return FALSE;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/* CheckBox( prompt, title, optionstring [, [default][, xpos, ypos]] )      */
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnCheckBox( int narg, LPVAR lparg, LPVAR lpResult )
{
   CHECKBOXSTRUCT CheckBoxStruct;
   USHORT usLen;
   NPSTR npszTemp2;
   LPSTR lpszTemp;
   LPSTR lpszOption;
   int idDialog;
   int i;
   int ret;
   
   if ( narg < 3 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 6 )
      return WBERR_ARGTOOMANY;

   memset( &CheckBoxStruct, 0 , sizeof( CHECKBOXSTRUCT ));

   // 1st arg: Prompt   
   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;

   lpszTemp = WBDerefZeroTermHlstr( lparg[0].var.tString );
   CheckBoxStruct.lpszPrompt = WBReplaceMetaChars(lpszTemp, TRUE);
   if (CheckBoxStruct.lpszPrompt == NULL)
      return WBERR_OUTOFMEMORY;

   // 2ng arg: Title
   if ( lparg[1].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   CheckBoxStruct.lpszTitle = WBDerefZeroTermHlstr( lparg[1].var.tString );

   // 3rd arg: Options      
   if ( lparg[2].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   usLen = WBGetHlstrLen( lparg[2].var.tString );
   npszTemp2 = (NPSTR)WBLocalAlloc( LPTR, usLen + 1 );
   if ( npszTemp2 == NULL )
      return WBERR_OUTOFMEMORY;
   
   strcpy( npszTemp2, WBDerefZeroTermHlstr( lparg[2].var.tString ));

   lpszOption = strtok( npszTemp2, "|" );
   
   for ( i = 0; i < MAXOPTIONS; i++ ) {
      if ( lpszOption == NULL )
         break;
      CheckBoxStruct.lpszOption[i] = lpszOption;
      lpszOption = strtok( NULL, "|" );
   }
   
   if ( i == 0 ) 
      return WBERR_NOOPTIONS;
      
   CheckBoxStruct.ciOption = i;
   
   // 4rd arg: Default
   if ( lparg[3].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   CheckBoxStruct.dwCheckBits = (DWORD)lparg[3].var.tLong;
   
   if ( narg >= 5 ) {
   
      if ( narg != 6 )
         return WBERR_ARGTOOFEW;
               
      if ( lparg[4].type == VT_EMPTY )
         return WBERR_ARGRANGE;
      CheckBoxStruct.xPos = (UINT)lparg[4].var.tLong;
      CheckBoxStruct.yPos = (UINT)lparg[5].var.tLong;
   }
   
   if (g_npTask->fLargeDialogs)
      idDialog = IDD_CHECKBOX_EX;
   else
      idDialog = IDD_CHECKBOX;
         
   ret = DialogBoxParam( g_hinstDLL, 
                         MAKEINTRESOURCE(idDialog),
                         g_npTask->hwndClient,
                         (DLGPROC)WBCheckBoxDlgProc,
                         (LPARAM)(LPCHECKBOXSTRUCT)&CheckBoxStruct );
                    
   WBRestoreTaskState();                    

   switch ( ret ) {
   
      case IDOK:
         lpResult->var.tLong = g_npTask->lOption;
         break;
         
      case IDCANCEL:
         lpResult->var.tLong = -1;
         break;
        
      default:
         assert( FALSE );
         return WBERR_INTERNAL_ERROR;
         
   }
   
   WBLocalFree( (HLOCAL)npszTemp2 );
   WBLocalFree( (HLOCAL)LOWORD(CheckBoxStruct.lpszPrompt));
   
   lpResult->type = VT_I4;
   return 0;
}
      
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
BOOL CALLBACK __export 
WBCheckBoxDlgProc( HWND hwndDlg, UINT msg, WPARAM wParam, LPARAM lParam )
{
   int i;
   DWORD dwBitMask;
   BOOL fSetCheck;
   
   WBRestoreTaskState();
   
   switch ( msg ) {
   
      case WM_INITDIALOG:
         {             
         LPCHECKBOXSTRUCT lpCheckBoxStruct;
         int nScreenWidth, nScreenHeight;
         int xPos, yPos;
         int nLogPixelsX, nLogPixelsY;
         RECT rcParent, rcDialog;
         HWND hwndButton;
         HDC hdc;
         
         lpCheckBoxStruct = (LPCHECKBOXSTRUCT)lParam;
         SetWindowText( hwndDlg, lpCheckBoxStruct->lpszTitle );
         SetDlgItemText( hwndDlg, IDC_EDIT, lpCheckBoxStruct->lpszPrompt );

         dwBitMask = 0x00000001L;         
         for ( i = 0; i < lpCheckBoxStruct->ciOption; i++ ) {
            hwndButton = GetDlgItem(hwndDlg, IDC_CHECK1 + i);
            SetWindowText( hwndButton, 
                           lpCheckBoxStruct->lpszOption[i] );
            if ( lpCheckBoxStruct->dwCheckBits & dwBitMask )
               fSetCheck = TRUE;
            else
               fSetCheck = FALSE;
            SendMessage( hwndButton,
                         BM_SETCHECK,
                         fSetCheck,
                         0L );
            dwBitMask <<= 1;                                
         }
         
         g_npTask->lOption = (long)lpCheckBoxStruct->dwCheckBits;
          
         // Disable and hide unused options
         for ( ; i < MAXOPTIONS; i++ ) {
            hwndButton = GetDlgItem( hwndDlg, IDC_CHECK1 + i );
            EnableWindow( hwndButton, FALSE );
            ShowWindow( hwndButton, SW_HIDE );
         }

         GetWindowRect( hwndDlg, &rcDialog );
         nScreenWidth = GetSystemMetrics( SM_CXSCREEN );
         nScreenHeight = GetSystemMetrics( SM_CYSCREEN );
         
         if ( lpCheckBoxStruct->xPos != 0 ) {         
            hdc = GetDC( hwndDlg );
            nLogPixelsX = GetDeviceCaps( hdc, LOGPIXELSX );
            nLogPixelsY = GetDeviceCaps( hdc, LOGPIXELSY );
            ReleaseDC( hwndDlg, hdc );
            xPos = (int)((long)lpCheckBoxStruct->xPos * 
                        (long)nLogPixelsX / 1440L);
            yPos = (int)((long)lpCheckBoxStruct->yPos * 
                        (long)nLogPixelsY / 1440L);
         }
         else {
            xPos = (nScreenWidth - (rcDialog.right - rcDialog.left)) / 2;
            yPos = (nScreenHeight - (rcDialog.bottom - rcDialog.top)) / 3;
         }
         
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

         SetFocus( GetDlgItem( hwndDlg, IDOK ));
         return FALSE;
         }
         
      case WM_COMMAND:
         switch ( wParam ) {
         
            case IDC_CHECK1:
            case IDC_CHECK2:
            case IDC_CHECK3:
            case IDC_CHECK4:
            case IDC_CHECK5:
            case IDC_CHECK6:
            case IDC_CHECK7: 
            case IDC_CHECK8: 
            case IDC_CHECK9: 
            case IDC_CHECK10:
               dwBitMask = 0x00000001L << (wParam - IDC_CHECK1);
               g_npTask->lOption ^= dwBitMask;
               return TRUE;           
               
            case IDOK:
               EndDialog( hwndDlg, IDOK );
               return TRUE;
               
            case IDCANCEL:
               EndDialog( hwndDlg, IDCANCEL );
               return TRUE;
               
         }
         break;
         
      case WM_CTLCOLOR:
         if (g_npTask->fLargeDialogs)
         {
            if ((HWND)LOWORD(lParam) == GetDlgItem(hwndDlg, IDC_EDIT)) {
               SetBkColor((HDC)wParam, RGB(192, 192, 192));
               return (BOOL)g_hbrPrompt;
            }
         }
         break;
   }
   
   return FALSE;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnMsgBox( int narg, LPVAR lparg, LPVAR lpResult )
{
   NPSTR npszText;
   LPSTR lpszTitle;
   LPSTR lpch;
   USHORT usLen;
   long lStyle;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 3 )
      return WBERR_ARGTOOMANY;

   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   usLen = WBGetHlstrLen( lparg[0].var.tString );
   
   // Truncate message to 1024 characters at most
   usLen = min( usLen, 1024 );
   
   npszText = (NPSTR)WBLocalAlloc( LPTR, usLen + 1 );
   if ( npszText == NULL )
      return WBERR_OUTOFMEMORY;
   
   memcpy( npszText, WBDerefHlstr( lparg[0].var.tString ), usLen);
   *(npszText+usLen) = '\0';

   // If length is greater than 255, truncate to 255 if first 255 characters
   // contains no space   
   if ( usLen > 255 ) {
      lpch = strchr( npszText, ' ' );
      if ( lpch == NULL || (lpch - (LPSTR)npszText > 255 ) ) 
         *(npszText+255) = '\0';
   }

   // Replace '|' by new line character in message text   
   lpch = strchr( npszText, '|' );
   while ( lpch != NULL ) {
      *lpch = '\n';
      lpch = strchr( lpch+1, '|' );
   }
   
   if ( narg >= 2 ) {
      if ( lparg[1].type == VT_EMPTY )
         lStyle = 0;
      else
         lStyle = lparg[1].var.tLong;
   }
   else
      lStyle = 0;

   if ( narg == 3 ) {
      lpszTitle = WBDerefZeroTermHlstr( lparg[2].var.tString );
   }
   else {
      lpszTitle = g_szAppTitle;
   }
   
   lpResult->var.tLong = (long)WBMessageBox( g_npTask->hwndClient, 
                                             (LPSTR)npszText, 
                                             lpszTitle, 
                                             (UINT)lStyle );
   WBLocalFree( (HLOCAL)npszText );
                                                
   lpResult->type = VT_I4;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnExitWindows( int narg, LPVAR lparg, LPVAR lpResult )
{
   BOOL fOk;
   WORD wExit;
   HWND hwndPM;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;
      
   if ( narg == 1 ) 
      wExit = (WORD)lparg[0].var.tLong;
   else 
      wExit = 0;

   WBBroadcastEvent( (UINT)wExit, 0, 0L );
   
   // Try and set the focus to Program Manager in order to avoid
   // a possible problem in Windows 3.1
   hwndPM = FindWindow(NULL, "Program Manager");
   if (hwndPM != NULL)
      SetFocus(hwndPM);
      
   fOk = ExitWindows( MAKELONG( wExit, 0 ), 0 );
   WBRestoreTaskState();
   
   lpResult->var.tLong = (long)( fOk ? WB_TRUE : WB_FALSE );
   lpResult->type = VT_I4;
   return 0;
}
         
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnExitWindowsExec( int narg, LPVAR lparg, LPVAR lpResult )
{
   BOOL fOk;
   LPSTR lpszExe;
   LPSTR lpszParms;
   HWND hwndPM;
   
   if ( narg < 2 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 2 )
      return WBERR_ARGTOOMANY;

   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   lpszExe = WBDerefZeroTermHlstr( lparg[0].var.tString );
   if ( *lpszExe == '\0' )
      return WBERR_ARGRANGE;

   lpszParms = WBDerefZeroTermHlstr( lparg[1].var.tString );
      
   WBBroadcastEvent( EW_RESTARTWINDOWS, 0, 0L );
   
   // Try and set the focus to Program Manager in order to avoid
   // a possible problem in Windows 3.1
   hwndPM = FindWindow(NULL, "Program Manager");
   if (hwndPM != NULL)
      SetFocus(hwndPM);
      
   fOk = ExitWindowsExec( lpszExe, lpszParms );
   WBRestoreTaskState();
   
   lpResult->var.tLong = (long)( fOk ? WB_TRUE : WB_FALSE );
   lpResult->type = VT_I4;
   return 0;
}
         
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnShowWaitCursor( int narg, LPVAR lparg, LPVAR lpResult )
{
   UNREFERENCED_PARAM( lparg );
   
   if ( narg != 0 )
      return WBERR_EXPECT_NOARG;
      
   WBShowWaitCursor();
   
   lpResult->var.tLong = WB_TRUE;
   lpResult->type = VT_I4;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnRestoreCursor( int narg, LPVAR lparg, LPVAR lpResult )
{
   UNREFERENCED_PARAM( lparg );
   
   if ( narg != 0 )
      return WBERR_EXPECT_NOARG;
      
   WBRestoreCursor();
   
   lpResult->var.tLong = WB_TRUE;
   lpResult->type = VT_I4;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/* Shell( commandstring [, [windowstyle], waitflag]] )                      */
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnShell( int narg, LPVAR lparg, LPVAR lpResult )
{
   VARIANT vTemp;
   LPSTR lpszCmdLine;
   UINT fuCmdShow;
   BOOL fWait;
   UINT ret;
   ERR err;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 3 )
      return WBERR_ARGTOOMANY;
  
   // arg 1: commandstring    
   if ( lparg[0].type == VT_EMPTY )
      return WBERR_ARGRANGE;
   lpszCmdLine = WBDerefZeroTermHlstr( lparg[0].var.tString );
   
   while ( isspace( *lpszCmdLine ) )
      lpszCmdLine++;
      
   if ( *lpszCmdLine == '\0' )
      return WBERR_ARGRANGE;

   // arg2: windowstyle      
   if ( narg >= 2 ) {
      switch ( lparg[1].type ) {
      
         case VT_EMPTY:
            fuCmdShow = SW_SHOWNORMAL;
            break;
            
         default:
            if ( (err = WBMakeLongVariant( &lparg[1] )) != 0 )
               return err;
            fuCmdShow = (UINT)lparg[1].var.tLong;
      }
   }
   else {
      fuCmdShow = SW_SHOWNORMAL;
   }
   
   if ( fuCmdShow == 0 )
      return WBERR_ARGRANGE;

   // arg3: waitflag   
   if ( narg == 3 ) {
      fWait = (BOOL)(lparg[2].var.tLong == 0 ? FALSE : TRUE );
   }
   else {
      fWait = FALSE;
   }

   WBShowWaitCursor();   
   
   ret = WinExec( lpszCmdLine, fuCmdShow );
   
   WBRestoreTaskState();
   WBRestoreCursor();
      
   if ( ret < 32 )
      return WBERR_WINEXEC;

   if ( fWait ) {
      vTemp.var.tLong = (long)ret;
      vTemp.type = VT_I4;
      return WBFnShellWait( 1, &vTemp, lpResult );
   }
   else {
      lpResult->var.tLong = (long)ret;
      lpResult->type = VT_I4;
      return 0;
   }
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnShellWait( int narg, LPVAR lparg, LPVAR lpResult )
{
   WNDCLASS  wc;
   MSG msg;
   HWND hwnd;
   BOOL fClassCreated;
   TASKENTRY te;
   BOOL fMore;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;
   
   g_npTask->ShellWaitInfo.hinstChild = (HINSTANCE)LOWORD(lparg[0].var.tLong);

   // Walk the task list and find task info corresponding to specified
   // instance handle.
   te.dwSize = sizeof( TASKENTRY );
   fMore = TaskFirst( &te );
    
   while ( fMore ) {
      if ( te.hInst == g_npTask->ShellWaitInfo.hinstChild )
         break;
      fMore = TaskNext( &te );
   }

   // If no corresponding task was found, either the instance handle is
   // invalid or the task has already ended. In this case, set exit code
   // to -1 and exit.
   if ( !fMore ) {
      g_npTask->ShellWaitInfo.hinstChild = NULL;
      lpResult->var.tLong = -1;
      lpResult->type = VT_I4;
      return 0;
   }
   
   // Save the module handle corresponding to the specified instance handle
   g_npTask->ShellWaitInfo.hmodChild = te.hModule;
    
   if ( GetClassInfo( g_hinstDLL, szWaitClass, &wc ) == 0 ) {
      wc.style         = 0;
      wc.lpfnWndProc   = (WNDPROC)WBShellWaitWndProc;
      wc.cbClsExtra    = 0;
      wc.cbWndExtra    = 0;
      wc.hInstance     = g_hinstDLL;
      wc.hIcon         = NULL;
      wc.hCursor       = NULL;
      wc.hbrBackground = NULL;
      wc.lpszMenuName  = NULL;
      wc.lpszClassName = szWaitClass;
   
      if ( RegisterClass( &wc ) == NULL )
         return WBERR_REGISTERCLASS;
         
      fClassCreated = TRUE;
   }
   else {
      fClassCreated = FALSE;
   }

   hwnd = CreateWindowEx( 0,
                          szWaitClass,
                          "",
                          WS_POPUP,
                          0,
                          0,
                          0,
                          0,
                          NULL,
                          NULL,
                          g_hinstDLL,
                          NULL );

   if ( hwnd == NULL ) {
      UnregisterClass( szWaitClass, g_hinstDLL );
      return WBERR_CREATEWINDOW;
   }
   
   g_npTask->ShellWaitInfo.hwndWait = hwnd;
   
   ShowWindow( hwnd, SW_HIDE );
   UpdateWindow( hwnd );
   
   EnableWindow( g_npTask->hwndClient, FALSE );
   
   while ( IsWindow( hwnd ) && GetMessage( &msg, NULL, NULL, NULL ) ) {
      DispatchMessage( &msg );
      WBRestoreTaskState();
   }
   
   EnableWindow( g_npTask->hwndClient, TRUE );
   SetWindowPos( g_npTask->hwndClient, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE|SWP_SHOWWINDOW );
   
   if ( fClassCreated )
      UnregisterClass( szWaitClass, g_hinstDLL );

   g_npTask->ShellWaitInfo.hinstChild = NULL;
   lpResult->var.tLong = g_npTask->ShellWaitInfo.lExitCode;
   lpResult->type = VT_I4;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
LRESULT CALLBACK __export 
WBShellWaitWndProc( HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam )
{
   WBRestoreTaskState();
   
   switch ( message ) {
   
      case WM_CREATE:
         NotifyRegister( NULL, 
                         (LPFNNOTIFYCALLBACK)WBNotifyCallback, 
                         NF_NORMAL );
         return 0;
         
      case WM_CLOSE:
         NotifyUnRegister( NULL );
         DestroyWindow( hwnd );
         return 0;
   }
   
   return DefWindowProc( hwnd, message, wParam, lParam );
}
         
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
BOOL CALLBACK __export WBNotifyCallback( WORD wID, DWORD dwData )
{
   HTASK hTask;     // handle of the task that called the notification call back
   NPWBTASK npt;
   TASKENTRY te;
   
   // Obtain info about the task that is terminating
   hTask = GetCurrentTask();
   te.dwSize = sizeof(TASKENTRY);
   TaskFindHandle(&te, hTask);
      
   // Check for task exiting
   switch ( wID ) {
    
      case NFY_EXITTASK:
            
         // Search our own task list for a host application that is 
         // the parent of the terminating task
         npt = WBFindTaskParent( te.hInst );
               
         if ( npt != NULL ) {
            // Task terminating is spawned by one of our host apps.
            // Save the exit code.
            npt->ShellWaitInfo.lExitCode = (long)(WORD)LOWORD(dwData);
            PostMessage( npt->ShellWaitInfo.hwndWait, 
                         WM_CLOSE,
                         0, 0L );
         }
         break;
         
    }

    // Pass notification to other callback functions
    return FALSE;

}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
NPWBTASK WBFindTaskParent( HINSTANCE hinst )
{
   NPWBTASK npt;
   
   npt = g_npTaskFirst;
   while ( npt != NULL ) {
      if ( npt->ShellWaitInfo.hinstChild == hinst )
         break;
      npt = npt->npNext;
   }
   
   return npt;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnFindWindow( int narg, LPVAR lparg, LPVAR lpResult )
{
   WNDENUMPROC wndenmprc;
   FINDWINDOWSTRUCT FindWindowStruct;
   BOOL fWindowExists;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;
      
   FindWindowStruct.lpszTitle = WBDerefZeroTermHlstr( lparg[0].var.tString );

   wndenmprc = (WNDENUMPROC)MakeProcInstance( (FARPROC)EnumWndProc, g_hinstDLL );
   fWindowExists = !EnumWindows( wndenmprc, (LPARAM)(LPVOID)&FindWindowStruct );
   FreeProcInstance( (FARPROC)wndenmprc );

   lpResult->var.tLong = (long)(fWindowExists ? (WORD)FindWindowStruct.hwnd : 0 );
   lpResult->type = VT_I4;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
BOOL CALLBACK __export EnumWndProc( HWND hwnd, LPARAM lParam )
{
   char szTitle[MAXTEXT];
   int cb;
   FINDWINDOWSTRUCT FAR* lpFindWindowStruct;
   
   lpFindWindowStruct = (FINDWINDOWSTRUCT FAR*)lParam;
   
   cb = GetWindowText( hwnd, szTitle, sizeof(szTitle)-1 );

   if ( cb != 0 ) {
      if ( _strnicmp( szTitle, 
                      lpFindWindowStruct->lpszTitle, 
                      strlen( lpFindWindowStruct->lpszTitle ) ) == 0 ) {
         lpFindWindowStruct->hwnd = hwnd;                   
         return FALSE;
      }
   }

   return TRUE;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnGetInstance( int narg, LPVAR lparg, LPVAR lpResult )
{
   HWND hwnd;
   HTASK htask;
   TASKENTRY te;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;

   hwnd = (HWND)(WORD)lparg[0].var.tLong;
   
   if ( IsWindow( hwnd ) ) {
      htask = GetWindowTask( hwnd );
      te.dwSize = sizeof(TASKENTRY);
      if ( TaskFindHandle( &te, htask ) ) {
         lpResult->var.tLong = (long)(WORD)te.hInst;
         lpResult->type = VT_I4;
         return 0;
      }
   }
   
   lpResult->var.tLong = 0;
   lpResult->type = VT_I4;
   return 0;
}

   