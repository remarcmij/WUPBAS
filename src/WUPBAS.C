#define STRICT
#include <windows.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <time.h>
#include <stdarg.h>
#include "wupbas.h"
#include "wbmem.h"
#include "wbfunc.h"
#include "wbfileio.h"
#include "wbdllcal.h"
#include "wupbas.rcv"

NPWBTASK g_npTaskFirst;
NPWBTASK g_npTask;
NPWBCONTEXT g_npContext;

HINSTANCE g_hinstDLL;
WORD g_wDataSeg;
HCURSOR g_hcurWait;
char g_szAppTitle[] = "Winup Basic";
char g_szDLLPath[_MAX_PATH];

int g_cClients = 0;
HBRUSH g_hbrPrompt;

static HTASK htaskLast = NULL;
static char szDefDatimFmt[] = "dd-mmm-yy hh:mm:ss";
static char szUndefinedMsg[] = "Undefined error";

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
int FAR PASCAL
LibMain( HANDLE hModule, WORD wDataSeg, WORD cbHeapSize, LPSTR lpszCmdLine)
{
   LPSTR lpch;

   UNREFERENCED_PARAM( lpszCmdLine );

   if ( cbHeapSize != 0 )
      UnlockData( 0 );

   g_hinstDLL = hModule;
   g_wDataSeg = wDataSeg;

   g_hcurWait = LoadCursor( NULL, IDC_WAIT );

   WBRegisterModel_SListObject();
   WBRegisterModel_FileObject();

   GetModuleFileName( hModule, g_szDLLPath, sizeof( g_szDLLPath ) );
   lpch = strrchr( g_szDLLPath, '\\' );
   if ( lpch != NULL )
      *(lpch+1) = '\0';

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
int WINAPI __export WBInitialize( HWND hwnd )
{
   VARIANT vTemp;
   ERR err;
   HDC hdc;

#ifdef DEBUGOUTPUT
   {
   char szMessage[64];

   wsprintf( szMessage, "t WUBBAS: WBInitialize called by task %X\r\n",
      GetCurrentTask() );
   OutputDebugString( szMessage );
   }
#endif

   if ( !IsWindow( hwnd ) ) {
      MessageBeep( MB_ICONSTOP );
      MessageBox( NULL,
                  "WBInitialize: Invalid HWND parameter",
                  g_szAppTitle,
                  MB_OK | MB_ICONHAND | MB_TASKMODAL );
      return WBERR_INITIALIZATION;
   }

   g_npTask = (NPWBTASK)LocalAlloc( LPTR, sizeof( WBTASK ) );

   if ( g_npTask == NULL ) {

      MessageBeep( MB_ICONSTOP );
      WBMessageBox( hwnd,
                    "Not enough memory to set up task database.",
                    g_szAppTitle,
                    MB_OK | MB_ICONSTOP );

      return WBERR_INITIALIZATION;
   }

   // Save task handle and main window handle of client app
   g_npTask->htaskClient  = GetCurrentTask();
   g_npTask->hwndClient   = hwnd;

   // Initialize the suballocation memory pools
   g_npTask->hStrDescPool = WBSubAllocInit();
   g_npTask->hStrPool     = WBSubAllocInit();
   g_npTask->hTypePool    = WBSubAllocInit();
   g_npTask->hLibPool     = WBSubAllocInit();

   g_npTask->npvtGlobal = (NPVARDEF NEAR*)WBLocalAlloc( LPTR,
                                    sizeof(VARDEF) * VAR_TABLE_SIZE );
   if ( g_npTask->npvtGlobal == NULL )
      return WBERR_INITIALIZATION;


   // Add internal functions to the declarations hash table
   WBInitFunctionTable();
                                       
   // insert task descriptor at top of task list
   g_npTask->npNext = g_npTaskFirst;
   g_npTaskFirst = g_npTask;

   // save task handle of current task, so we can skip unnessary
   // task context switching
   htaskLast = g_npTask->htaskClient;

   // Initialize Compare Text option to TRUE
   g_npTask->fCompareText = TRUE;

   // Initialize the date/time format string to its default
   strcpy( g_npTask->szDatimFmt, szDefDatimFmt );

   // Extract date/time format characteristics
   vTemp.type = VT_DATE;
   vTemp.var.tLong = (long)time( NULL );
   err = WBVariantDateToString( &vTemp );
   if ( err == 0 )
      WBDestroyHlstrIfTemp( vTemp.var.tString );


   if (g_cClients++ == 0 ) {
      hdc = GetDC(NULL);
      g_hbrPrompt = CreateSolidBrush(GetNearestColor(hdc, RGB(192, 192, 192)));
      ReleaseDC(NULL, hdc);
      
   }
   
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
UINT WINAPI __export WBAboutString( LPSTR lpszBuffer, UINT cb )
{
   char szBuffer[MAXTEXT];
   size_t size;
               
   wsprintf( szBuffer, "%s Version %s", VER_FILEDESCRIPTION_STR,
      VER_FILEVERSION_STR );
      
   size = min( strlen( szBuffer ), (size_t)cb-1 );
   memcpy( lpszBuffer, szBuffer, size );
   *(lpszBuffer+size) = '\0';

   return (UINT)size;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
int WINAPI __export WBRegisterLineCallBacks(
   LPFNNEXTLINE lpNextLine,
   LPFNGOTOLINE lpGotoLine )
{
   NPWBCONTEXT npContext;
   NPVARDEF NEAR* npvtLocal;
   ERR err;

   if ( (err = WBRestoreTaskState()) != 0 )
      return err;

   if ( lpNextLine == NULL || lpGotoLine == NULL ) {
      MessageBeep( MB_ICONSTOP );
      MessageBox( NULL,
                  "Null pointer(s) passed to WBRegisterLineCallBacks",
                  g_szAppTitle,
                  MB_OK | MB_ICONHAND | MB_TASKMODAL );
      return WBERR_INITIALIZATION;
   }

   npContext = WBPushContext();
   npvtLocal  = (NPVARDEF NEAR*)WBLocalAlloc( LPTR,
                                    sizeof(VARDEF) * VAR_TABLE_SIZE );
   if ( npContext == NULL || npvtLocal == NULL ) {
      MessageBeep( MB_ICONSTOP );
      WBMessageBox( g_npTask->hwndClient,
                    "Not enough memory to set up execution context.",
                    g_szAppTitle,
                    MB_OK | MB_ICONSTOP );

      return WBERR_INITIALIZATION;
   }

   npContext->npvtLocal = npvtLocal;
   npContext->lpfnNextLine = lpNextLine;
   npContext->lpfnGotoLine = lpGotoLine;

   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
int WINAPI __export WBRegisterCmdCallback( LPFNEXTCMD lpExtCmd )
{
   UINT err;

#ifdef DEBUGOUTPUT
   {
      char szMessage[64];
      wsprintf( szMessage, "t WUBBAS: WBRegisterCmdCallBack called by task %X\r\n",
         GetCurrentTask() );
      OutputDebugString( szMessage );
   }
#endif

   if ( (err = WBRestoreTaskState()) != 0 )
      return err;

   g_npTask->lpfnExtCmd = lpExtCmd;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
int WINAPI __export WBRegisterBroadcastHandler( LPFNBROADCASTHANDLER lpHandler )
{
   UINT err;

   if ( (err = WBRestoreTaskState()) != 0 )
      return err;

   g_npTask->lpfnBroadcastHandler = lpHandler;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
int WINAPI __export WBRun( void )
{
   int nRet;

#ifdef DEBUGOUTPUT
   {
      char szMessage[64];
      wsprintf( szMessage, "t WUBBAS: WBRun called by task %X\r\n",
         GetCurrentTask() );
      OutputDebugString( szMessage );
   }
#endif

   nRet = 0;

   if ( (nRet = WBRestoreTaskState()) != 0 )
      return nRet;

   if ( (nRet = Catch( (int FAR*)(g_npTask->CatchBuf) )) != 0 ) {
      NULL;
   }
   else {

      if ( g_npTask->npContext->lpfnNextLine == NULL ||
           g_npTask->npContext->lpfnGotoLine == NULL ) 
         WBRuntimeError( WBERR_NOLINECALLBACKS );

      // Force reading of initial line
      g_npTask->npContext->fEol = TRUE;
      NextToken();

      while ( LastToken() != TOKEN_EOF ) {
         WBStatement();
      }
   }

   WBDestroyOrphanTempHlstr();

   // All error codes > 0 and below or equal WBERR_BEGIN are decremented by one.
   // The STOP stament produces error code 3, which is returned as 2.
   // The EXIT stament produces error code 2, which is returned as 1.
   // The END statement produces error code 1, which becomes error code 0,
   // meaning: no error.
   if ( nRet > 0 && nRet <= WBERR_BEGIN )
      nRet--;

   return (UINT)nRet;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
int WINAPI __export WBRunFile( LPSTR lpszFileName )
{
   VARIANT vResult;
   NPWBHASHENTRY npHash;
   UINT fuDeclType;
   ERR err;
   int nRet;

#ifdef DEBUGOUTPUT
   {
      char szMessage[64];
      wsprintf( szMessage, "t WUBBAS: WBRunFile called by task %X\r\n",
         GetCurrentTask() );
      OutputDebugString( szMessage );
   }
#endif

   nRet = 0;

   if ( (nRet = WBRestoreTaskState()) != 0 )
      return nRet;


   if ( (nRet = Catch( (int FAR*)(g_npTask->CatchBuf) )) != 0 ) {
      NULL;
   }
   else {
      if ( (err = WBLoadProgram( lpszFileName )) != 0 )
         WBRuntimeError( err );

      npHash.VoidPtr = WBLookupDeclaration( "MAIN", &fuDeclType );
      if ( fuDeclType != DECL_USERFUNC )
         WBRuntimeError( WBERR_NO_SUBMAIN );

      if ( (err = WBCallUserFunction( npHash.UserFunc, 0, NULL, &vResult )) > WBERR_EXIT )
         WBRuntimeError( err );
   }

   WBDestroyOrphanTempHlstr();

   // All error codes > 0 and below or equal WBERR_BEGIN are decremented by one.
   // The STOP stament produces error code 3, which is returned as 2.
   // The EXIT stament produces error code 2, which is returned as 1.
   // The END statement produces error code 1, which becomes error code 0,
   // meaning: no error.
   if ( nRet > 0 && nRet <= WBERR_BEGIN )
      nRet--;

   return (UINT)nRet;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
int WINAPI __export WBDeInitialize( void )
{
   ERR err;
   NPWBTASK npt, nptPrev;
   NPWBCONTEXT npcontext, npcontemp;

#ifdef DEBUGOUTPUT
   {
      char szMessage[64];
      wsprintf( szMessage, "t WUBBAS: WBDeInitialize called by task %X\r\n",
      GetCurrentTask() );
      OutputDebugString( szMessage );
   }
#endif

   if ( (err = WBRestoreTaskState()) != 0 )
      return err;

   // Unload all explicitly loaded DLL's
   WBUnloadExplicitLinks();

   // Destroy execution contexts
   npcontext = g_npTask->npContext;
   while ( npcontext != NULL ) {
      if ( npcontext->npvtLocal != NULL ) {
         // Destroy local variable table
         WBDestroyVarTable( npcontext->npvtLocal );
         WBLocalFree( (HLOCAL)npcontext->npvtLocal );
      }
      npcontemp = npcontext->npNext;
      WBLocalFree( (HLOCAL)npcontext );
      npcontext = npcontemp;
   }

#ifdef _DEBUG
   if ( g_npTask->lpLibHead != NULL )
      WBDestroyLib( g_npTask->lpLibHead );
#endif

   // Destroy declaration table
   WBDestroyDeclarationTable();
      
   // Destroy global variable table
   WBDestroyVarTable( g_npTask->npvtGlobal );
   WBLocalFree( (HLOCAL)g_npTask->npvtGlobal );
   
   // Destroy list of loaded program text files
   WBDestroyLoadLibList();
   
#ifdef DEBUGOUTPUT
   // check for local heap memory leaks (except default local heap)
   WBSubCheckMemoryLeaks();
#endif

   // Free local heaps
   WBDestroyLocalHeap( g_npTask->hStrPool );
   WBDestroyLocalHeap( g_npTask->hStrDescPool );
   WBDestroyLocalHeap( g_npTask->hTypePool );
   WBDestroyLocalHeap( g_npTask->hLibPool );

#ifdef DEBUGOUTPUT
   // check for memory leaks in default local heap
   WBLocalCheckMemoryLeaks();
#endif

   // search task list for current task
   npt = g_npTaskFirst;
   nptPrev = NULL;
   while ( npt != NULL ) {
      if ( npt == g_npTask )
         break;
      nptPrev = npt;
      npt = npt->npNext;
   }
   assert( npt != NULL );

   // remove task descriptor from task list
   if ( nptPrev == NULL ) {
      // current task is first in list
      g_npTaskFirst = g_npTask->npNext;
   }
   else {
      // current task is elsewhere in list
      nptPrev->npNext = g_npTask->npNext;
   }

   // Free task descriptor
   LocalFree( (HLOCAL)g_npTask );

   // Invalidate the saved task handle, since it is no longer valid
   htaskLast = NULL;

   if (--g_cClients == 0) 
      DeleteObject(g_hbrPrompt);
      
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
BOOL WINAPI __export WBSetLongVariable( LPCSTR lpszVarName, long lValue )
{
   NPVARDEF npvar;
   char szVarName[MAXTOKEN];
   VARIANT v;
   size_t len;

   if ( WBRestoreTaskState() != 0 )
      return FALSE;

   len = min( strlen( lpszVarName ), MAXTOKEN-1 );
   _fmemcpy( szVarName, lpszVarName, len );
   szVarName[len] = '\0';
   strupr( szVarName );

   v.var.tLong = lValue;
   v.type = VT_I4;

   npvar = WBInsertGlobalVariable( szVarName );
   WBSetVariable( npvar, &v );

   return TRUE;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
BOOL WINAPI __export WBSetStringVariable( LPCSTR lpszVarName, LPSTR lpszValue )
{
   NPVARDEF npvar;
   char szVarName[MAXTOKEN];
   VARIANT v;
   HLSTR hlstr;
   size_t len;

   if ( WBRestoreTaskState() != 0 )
      return FALSE;

   len = min( strlen( lpszVarName ), MAXTOKEN-1 );
   _fmemcpy( szVarName, lpszVarName, len );
   szVarName[len] = '\0';
   strupr( szVarName );

   hlstr = WBCreateHlstr( lpszValue, (USHORT)strlen( lpszValue ));
   if ( hlstr == NULL )
      return FALSE;

   v.var.tString = hlstr;
   v.type = VT_STRING;

   npvar = WBInsertGlobalVariable( szVarName );
   WBSetVariable( npvar, &v );

   return TRUE;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBRuntimeError( ERR err )
{
   // Size of szMessage: program lines have a max length of MAXTEXT.
   // Additional message text must not exceed the extra bytes specified
   // below.
   char szMessage[MAXTEXT+256];
   LPSTR lpch, lpch2;
   int len1, len2;

   if ( err == 0 )
      assert( err != 0 );

   if ( g_npTask->nWaitCursorCount != 0 ) {
      SetCursor( g_npTask->hcurPrev );
      ReleaseCapture();
      g_npTask->nWaitCursorCount = 0;
   }

   if ( err > WBERR_EXIT ) {

      if ( err == g_npTask->nCustomError ) {
         len1 = min( strlen( g_npTask->lpszCustomErrorText ), 80 );
         memcpy( szMessage, g_npTask->lpszCustomErrorText, len1 );
         szMessage[len1] = '\0';
      }
      else {
         if ( LoadString( g_hinstDLL, err, szMessage, sizeof( szMessage ) ) == 0 )
            strcpy( szMessage, szUndefinedMsg );
      }
     
      lpch = szMessage + strlen( szMessage ) - 1;
      if ( !ispunct( *lpch ) )
         strcat( szMessage, "." );
      strcat( szMessage, "\n" );

      if ( g_npContext != NULL ) {
      
         len1 = (size_t)strlen( szMessage );
   
         lpch = g_npContext->szProgText;
         while ( isspace( *lpch ) )
            lpch++;
   
         len2 = (LPSTR)g_npContext->npchPrevPos - lpch;
         if ( len2 > 0 ) {
            memcpy( szMessage + len1, lpch, (size_t)len2 );
            len1 += len2;
         }
         
         lpch = (LPSTR)g_npContext->npchPrevPos;
         while ( isspace( *lpch ) )
            lpch++;
            
         memcpy( szMessage + len1, "  ??  ", 6 );
         len1 += 6;
         
         // Copy at most 30 character after error position
         if ( strlen( lpch ) > 30 ) {
            lpch2 = lpch + 30;
            // Try and not break a word or number
            while ( lpch2 >= lpch && isalnum( *lpch2 ) )
               lpch2--;
            while ( lpch2 >= lpch && isspace( *lpch2 ) )
               lpch2--;
            lpch2++;
            if ( lpch2 == lpch ) 
               lpch2 = lpch + 30;
            len2 = lpch2 - lpch;
            memcpy( szMessage + len1, lpch, len2 );
            len1 += len2;
            strcpy( szMessage + len1, "..." );
         }
         else {
            strcpy( szMessage + len1, lpch );
         }
         
         lpch = (LPSTR)g_npContext->npszLibName;
         if ( lpch && *lpch ) {
            lpch2 = szMessage + strlen( szMessage );
            wsprintf( lpch2, "\n\n%s, line %u", lpch,
                      g_npContext->uLineNum );
         }
      }                                      

      MessageBeep( MB_ICONEXCLAMATION );
      MessageBox( g_npTask->hwndClient,
                  szMessage,
                  "Winup Basic Error",
                  MB_ICONEXCLAMATION | MB_OK );
   }

   Throw( (int FAR*)(g_npTask->CatchBuf), err );
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WINAPI __export WBSetErrorMessage( ERR err, LPSTR lpszErrorText )
{
   g_npTask->nCustomError = err;
   g_npTask->lpszCustomErrorText = lpszErrorText;

   return err;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
HWND WINAPI __export WBGetHostWindow( void )
{
   return g_npTask->hwndClient;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBStopExecution( ERR err )
{
   char szMessage[MAXTEXT];

   if ( err > WBERR_EXIT ) {

      if ( LoadString( g_hinstDLL, err, szMessage, sizeof( szMessage ) ) == 0 )
         strcpy( szMessage, szUndefinedMsg );

      MessageBeep( MB_ICONEXCLAMATION );
      WBMessageBox( g_npTask->hwndClient,
                  szMessage,
                  g_szAppTitle,
                  MB_ICONEXCLAMATION | MB_OK );
   }

   Throw( (int FAR*)(g_npTask->CatchBuf), err+1 );
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBRestoreTaskState( void )
{
   NPWBTASK npt;
   HTASK htaskClient;

   // Get handle of requesting task
   htaskClient = GetCurrentTask();

   // If requesting task is same as previous task, then the task descriptor
   // is already valid
   if ( htaskClient == htaskLast )
      return 0;

   // search task list for the current caller's task handle
   npt = g_npTaskFirst;
   while ( npt != NULL ) {
      if ( npt->htaskClient == htaskClient )
         break;
      npt = npt->npNext;
   }

   // return error if task not found
   if ( npt == NULL )
      return WBERR_TASK_NOT_REGISTERED;

   // save the task handle of the current task
   htaskLast = htaskClient;

   // save pointer to current task descriptor
   g_npTask = npt;
   g_npContext = npt->npContext;

   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
NPWBCONTEXT WBPushContext( void )
{
   NPWBCONTEXT npContext;
   
   npContext = (NPWBCONTEXT)WBLocalAlloc( LPTR, sizeof(WBCONTEXT) );
   if ( npContext != NULL ) {
      // Insert new context at top of context list and make it the
      // current context
      npContext->npNext = g_npTask->npContext;
      g_npTask->npContext = g_npContext = npContext;
   }

   return npContext;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBPopContext( void )
{
   NPWBCONTEXT npContext;

   npContext = g_npContext;
   g_npTask->npContext = g_npContext = npContext->npNext;
   
   WBLocalFree( (HLOCAL)npContext );
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
LPVOID WBNextLine( LPSTR lpszBuffer )
{
   LPVOID lpLine;

   if ( g_npContext->lpfnNextLine != NULL ) {
      lpLine = g_npContext->lpfnNextLine( lpszBuffer, MAXTEXT );
      WBRestoreTaskState();
   }
   else {
      lpLine = NULL;
      assert( FALSE );
   }

   return lpLine;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBGotoLine( LPVOID lpLine )
{
   if ( g_npContext->lpfnGotoLine != NULL ) {
      g_npContext->lpfnGotoLine( lpLine );
      WBRestoreTaskState();
   }
   else {
      assert( FALSE );
   }
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
int WBExtCmd( LPSTR lpszCommand )
{
   int nErr;
   if ( g_npTask->lpfnExtCmd == NULL )
      WBRuntimeError( WBERR_SYNTAX );

   nErr = g_npTask->lpfnExtCmd( lpszCommand );
   WBRestoreTaskState();

   return nErr;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBBroadcastEvent( UINT uEvent, WORD wParam, LONG lParam )
{
   if ( g_npTask->lpfnBroadcastHandler != NULL ) {
      g_npTask->lpfnBroadcastHandler( uEvent, wParam, lParam );
      WBRestoreTaskState();
   }
}
