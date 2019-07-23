#define STRICT
#include <windows.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <time.h>
#include "wupbas.h"
#include "wbfunc.h"
#include "wbfnstr.h"

extern NPWBTASK g_npTask;
extern HINSTANCE g_hinstDLL;

static char *lpszDayName[] = {
   "Sunday",
   "Monday",
   "Tuesday",
   "Wednesday",
   "Thursday",
   "Friday",
   "Saturday"
};

static char *lpszMonthName[] = {
   "January",
   "February",
   "March",
   "April",
   "May",
   "June",
   "July",
   "August",
   "September",
   "October",
   "November",
   "December"
};

static char NEAR szWaitClass[] = "WupBas:Wakeup";

LRESULT CALLBACK __export 
WBWakeupWndProc( HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam );

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnNow( int narg, LPVAR lparg, LPVAR lpResult )
{
   UNREFERENCED_PARAM( lparg );
   
   if ( narg != 0 )
      return WBERR_EXPECT_NOARG;
   
   lpResult->var.tLong = (long)time( NULL );
   lpResult->type = VT_DATE;
   return 0;
}
     
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnDay( int narg, LPVAR lparg, LPVAR lpResult )
{
   ERR err;
   struct tm *ptm;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;
      
   switch ( lparg[0].type ) {
   
      case VT_DATE:
         break;
         
      case VT_STRING:
         if ( (err = WBVariantStringToOther( &lparg[0], VT_DATE )) != 0 )
            return err;
         break;
         
      default:
         return WBERR_ARGTYPEMISMATCH;
   }
   
   ptm = localtime( &((time_t)lparg[0].var.tLong) );
   
   lpResult->var.tLong = (long)ptm->tm_mday;
   lpResult->type = VT_I4;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnWeekDay( int narg, LPVAR lparg, LPVAR lpResult )
{
   ERR err;
   LPSTR lpsz;
   int nFormat;
   struct tm *ptm;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 2 )
      return WBERR_ARGTOOMANY;
      
   switch ( lparg[0].type ) {
   
      case VT_DATE:
         break;
         
      case VT_STRING:
         if ( (err = WBVariantStringToOther( &lparg[0], VT_DATE )) != 0 )
            return err;
         break;
         
      default:
         return WBERR_ARGTYPEMISMATCH;
   }
   
   if ( narg == 2 ) {
      if ( lparg[1].type != VT_STRING )
         return WBERR_ARGTYPEMISMATCH;
      lpsz = WBDerefZeroTermHlstr( lparg[1].var.tString );
      if ( stricmp( lpsz, "d" ) == 0 )
         nFormat = 0;
      else if ( stricmp( lpsz, "dd" ) == 0 )
         nFormat = 0;
      else if ( stricmp( lpsz, "ddd" ) == 0 )
         nFormat = 1;
      else if ( stricmp( lpsz, "dddd" ) == 0 )
         nFormat = 2;
      else
         return WBERR_ARGRANGE;
   }
   else {
      nFormat = 0;
   }
         
   ptm = localtime( &((time_t)lparg[0].var.tLong) );
   
   switch ( nFormat ) {
      
      case 0:
         lpResult->var.tLong = (long)ptm->tm_wday;
         lpResult->type = VT_I4;
         break;
         
      case 1:
         lpResult->var.tString = WBCreateTempHlstr( lpszDayName[ptm->tm_wday], 3 );
         if ( lpResult->var.tString == NULL )
            return WBERR_STRINGSPACE;
         lpResult->type = VT_STRING;
         break;

      case 2:                                                    
         lpResult->var.tString = WBCreateTempHlstr( lpszDayName[ptm->tm_wday],
                                    (USHORT)strlen(lpszDayName[ptm->tm_wday]) );
         if ( lpResult->var.tString == NULL )
            return WBERR_STRINGSPACE;
         lpResult->type = VT_STRING;
         break;
         
      default:
         return WBERR_INTERNAL_ERROR;
   }
   
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnMonth( int narg, LPVAR lparg, LPVAR lpResult )
{
   ERR err;
   LPSTR lpsz;
   int nFormat;
   struct tm *ptm;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 2 )
      return WBERR_ARGTOOMANY;
      
   switch ( lparg[0].type ) {
   
      case VT_DATE:
         break;
         
      case VT_STRING:
         if ( (err = WBVariantStringToOther( &lparg[0], VT_DATE )) != 0 )
            return err;
         break;
         
      default:
         return WBERR_ARGTYPEMISMATCH;
   }
   
   if ( narg == 2 ) {
      if ( lparg[1].type != VT_STRING )
         return WBERR_ARGTYPEMISMATCH;
      lpsz = WBDerefZeroTermHlstr( lparg[1].var.tString );
      if ( stricmp( lpsz, "m" ) == 0 )
         nFormat = 0;
      else if ( stricmp( lpsz, "mm" ) == 0 )
         nFormat = 0;
      else if ( stricmp( lpsz, "mmm" ) == 0 )
         nFormat = 1;
      else if ( stricmp( lpsz, "mmmm" ) == 0 )
         nFormat = 2;
      else
         return WBERR_ARGRANGE;
   }
   else {
      nFormat = 0;
   }
         
   ptm = localtime( &((time_t)lparg[0].var.tLong) );
   
   switch ( nFormat ) {
      
      case 0:
         lpResult->var.tLong = (long)ptm->tm_mon+1;
         lpResult->type = VT_I4;
         break;
         
      case 1:
         lpResult->var.tString = WBCreateTempHlstr( lpszMonthName[ptm->tm_mon], 3 );
         if ( lpResult->var.tString == NULL )
            return WBERR_STRINGSPACE;
         lpResult->type = VT_STRING;
         break;

      case 2:                                                    
         lpResult->var.tString = WBCreateTempHlstr( lpszMonthName[ptm->tm_mon],
                                    (USHORT)strlen(lpszMonthName[ptm->tm_mon]) );
         if ( lpResult->var.tString == NULL )
            return WBERR_STRINGSPACE;
         lpResult->type = VT_STRING;
         break;
         
      default:
         return WBERR_INTERNAL_ERROR;
   }
   
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnYear( int narg, LPVAR lparg, LPVAR lpResult )
{
   ERR err;
   LPSTR lpsz;
   int nFormat;
   struct tm *ptm;
   char szYear[5];
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 2 )
      return WBERR_ARGTOOMANY;
      
   switch ( lparg[0].type ) {
   
      case VT_DATE:
         break;
         
      case VT_STRING:
         if ( (err = WBVariantStringToOther( &lparg[0], VT_DATE )) != 0 )
            return err;
         break;
         
      default:
         return WBERR_ARGTYPEMISMATCH;
   }
   
   if ( narg == 2 ) {
      if ( lparg[1].type != VT_STRING )
         return WBERR_ARGTYPEMISMATCH;
      lpsz = WBDerefZeroTermHlstr( lparg[1].var.tString );
      if ( stricmp( lpsz, "yy" ) == 0 )
         nFormat = 0;
      else if ( stricmp( lpsz, "yyyy" ) == 0 )
         nFormat = 1;
      else
         return WBERR_ARGRANGE;
   }
   else {
      nFormat = 0;
   }
         
   ptm = localtime( &((time_t)lparg[0].var.tLong) );
   itoa( ptm->tm_year + 1900, szYear, 10 );
   
   switch ( nFormat ) {
      
      case 0:
         lpResult->var.tString = WBCreateTempHlstr( szYear+2 , 2);
         if ( lpResult->var.tString == NULL )
            return WBERR_STRINGSPACE;
         lpResult->type = VT_STRING;
         break;

      case 1:                                                    
         lpResult->var.tString = WBCreateTempHlstr( szYear, 4 );
         if ( lpResult->var.tString == NULL )
            return WBERR_STRINGSPACE;
         lpResult->type = VT_STRING;
         break;
         
      default:
         return WBERR_INTERNAL_ERROR;
   }
   
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnTimer( int narg, LPVAR lparg, LPVAR lpResult )
{
   UNREFERENCED_PARAM( lparg );
   
   if ( narg != 0 )
      return WBERR_EXPECT_NOARG;
      
   lpResult->var.tLong = (long)GetTickCount();
   lpResult->type = VT_I4;
   return 0;
}
     
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnWakeup( int narg, LPVAR lparg, LPVAR lpResult )
{
   WNDCLASS  wc;
   MSG msg;
   HWND hwnd;
   BOOL fClassCreated;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;
   
   // Attempt to convert argument to a long variant. If this fails
   // it still may be a string representing a valid hh:mm:ss time.
   WBMakeLongVariant( &lparg[0] );
   
   switch ( lparg[0].type ) {
   
      case VT_I4:
         if ( lparg[0].var.tLong < 0L )
            return WBERR_ARGRANGE;
         g_npTask->lWakeupTime = (long)time( NULL ) + lparg[0].var.tLong;
         break;

      case VT_STRING:
         {
            struct tm tim;
            LPSTR lpsz;
            int hh, mm, ss;
            int i;
            time_t now;
            
            hh = mm = ss = 0;
                  
            lpsz = WBDerefZeroTermHlstr( lparg[0].var.tString );
            while ( isspace( *lpsz ) )
               lpsz++;
               
            for ( i = 0; i < 2; i++ ) {
               if ( !isdigit( *lpsz ) )
                  break;
               hh = hh * 10 + (*lpsz++ - '0');
               
            }
            if ( *lpsz == ':' ) {
               lpsz++;
               for ( i = 0; i < 2; i++ ) {
                  if ( !isdigit( *lpsz ) )
                     break;
                  mm = mm * 10 + (*lpsz++ - '0');
               }
               if ( *lpsz == ':' ) {
                  lpsz++;
                  for ( i = 0; i < 2; i++ ) {
                     if ( !isdigit( *lpsz ) )
                        break;
                     ss = ss * 10 + (*lpsz++ - '0');
                  }
               }
            }
            
            if ( *lpsz != '\0' )
               return WBERR_ARGRANGE;
            
            if ( hh > 23 || mm > 59 || ss > 59 )
               return WBERR_ARGRANGE;

            time( &now );               
            memcpy( &tim, localtime( &now ), sizeof(struct tm) );
            tim.tm_hour = hh;
            tim.tm_min = mm;
            tim.tm_sec = ss;
            tim.tm_isdst = -1;
            g_npTask->lWakeupTime = (long)mktime( &tim );
            if ( now >= (time_t)g_npTask->lWakeupTime ) {
               lpResult->var.tLong = 0L;
               lpResult->type = VT_I4;
               return 0;
            }
         }
         break;
            
      default:
         return WBERR_ARGTYPEMISMATCH;      
   }
      
   if ( GetClassInfo( g_hinstDLL, szWaitClass, &wc ) == 0 ) {
      wc.style         = 0;
      wc.lpfnWndProc   = (WNDPROC)WBWakeupWndProc;
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
                          g_npTask->hwndClient,
                          NULL,
                          g_hinstDLL,
                          NULL );

   if ( hwnd == NULL ) {
      UnregisterClass( szWaitClass, g_hinstDLL );
      return WBERR_CREATEWINDOW;
   }
   
   ShowWindow( hwnd, SW_HIDE );
   EnableWindow( g_npTask->hwndClient, FALSE );
   
   while ( IsWindow( hwnd ) && GetMessage( &msg, NULL, NULL, NULL ) ) {
      DispatchMessage( &msg );
   }
   
   EnableWindow( g_npTask->hwndClient, TRUE );
   
   if ( fClassCreated )
      UnregisterClass( szWaitClass, g_hinstDLL );
      
   lpResult->var.tLong = 0L;
   lpResult->type = VT_I4;
   return 0;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
LRESULT CALLBACK __export 
WBWakeupWndProc( HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam )
{
   time_t now;
   
   WBRestoreTaskState();
   
   switch ( message ) {
   
      case WM_CREATE:
         if ( SetTimer( hwnd, 1, 1000, NULL ) == 0 ) 
            PostMessage( hwnd, WM_CLOSE, 0, 0L );
         return 0;
         
      case WM_TIMER:
         time( &now );
         if ( now >= (time_t)g_npTask->lWakeupTime ) 
            PostMessage( hwnd, WM_CLOSE, 0, 0L );
         return 0;
         
      case WM_CLOSE:
         DestroyWindow( hwnd );
         return 0;
   }
   
   return DefWindowProc( hwnd, message, wParam, lParam );
}
         

