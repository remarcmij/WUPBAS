#define STRICT
#include <windows.h>
#include <string.h>

#include "wupbas.h"

extern HCURSOR g_hcurWait;

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
int WBMessageBox( HWND hwndParent, LPCSTR lpszText, LPCSTR lpszTitle, UINT fuStyle )
{
   int nReturn;
   LPSTR lpszTemp;
   
   lpszTemp = WBReplaceMetaChars((LPSTR)lpszText, FALSE);
   if (lpszTemp == NULL)
      WBRuntimeError(WBERR_OUTOFMEMORY);

   nReturn = MessageBox( hwndParent, lpszTemp, lpszTitle, fuStyle );
   WBRestoreTaskState();
   
   WBLocalFree((HLOCAL)LOWORD(lpszTemp));
   
   return nReturn;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBShowWaitCursor( void )
{
   if ( g_npTask->nWaitCursorCount++ == 0 ) {
      SetCapture( g_npTask->hwndClient );
      g_npTask->hcurPrev = SetCursor( g_hcurWait );
   }
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBRestoreCursor( void )
{
   if ( g_npTask->nWaitCursorCount != 0 ) {
      if ( --g_npTask->nWaitCursorCount == 0 ) {
         SetCursor( g_npTask->hcurPrev );
         ReleaseCapture();
      }
   }
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
LPSTR WBReplaceMetaChars(LPSTR lpszText, BOOL fExpand)
{
   LPSTR lpszResult, lpch, lpch1, lpch2;
   UINT cbOrg, cbResult, cbSub;
   
   cbOrg = cbResult = strlen(lpszText) + 1;

   // If fExpand is True, the "|" char should be translated to \r\n.
   // If False translate to \n only.
   if (fExpand)
   {
      lpch = strchr(lpszText, '|');
      while (lpch != NULL) 
      {
         lpch = strchr(lpch+1, '|');
         cbResult++;
      }
   }
      
   lpszResult = (LPSTR)WBLocalAlloc(LPTR, cbResult);
   if (LOWORD(lpszResult) == NULL)
      return NULL;

   lpch = lpszText;   
   lpch1 = strchr(lpch, '|');
   lpch2 = lpszResult;
   
   while (lpch1 != NULL) 
   {
      cbSub = lpch1 - lpch;
      memcpy(lpch2, lpch, cbSub);
      lpch2 += cbSub;
      if (fExpand)
         *lpch2++ = '\r';
      *lpch2++ = '\n';
      lpch = lpch1+1;
      lpch1 = strchr(lpch, '|');
   }
   
   cbSub = cbOrg - (lpch - lpszText);
   memcpy(lpch2, lpch, cbSub);
   
   lpch = strchr(lpszResult, '~');
   while (lpch != NULL) 
   {
      *lpch = '\t';
      lpch = strchr(lpch+1, '~');
   }
   
   return lpszResult;
}
