#define STRICT
#include <windows.h>
#include <string.h>
#include "wupbas.h"
#include "wbmem.h"

extern WORD g_wDataSeg;

#define GUARD_BYTE   (BYTE)'\xCC'

typedef struct tagLSTRING {
   HMEM   hMem;
   USHORT usSize;
} LSTRING;
typedef LSTRING FAR* LPLSTRING;

BOOL WBCreateString( LPLSTRING lps, LPVOID pb, USHORT cbLen );

#ifdef _DEBUG
void WBAddGuardByte( LPLSTRING lps );
BOOL WBCheckGuardByte( LPLSTRING lps );
#endif

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
HLSTR WINAPI FAR __export WBCreateHlstr( LPVOID pb, USHORT cbLen )
{
   LPLSTRING lps;

   // String descriptors for non-tempory HLSTR's are allocated from
   // the string descriptor local heap pool
   lps = (LPLSTRING)WBSubAlloc( g_npTask->hStrDescPool,
                                LPTR,
                                sizeof( LSTRING ));
   if ( lps == NULL )
      return NULL;

   if ( !WBCreateString( lps, pb, cbLen ) ) {
      WBSubFree( (HMEM)lps );
      return NULL;
   }

   return (HLSTR)lps;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
HLSTR WINAPI FAR __export WBCreateHlstrFromTemp( HLSTR hlstr )
{
   LPLSTRING lps1, lps2;
   LPSTR lpsz;
   USHORT usSize;
   BOOL fOk;

   // String descriptors for non-tempory HLSTR's are allocated from
   // the string descriptor local heap pool
   lps1 = (LPLSTRING)WBSubAlloc( g_npTask->hStrDescPool,
                                 LPTR,
                                 sizeof( LSTRING ));
   if ( lps1 == NULL )
      return NULL;

   lps2 = (LPLSTRING)hlstr;

   // If the HLSTR argument is a temp string, assign the string data
   // to the new temp HLSTR, and set the old HLSTR to the NULL string.
   // In other words, the old HLSTR remains a valid HLSTR, but becomes
   // of zero-length. By doing we can avoid unnessary copying of string
   // string data.
   if ( HIWORD(lps2) == g_wDataSeg ) {
      *lps1 = *lps2;
      lps2->hMem = NULL;
      lps2->usSize = 0;
   }

   // If the HLSTR argument is not a temp string, the string data needs
   // to be copied.
   else {
      lpsz = WBLockHlstrLen( hlstr, &usSize );
      fOk = WBCreateString( lps1, lpsz, usSize );
      WBUnlockHlstr( hlstr );
      if ( !fOk ) {
         WBSubFree( (HMEM)lps1 );
         return NULL;
      }
   }

   return (HLSTR)lps1;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
HLSTR WINAPI FAR __export WBCreateTempHlstr( LPVOID pb, USHORT cbLen )
{
   LPLSTRING lps;
   int i;

   // Search for a free slot in the temporary HLSTR table
   for ( i = 0; i < MAXTEMPHLSTR; i++ ) {
      if ( g_npTask->TempHlstrTable[i] == NULL )
         break;
   }

   // Return zero if no available slot
   if ( i == MAXTEMPHLSTR )
      return NULL;

   // String descriptors for tempory HLSTR's are allocated from
   // the local heap in the default data segment
   lps = (LPLSTRING)WBLocalAlloc( LPTR, sizeof( LSTRING ));
   if ( lps == NULL )
      return NULL;

   if ( !WBCreateString( lps, pb, cbLen ) ) {
      WBLocalFree( (HLOCAL)LOWORD(lps) );
      return NULL;
   }

   g_npTask->TempHlstrTable[i] = (HLSTR)lps;
   return (HLSTR)lps;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
HLSTR WINAPI FAR __export WBCreateTempHlstrFromTemp( HLSTR hlstr )
{
   LPLSTRING lps1, lps2;
   LPSTR lpsz;
   USHORT usSize;
   BOOL fOk;
   int i;

   // Search for a free slot in the temporary HLSTR table
   for ( i = 0; i < MAXTEMPHLSTR; i++ ) {
      if ( g_npTask->TempHlstrTable[i] == NULL )
         break;
   }

   // Return zero if no available slot
   if ( i == MAXTEMPHLSTR )
      return NULL;

   // String descriptors for tempory HLSTR's are allocated from
   // the local heap in the default data segment
   lps1 = (LPLSTRING)WBLocalAlloc( LPTR, sizeof( LSTRING ));
   if ( lps1 == NULL )
      return NULL;

   lps2 = (LPLSTRING)hlstr;

   // If the HLSTR argument is a temp string, assign the string data
   // to the new temp HLSTR, and set the old HLSTR to the NULL string.
   // In other words, the old HLSTR remains a valid HLSTR, but becomes
   // of zero-length. By doing we can avoid unnessary copying of string
   // string data.
   if ( HIWORD(lps2) == g_wDataSeg ) {
      *lps1 = *lps2;
      lps2->hMem = NULL;
      lps2->usSize = 0;
   }

   // If the HLSTR argument is not a temp string, the string data needs
   // to be copied.
   else {

      lpsz = WBLockHlstrLen( hlstr, &usSize );
      fOk = WBCreateString( lps1, lpsz, usSize );
      WBUnlockHlstr( hlstr );
      if ( !fOk ) {
         WBLocalFree( (HLOCAL)LOWORD(lps1) );
         return NULL;
      }
   }

   g_npTask->TempHlstrTable[i] = (HLSTR)lps1;
   return (HLSTR)lps1;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
BOOL WBCreateString( LPLSTRING lps, LPVOID pb, USHORT cbLen )
{
   HMEM hmem;
   LPVOID lpv;
   USHORT cb;

   assert( lps != NULL );

   if ( cbLen > 0 ) {
      // Request an extra byte in case we need to zero-terminate the
      // string at some point in time (something calling
      // WBDerefZeroTermHlstr()
      cb = cbLen + 1;

#ifdef _DEBUG
      // Request an additional byte in a debug compile to store a
      // guard byte (check against writing out of bounds)
      cb++;
#endif

      hmem = WBSubAlloc( g_npTask->hStrPool, LMEM_MOVEABLE, cb );
      if ( hmem == NULL )
         return FALSE;

      lps->hMem = hmem;
      lps->usSize = cbLen;

#ifdef _DEBUG
      WBAddGuardByte( lps );
#endif

      if ( pb != NULL ) {
         lpv = WBSubLock( hmem );
         memcpy( lpv, pb, (size_t)cbLen );
         WBSubUnlock( hmem );
      }
   }
   else {
      lps->hMem = NULL;
      lps->usSize = 0;
   }

   return TRUE;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WINAPI FAR __export WBResizeHlstr( HLSTR hlstr, USHORT newCbLen )
{
   LPLSTRING lps;
   HMEM hmem;
   USHORT cb;

   lps = (LPLSTRING)hlstr;
   assert( lps != NULL );

   // Request an extra byte in case we need to zero-terminate the
   // string at some point in time (something calling
   // WBDerefZeroTermHlstr()
   cb = newCbLen + 1;
#ifdef _DEBUG
   // Request an additional byte in a debug compile to store a
   // guard byte (check against writing out of bounds)
   cb++;
#endif

   if ( lps->usSize != newCbLen ) {
      if ( newCbLen > 0 ) {
         if ( lps->usSize == 0 ) {
            assert( lps->hMem == NULL );
            hmem = WBSubAlloc( g_npTask->hStrPool,
                               LMEM_MOVEABLE,
                               cb );
         }
         else {
            assert( lps->hMem != NULL );
            assert( WBCheckGuardByte( lps ) );
            if ( newCbLen > lps->usSize ) {
               // If growing, resize memory
               hmem = WBSubReAlloc( lps->hMem, cb, LMEM_MOVEABLE );
            }
            else {
               // If shrinking, keep same memory block
               hmem = lps->hMem;
            }
         }

         if ( hmem == NULL )
            return WBERR_STRINGSPACE;

         lps->usSize = newCbLen;
         lps->hMem = hmem;

#ifdef _DEBUG
         WBAddGuardByte( lps );
#endif

      }
      else {
         if ( lps->usSize == 0 ) {
            assert( lps->hMem == NULL );
         }
         else {
            assert( lps->hMem != NULL );
            assert( WBCheckGuardByte( lps ) );
            WBSubFree( lps->hMem );
            lps->usSize = 0;
            lps->hMem = NULL;
         }
      }
   }

   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WINAPI FAR __export WBDestroyHlstr( HLSTR hlstr )
{
   LPLSTRING lps;

   if ( !WBDestroyHlstrIfTemp( hlstr ) ) {

      lps = (LPLSTRING)hlstr;
      assert( lps != NULL );

      if ( lps->usSize > 0 ) {
         assert( lps->hMem != NULL );
         assert( WBCheckGuardByte( lps ) );
         WBSubFree( lps->hMem );
      }
      else {
         assert( lps->hMem == NULL );
      }

      WBSubFree( (HMEM)hlstr );
   }
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
USHORT WINAPI FAR __export WBGetHlstrLen( HLSTR hlstr )
{
   LPLSTRING lps;
   USHORT usSize;

   lps = (LPLSTRING)hlstr;
   assert( lps != NULL );

   if ( lps->usSize == 0 ) {
      assert( lps->hMem == NULL );
      usSize = 0;
   }
   else {
      assert( lps->hMem != NULL );
      assert( WBCheckGuardByte( lps ) );
      usSize = lps->usSize;
   }

   return usSize;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
USHORT WINAPI FAR __export WBGetHlstr( HLSTR hlstr, LPVOID pb, USHORT cbLen )
{
   LPLSTRING lps;
   USHORT usSize;
   LPVOID lpv;

   lps = (LPLSTRING)hlstr;
   assert( lps != NULL );

   if ( lps->usSize == 0 ) {
      assert( lps->hMem == NULL );
      usSize = 0;
   }
   else {
      assert( lps->hMem != NULL );
      assert( WBCheckGuardByte( lps ) );
      lpv = WBSubLock( lps->hMem );
      if ( lpv == NULL ) {
         assert( FALSE );
         return 0;
      }
      usSize = min( lps->usSize, cbLen );
      if ( usSize > 0 )
         memcpy( pb, lpv, (size_t)usSize );

      WBSubUnlock( lps->hMem );
   }

   WBDestroyHlstrIfTemp( hlstr );

   return usSize;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WINAPI FAR __export WBSetHlstr( HLSTR FAR* phlstr, LPVOID pb, USHORT cbLen )
{
   LPLSTRING lps1, lps2;
   LPVOID lpv;
   BOOL fOk;

   // phlstr == NULL is not supported
   assert( phlstr != NULL );
   assert( *phlstr != NULL );

   if ( cbLen == (USHORT)-1 ) {
      // This is a "copy string" request with the HLSTR of the source
      // string in pb

      // Destroy old string
      lps2 = (LPLSTRING)*phlstr;
      if ( lps2->usSize > 0 ) {
         assert( lps2->hMem != NULL );
         WBSubFree( lps2->hMem );
      }

      lps1 = (LPLSTRING)pb;

      if ( lps1->usSize > 0 ) {
         assert( WBCheckGuardByte( lps1 ) );
         lpv = WBSubLock( lps1->hMem );
      }
      else
         lpv = NULL;

      // Create new string
      fOk = WBCreateString( (LPLSTRING)*phlstr, lpv, lps1->usSize );

      if ( lps1->usSize > 0 )
         WBSubUnlock( lps1->hMem );

      if ( !fOk )
         return WBERR_STRINGSPACE;

      // Destroy source HLSTR if temporary
      WBDestroyHlstrIfTemp( (HLSTR)pb );
      return 0;
   }
   else {
      // Create a new string and copy the string data into it
      if ( !WBCreateString( (LPLSTRING)*phlstr, pb, cbLen ) )
         return WBERR_STRINGSPACE;

      return 0;
   }
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
LPSTR WINAPI FAR __export WBDerefHlstr( HLSTR hlstr )
{
   LPLSTRING lps;
   LPSTR lpsz;

   lps = (LPLSTRING)hlstr;
   assert( lps != NULL );

   if ( lps->usSize == 0 ) {
      assert( lps->hMem == NULL );
      lpsz = NULL;
   }
   else {
      assert( lps->hMem != NULL );
      assert( WBCheckGuardByte( lps ) );
      lpsz = (LPSTR)WBSubLock( lps->hMem );
      if ( lpsz == NULL ) {
         assert( FALSE );
         return NULL;
      }
      WBSubUnlock( lps->hMem );
   }

   return lpsz;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
LPSTR WINAPI FAR __export WBDerefZeroTermHlstr( HLSTR hlstr )
{
   LPSTR lpsz;
   USHORT usSize;

   assert( hlstr != NULL );

   lpsz = WBDerefHlstrLen( hlstr, &usSize );
   if ( lpsz == NULL )
      lpsz = "";
   else {
      assert( usSize != 0 );
      *(lpsz + usSize) = '\0';
   }

   return lpsz;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
LPSTR WINAPI FAR __export WBDerefHlstrLen( HLSTR hlstr, USHORT FAR* pcbLen )
{
   LPLSTRING lps;
   LPSTR lpsz;

   assert( hlstr != NULL );
   assert( pcbLen != NULL );

   lps = (LPLSTRING)hlstr;

   *pcbLen = lps->usSize;

   if ( lps->usSize == 0 ) {
      assert( lps->hMem == NULL );
      lpsz = NULL;
   }
   else {
      assert( lps->hMem != NULL );
      assert( WBCheckGuardByte( lps ) );
      lpsz = (LPSTR)WBSubLock( lps->hMem );
      if ( lpsz == NULL ) {
         assert( FALSE );
         *pcbLen = 0;
         return NULL;
      }
      WBSubUnlock( lps->hMem );
   }

   return lpsz;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
LPSTR WINAPI FAR __export WBLockHlstr( HLSTR hlstr )
{
   LPLSTRING lps;
   LPSTR lpsz;

   lps = (LPLSTRING)hlstr;
   assert( lps != NULL );

   if ( lps->usSize == 0 ) {
      assert( lps->hMem == NULL );
      lpsz = NULL;
   }
   else {
      assert( lps->hMem != NULL );
      assert( WBCheckGuardByte( lps ) );
      lpsz = (LPSTR)WBSubLock( lps->hMem );
      if ( lpsz == NULL ) {
         assert( FALSE );
         return NULL;
      }
   }

   return lpsz;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
LPSTR WINAPI FAR __export WBLockHlstrLen( HLSTR hlstr, USHORT FAR* pcbLen )
{
   LPLSTRING lps;
   LPSTR lpsz;

   assert( hlstr != NULL );
   assert( pcbLen != NULL );

   lps = (LPLSTRING)hlstr;

   *pcbLen = lps->usSize;

   if ( lps->usSize == 0 ) {
      assert( lps->hMem == NULL );
      lpsz = NULL;
   }
   else {
      assert( lps->hMem != NULL );
      assert( WBCheckGuardByte( lps ) );
      lpsz = (LPSTR)WBSubLock( lps->hMem );
      if ( lpsz == NULL ) {
         assert( FALSE );
         *pcbLen = 0;
         return NULL;
      }
   }

   return lpsz;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WINAPI FAR __export WBUnlockHlstr( HLSTR hlstr )
{
   LPLSTRING lps;

   assert( hlstr != NULL );

   lps = (LPLSTRING)hlstr;
   if ( lps->hMem != 0 ) {
      assert( lps->usSize != 0 );
      WBSubUnlock( lps->hMem );
   }
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
int WBDestroyOrphanTempHlstr( void )
{
   int i;
   BOOL fOk;
   int nOrphan = 0;

   for ( i = 0; i < MAXTEMPHLSTR; i++ ) {
      if ( g_npTask->TempHlstrTable[i] != NULL ) {
         fOk = WBDestroyHlstrIfTemp( g_npTask->TempHlstrTable[i] );
         // This MUST be successful
         assert( fOk );
         g_npTask->TempHlstrTable[i] = NULL;
         nOrphan++;
      }
   }

#ifdef DEBUGOUTPUT
   if ( nOrphan != 0 ) {
      char szMessage[128];
      wsprintf( szMessage,
         "wn WUPBAS: %u temporary HLSTR(s) freed\r\n",
         nOrphan );
      OutputDebugString( szMessage );
   }
#endif

   return nOrphan;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
BOOL WINAPI FAR __export WBDestroyHlstrIfTemp( HLSTR hlstr )
{
   LPLSTRING lps;
   int i;

   assert( hlstr != NULL );

   // Check for a temporary HLSTR. The string descriptor of temporary HLSTR's
   // are allocated in the default data segment
   if ( HIWORD(hlstr) == g_wDataSeg ) {

      // Search corresponding entry in temporary HLSTR table and
      // free the slot by setting it to NULL
      for ( i = 0; i < MAXTEMPHLSTR; i++ ) {
         if ( g_npTask->TempHlstrTable[i] == hlstr ) {
            g_npTask->TempHlstrTable[i] = NULL;
            break;
         }
      }

      assert( i != MAXTEMPHLSTR );

      lps = (LPLSTRING)hlstr;
      if ( lps->usSize > 0 ) {
         assert( lps->hMem != NULL );
         WBSubFree( lps->hMem );
      }
      else {
         assert( lps->hMem == NULL );
      }
      WBLocalFree( (HLOCAL)LOWORD(lps) );
      return TRUE;
   }

   return FALSE;
}

#ifdef _DEBUG
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBAddGuardByte( LPLSTRING lps )
{
   LPBYTE lpb;

   assert( lps != NULL );

   lpb = (LPBYTE)WBSubLock( lps->hMem );
   *(lpb + lps->usSize + 1) = GUARD_BYTE;
   WBSubUnlock( (HMEM)lps->hMem );
}
#endif


#ifdef _DEBUG
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
BOOL WBCheckGuardByte( LPLSTRING lps )
{
   LPBYTE lpb;
   BOOL fOk;

   assert( lps != NULL );

   // Check if guard byte is still there
   lpb = (LPBYTE)WBSubLock( lps->hMem );
   if (lpb == NULL)
      assert( FALSE );
   fOk = ( *(lpb + lps->usSize + 1) == GUARD_BYTE );
   WBSubUnlock( lps->hMem );
   return fOk;
}
#endif










