#define STRICT
#include <windows.h>
#include <string.h>
#include "wupbas.h"
#include "wbmem.h"

#define INITIALHEAPSIZE    4096
#define HEAP_WATERMARK     1024
#define GARBAGE            0xf9

HANDLE WBSubNewSegment( void );

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
HANDLE WBSubAllocInit( void )
{
   NPLOCALHEAPINFO npHeap;
   
   // Allocate head node for the linked list of local heap
   // descriptors belonging to this memory sub-allocation pool.
   npHeap = (NPLOCALHEAPINFO)WBLocalAlloc( LPTR, sizeof( LOCALHEAPINFO ) );
   if ( npHeap == NULL )
      return NULL;

   // Cross-link the head node to itself      
   npHeap->npNext = npHeap->npPrev = npHeap;
   
   // Return the near address of the head node as a handle to the 
   // memory sub-allocation pool.
   return (HANDLE)npHeap;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBDestroyLocalHeap( HANDLE hpool )
{
   NPLOCALHEAPINFO npHead, npHeap, npTemp;

   // Note that the head node itself has no heap segment assigned
   // to it.
   npHead = (NPLOCALHEAPINFO)hpool;
   npHeap = npHead->npNext;   
   
   while ( npHeap != npHead ) {
   
      // Free heap segment
      GlobalUnlock( npHeap->hglobal );
      GlobalFree( npHeap->hglobal );
      
      // Free heap descriptor
      npTemp = npHeap->npNext;
      WBLocalFree( (HLOCAL)npHeap );
      npHeap = npTemp;
   }
   
   // Free heap descriptor head node
   WBLocalFree( (HLOCAL)npHead );
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
HANDLE WBSubNewSegment( void )
{
   NPLOCALHEAPINFO npHeap;
   LPVOID lpv;
   WORD wSize;
   BOOL fOk;

   // Allocate new local heap descriptor, zero initialized
   npHeap = (NPLOCALHEAPINFO)WBLocalAlloc( LPTR, sizeof( LOCALHEAPINFO ) );
   if ( npHeap == NULL )
      return NULL;

   // Allocate a global segment in which to create the new local heap.
   // Note that Windows will grow the local heap automatically as
   // allocations are made, until the local heap has grown to it's
   // maximum size of 64KB.
   npHeap->hglobal = GlobalAlloc( GMEM_MOVEABLE | GMEM_ZEROINIT, 
                                  INITIALHEAPSIZE );
   if ( npHeap->hglobal == NULL )
      return NULL;
   
   // Lock the global segment. It is kept locked until program termination.
   lpv = GlobalLock( npHeap->hglobal );
   if ( lpv == NULL ) {
      GlobalFree( npHeap->hglobal );
      return NULL;
   }

   // Save the selector of the segment in the local heap descriptor.
   npHeap->wSeg = HIWORD( lpv );
   
   // Get the actual size of the segment, in preparation for 
   // LocalInit. Reserve 16 bytes as required by LocalInit.
   wSize = (WORD)GlobalSize( npHeap->hglobal ) - 16;
   
   // Initialize the new local heap
   fOk = LocalInit( npHeap->wSeg, 0, wSize );
   
   if ( fOk ) {
      // Undo lock left by LocalInit (one lock remaining)
      GlobalUnlock( npHeap->hglobal );                   
      
#ifdef DEBUGOUTPUT
      OutputDebugString( "t WUPBAS: new local heap segment allocated\r\n" );
#endif   
      // Mark the heap as avaible for allocations
      npHeap->fAvailable = TRUE;
      
      // Insert the heap descriptor in the list of all local heaps.
      npHeap->npLink = g_npTask->npLocalHeapList;
      g_npTask->npLocalHeapList = npHeap;
      
      // Return the near address of the heap as a handle.
      return (HANDLE)npHeap;
   }

   // If heap initialization or heap list entry allocation failed      
   // clean-up and return NULL.
   
   GlobalUnlock( npHeap->hglobal );
   GlobalFree( npHeap->hglobal );
   WBLocalFree( (HLOCAL)npHeap );
   return NULL;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
HMEM WBSubAlloc( HANDLE hpool, UINT fuFlags, USHORT cb )
{
   register NPLOCALHEAPINFO npHeap;
   NPLOCALHEAPINFO npHead;
   HLOCAL hloc;
   WORD wSeg;
   UINT fuFlagsExtra;

   assert( hpool != NULL );
      
   // Consider this an error in the debug version.
   assert( cb != 0 );

   // Must not request more than the maximum size of a local
   // heap (approx. 64KB).
   if ( cb > MAXSTRING )
      return NULL;

   // Add LMEM_NOCOMPACT to the allocation flags is heap compaction
   // has been turned off globally.   
   fuFlagsExtra = (UINT)(g_npTask->fNoLocalCompact ? LMEM_NOCOMPACT : 0 );

   // Walk down the chain of local heap segments until a segment is
   // found that is available, and in which the requested allocation
   // can be successfully made.   
   npHead = (NPLOCALHEAPINFO)hpool;
   npHeap = npHead->npNext;
   
   while ( npHeap != npHead ) {

      if ( npHeap->fAvailable ) {
      
         wSeg = npHeap->wSeg;
               
         _asm {   push  ds
                  mov   ax, wSeg
                  mov   ds, ax   }
         
         hloc = LocalAlloc( fuFlags | fuFlagsExtra, (UINT)cb );
               
#ifdef _DEBUG         
         // In the debug version, fill the allocated memory with defined
         // garbage for debugging purposes. (Unless zero-initialized memory
         // was requested.)
         if ( hloc != NULL && (fuFlags & LMEM_ZEROINIT) == 0 ) {
            USHORT usSize;
            void NEAR* pv;
            usSize = (USHORT)LocalSize( hloc );
            pv = LocalLock( hloc );
            memset( pv, GARBAGE, (size_t)usSize );
            LocalUnlock( hloc );
         }
#endif         
            
         _asm {   pop   ds }
            
         if ( hloc != NULL ) {
#ifdef _DEBUG         
            // In the debug version, keep a count of allocations so that
            // memory leaks can be reported.
            npHeap->cObjects++;
#endif            
            // Return a handle to the allocated memory. This handle contains
            // the heap segment selected in the high order word and the
            // handle returned by LocalAlloc in the low order word. However
            // the caller should treat the return value as just a handle.
            return (HMEM)MAKELONG( hloc, wSeg );
         }
         else if ( g_npTask->fNoLocalCompact && cb < HEAP_WATERMARK ) {
            // If the allocation was unsuccessful, and the size of the
            // requested memory block was less thane the watermark, then
            // mark the heap as unavailable. This means that no new 
            // allocations will be attempted in this heap.
            npHeap->fAvailable = FALSE;
         }
      }
      
      // Go check next heap segment
      npHeap = npHeap->npNext;
   }

   // If no allocation could be made in one of the available heap segments
   // already available, create a new heap segment and add it to the list
   // of heap segment for the current memory sub-allocation pool.   
   npHeap = (NPLOCALHEAPINFO)WBSubNewSegment();
   if ( npHeap != NULL ) {
   
      // Insert new heap descriptor at top of the list of heap descriptors
      // for the current heap pool.
      npHeap->npNext = npHead->npNext;
      npHeap->npPrev = npHead;
      npHead->npNext->npPrev = npHeap;
      npHead->npNext = npHeap;
      npHeap->hpool  = hpool;

      // Try and allocate memory from the new segment      
      wSeg = npHeap->wSeg;
      
      _asm {   push  ds
               mov   ax, wSeg
               mov   ds, ax   }
               
      hloc = LocalAlloc( fuFlags, (UINT)cb );
      
#ifdef _DEBUG         
      // In the debug version, fill the allocated memory with defined
      // garbage for debugging purposes. (Unless zero-initialized memory
      // was requested.)
      if ( hloc != NULL && (fuFlags & LMEM_ZEROINIT) == 0 ) {
         USHORT usSize;
         void NEAR* pv;
         usSize = (USHORT)LocalSize( hloc );
         pv = LocalLock( hloc );
         memset( pv, GARBAGE, (size_t)usSize );
         LocalUnlock( hloc );
      }
#endif         
      
      _asm {   pop   ds }
      
      if ( hloc != NULL ) {
#ifdef _DEBUG      
         // In the debug version, keep a count of allocations so that
         // memory leaks can be reported.
         npHeap->cObjects++;
#endif         
         // Return a handle to the allocated memory. This handle contains
         // the heap segment selected in the high order word and the
         // handle returned by LocalAlloc in the low order word. However
         // the caller should treat the return value as just a handle.
         return (HMEM)MAKELONG( hloc, wSeg );
      }
   }
    
   // If all fails, return NULL.
   return NULL;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
HMEM WBSubReAlloc( HMEM hmem, USHORT cb, UINT fuFlags )
{
   register NPLOCALHEAPINFO npHeap;
   WORD wSeg;
   HLOCAL hloc, hlocNew;
   HMEM hmemNew;
   LPVOID lpv1, lpv2;
   UINT fuFlagsExtra;
   USHORT usSizeOld;
#ifdef _DEBUG   
   USHORT usSizeNew;
#endif   
   
   assert( cb != 0 );
      
   // Extract segment selector and local handle from HMEM handle   
   wSeg = HIWORD( hmem );
   hloc = (HLOCAL)LOWORD( hmem );
   
   assert( wSeg != 0 );
   assert( hloc != NULL );
   
   // Search the list of local heap descriptors for a      
   // descriptor which matches the heap segment selector.
   npHeap = g_npTask->npLocalHeapList; 
   while ( npHeap != NULL ) {
      if ( npHeap->wSeg == wSeg )
         break;
      npHeap = npHeap->npLink;
   }
   assert( npHeap != NULL );
   
   // Add LMEM_NOCOMPACT to the allocation flags is heap compaction
   // has been turned off globally.   
   fuFlagsExtra = (UINT)(g_npTask->fNoLocalCompact ? LMEM_NOCOMPACT : 0 );
   
   // If the heap has been marked as available, try a resize in the current
   // local heap.
   if (  npHeap->fAvailable ) {
   
      _asm  {  push  ds
               mov   ax, wSeg
               mov   ds, ax   }
               
      usSizeOld = (USHORT)LocalSize( hloc );
      
#ifdef _DEBUG   
      // In the debugging version, if the memory block is shrinking, fill
      // the memory freed with defined garbage.
      if ( cb < usSizeOld ) {
         void NEAR* pv;
         pv = LocalLock( hloc );
         memset( (NPSTR)pv + cb, GARBAGE, (size_t)(usSizeOld-cb) );
         LocalUnlock( hloc );
      }
#endif      

      hlocNew = LocalReAlloc( hloc, (UINT)cb, fuFlags | fuFlagsExtra );
      if ( hlocNew != NULL ) 
         hloc = hlocNew;
         
#ifdef _DEBUG   
      // In the debugging version, if the memory block has grown, fill
      // the memory added with defined garbage (unless zero-initialized
      // memory was requested).      
      usSizeNew = (USHORT)LocalSize( hloc );
      if ( usSizeNew > usSizeOld && (fuFlags & LMEM_ZEROINIT) == 0  ) {
         void NEAR* pv;
         pv = LocalLock( hloc );
         memset( (NPSTR)pv + usSizeOld, GARBAGE, (size_t)(usSizeNew-usSizeOld) );
         LocalUnlock( hloc );
      }
#endif      
   
      _asm  {  pop   ds       }
   }
   
   // Return a handle to the allocated memory. This handle contains
   // the heap segment selected in the high order word and the
   // handle returned by LocalAlloc in the low order word. However
   // the caller should treat the return value as just a handle.
   if ( hlocNew != NULL ) 
      return (HMEM)MAKELONG( hloc, wSeg );
      
   // At this point the resize operation failed. If in non-compacting mode,
   // flag the heap as not available.
   if ( g_npTask->fNoLocalCompact )
      npHeap->fAvailable = FALSE;
      
   // Try allocating in a different local heap segment from the same pool,
   // and, if successful, copy the data from the old block to the new block.
   hmemNew = WBSubAlloc( npHeap->hpool, fuFlags, cb );
   if ( hmemNew != NULL ) {
      lpv1 = WBSubLock( hmem );
      lpv2 = WBSubLock( hmemNew );
      memcpy( lpv2, lpv1, usSizeOld );
      WBSubUnlock( hmem );
      WBSubUnlock( hmemNew );
      WBSubFree( hmem );
   }
   
   return hmemNew;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
HMEM WBSubFree( HMEM hmem )
{
   register NPLOCALHEAPINFO npHeap;
   WORD wSeg;
   HLOCAL hloc;
   USHORT usSize;

   // Extract segment selector and local handle from HMEM handle   
   wSeg = HIWORD( hmem );
   hloc = (HLOCAL)LOWORD( hmem );

   assert( wSeg != 0 );
   assert( hloc != NULL );
   
   _asm  {  push  ds
            mov   ax,wSeg
            mov   ds, ax   }

   usSize = (USHORT)LocalSize( hloc );
   
#ifdef _DEBUG   
   // In the debug version, fill the freed block with defined garbage.
   // Continued use after freeing the block will then almost certainly
   // be detected. 
   if ( usSize > 0 ) {
      void NEAR* pv;
      pv = LocalLock( hloc );   
      memset( pv, GARBAGE, (size_t)usSize );
      LocalUnlock( hloc );
   }
#endif   

   hloc = LocalFree( hloc );
   
   _asm  {  pop   ds       }

   // If freeing was unsuccessful, return the value of handle argument. 
   if ( hloc != NULL )
      return hmem;

   // Search the list of local heap descriptors for a      
   // descriptor which matches the heap segment selector.
   npHeap = g_npTask->npLocalHeapList;
   while ( npHeap != NULL ) {
      if ( npHeap->wSeg == wSeg )
         break;
      npHeap = npHeap->npLink;
   }
   assert( npHeap != NULL );
   
   // If the heap was marked unavailable and the size of the freed memory
   // is greater or equal than the watermark, mark the heap as available
   // again.
   if ( !npHeap->fAvailable && usSize >= HEAP_WATERMARK )
      npHeap->fAvailable = TRUE;
   
#ifdef _DEBUG
   // In the debug version, decrement the allocation counter used for 
   // memory leak detection.
   npHeap->cObjects--;
#endif
      
   // Return NULL to indicate success.
   return NULL;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
LPVOID WBSubLock( HMEM hmem )
{
   WORD wSeg;
   HLOCAL hloc;
   void NEAR* npv;
   
   wSeg = HIWORD( hmem );
   hloc = (HLOCAL)LOWORD( hmem );

   assert( wSeg != 0 );
   assert( hloc != NULL );
   
   _asm  {  push  ds
            mov   ax,wSeg
            mov   ds, ax   }

   npv = LocalLock( hloc );
   
   _asm  {  pop   ds       }
   
   if ( npv == NULL )
      return NULL;
   else
      return (LPVOID)MAKELONG( npv, wSeg );
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
BOOL WBSubUnlock( HMEM hmem )
{
   WORD wSeg;
   HLOCAL hloc;
   BOOL fOk;
   
   wSeg = HIWORD( hmem );
   hloc = (HLOCAL)LOWORD( hmem );

   assert( wSeg != 0 );
   assert( hloc != NULL );
   
   _asm  {  push  ds
            mov   ax,wSeg
            mov   ds, ax   }

   fOk = LocalUnlock( hloc );
   
   _asm  {  pop   ds       }

   return fOk;   
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBEnableLocalCompact( BOOL fEnable )
{
   NPLOCALHEAPINFO npHeap;
   WORD wSeg;
   
   if ( !fEnable ) {
      // Disable local compaction
      g_npTask->fNoLocalCompact = TRUE;
      return;
   }
   
   // Compact all local heaps (except default)
   npHeap = g_npTask->npLocalHeapList;   
   while ( npHeap != NULL ) {
   
      wSeg = npHeap->wSeg;
      
      _asm  {  push  ds
               mov   ax,wSeg
               mov   ds, ax   }
   
      LocalCompact( (UINT)-1 );
      
      _asm  {  pop   ds       }

      npHeap->fAvailable = TRUE;
      npHeap = npHeap->npLink;
   }
   
   g_npTask->fNoLocalCompact = FALSE;
}

#ifdef DEBUGOUTPUT
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBSubCheckMemoryLeaks( void )
{
   NPLOCALHEAPINFO npHeap;

   npHeap = g_npTask->npLocalHeapList;   

   while ( npHeap != NULL ) {   
      if ( npHeap->cObjects != 0 ) {
         char szMessage[128];
         wsprintf( szMessage,
            "wn WUBBAS: memory leak, heap segment %X, %u memory object(s) not freed\r\n",
            npHeap->wSeg,
            npHeap->cObjects );
         OutputDebugString( szMessage );
      }
      npHeap = npHeap->npLink;
   }
}      
#endif
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
#ifdef _DEBUG      
HLOCAL WBLocalAllocDebug( UINT fuAllocFlags, UINT fuAlloc )
{
   HLOCAL hloc;
   UINT uSize;
   void NEAR* pv;
   
   // Consider this an error
   assert( fuAlloc != 0 );

   hloc = LocalAlloc( fuAllocFlags, fuAlloc );
   
   if ( hloc != NULL ) {
   
      uSize = LocalSize( hloc );
      if ( uSize == 0 )
         assert( uSize != 0 );
      
      // If requested block is not zero-initialized then fill it
      // with defined garbage
      if ( (fuAllocFlags & LMEM_ZEROINIT) == 0 ) {
         pv = LocalLock( hloc );
         memset( pv, GARBAGE, (size_t)uSize );
         LocalUnlock( hloc );
      }
      
      g_npTask->cObjects++;
   }
   
   return hloc;
}
#endif      
      
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
#ifdef _DEBUG      
HLOCAL WBLocalReAllocDebug( HLOCAL hloc, UINT fuNewSize, UINT fuFlags )
{
   HLOCAL hlocNew;
   UINT uSizeOld, uSizeNew;
   
   uSizeOld = LocalSize( hloc );
   assert( uSizeOld != NULL );

   hlocNew = LocalReAlloc( hloc, fuNewSize, fuFlags );
   
   if ( hlocNew != NULL ) {
      uSizeNew = LocalSize( hlocNew );
      assert( uSizeNew != 0 );
   }
   
   return hlocNew;
}
#endif      
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
#ifdef _DEBUG
HLOCAL WBLocalFreeDebug( HLOCAL hloc )
{
   UINT uSize;
   void NEAR* pv;

   assert( hloc != NULL );
      
   uSize = LocalSize( hloc );
   assert( uSize != 0 );

   // Fill freed memory with defined garbage
   pv = LocalLock( hloc );
   memset( pv, GARBAGE, (size_t)uSize );
   LocalUnlock( hloc );
   
   hloc = LocalFree( hloc );
   assert( hloc == NULL );
   
   g_npTask->cObjects--;
   
   return hloc;
}
#endif   

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
#ifdef DEBUGOUTPUT
void WBLocalCheckMemoryLeaks( void )
{
   
   if ( g_npTask->cObjects != 0 ) {
      char szMessage[128];
      wsprintf( szMessage,
                "wn WUBBAS: memory leak, %u local heap object(s) not freed\r\n",
                g_npTask->cObjects );
      OutputDebugString( szMessage );
   }
}   
#endif
