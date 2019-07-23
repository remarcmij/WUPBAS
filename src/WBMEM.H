/* wbmem.h */

#ifndef _WBMEM_H_
#define _WBMEM_H_

HANDLE WBSubAllocInit( void );
void WBDestroyLocalHeap( HANDLE hpool );
HMEM   WBSubAlloc( HANDLE hpool, UINT fuFlags, USHORT cb );
HMEM   WBSubFree( HMEM hmem );
LPVOID WBSubLock( HMEM hmem );
BOOL   WBSubUnlock( HMEM hmem );
HMEM   WBSubReAlloc( HMEM hmem, USHORT cb, UINT fuFlags );
void WBEnableLocalCompact( BOOL fEnable );
#ifdef _DEBUG
void WBSubCheckMemoryLeaks( void );
void WBLocalCheckMemoryLeaks( void );
#endif

#endif
