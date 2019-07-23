#define STRICT
#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include "wupbas.h"

extern char NEAR g_szDLLPath[];

static BOOL fConflicts;
static ERR errEnumFunctions;

typedef int (CALLBACK FAR* LPENUMCALLBACK)( LPSTR, LPSTR, LPFUNCTION );
typedef void (WINAPI FAR* LPENUMFUNCTIONS)( LPENUMCALLBACK );

int WINAPI __export
WBRegisterFunction( LPSTR lpszName, LPSTR lpszTemplate, LPFUNCTION lpfn );

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBLoadDLLStatement( void )
{
   VARIANT vResult;
   LPSTR lpszFileName;
   char szFileName[_MAX_PATH];
   LPSTR lpch;
   UINT uPrevErrorMode;
   HINSTANCE hinstDLL;
   HMODULE hmodDLL;
   LPENUMFUNCTIONS lpEnumFunctions;
   NPDLLLISTENTRY npEntry;
   ERR err;

   errEnumFunctions = 0;

   if ( (err = WBExpression( &vResult )) != 0 )
      WBRuntimeError( err );
      
   if ( (err = WBVariantToString( &vResult )) != 0 )
      WBRuntimeError( err );

   lpszFileName = WBDerefZeroTermHlstr( vResult.var.tString );
   
   // Extract name.ext part from file name
   lpch = strrchr( lpszFileName, '\\' );  
   if ( lpch != NULL )
      lpch++;
   else
      lpch = lpszFileName;

   // Walk down the list of already loaded DLL's to see if the 
   // target DLL is already loaded      
   for ( npEntry = g_npTask->npDLLList; 
         npEntry != NULL; 
         npEntry = npEntry->npNext ) {

      // If the DLL is already loaded, we're done already         
      if ( stricmp( npEntry->szFileName, lpch ) == 0 ) 
         return;
   }
   
   strcpy( szFileName, lpszFileName );
   _strupr( szFileName );

   uPrevErrorMode = SetErrorMode( SEM_NOOPENFILEERRORBOX );
   hinstDLL = LoadLibrary( szFileName );
   SetErrorMode( uPrevErrorMode );
   
   WBDestroyHlstrIfTemp( vResult.var.tString );   
   
   if ( hinstDLL < HINSTANCE_ERROR )
      WBRuntimeError( WBERR_LOADEXTDLL );

   hmodDLL = GetModuleHandle( (LPSTR)MAKELONG(hinstDLL, NULL) );
   
   lpEnumFunctions = (LPENUMFUNCTIONS)GetProcAddress( hinstDLL, 
                                                      "EnumFunctions" );
   if ( lpEnumFunctions == NULL ) {
      FreeLibrary( hinstDLL );
      WBRuntimeError( WBERR_INVALID_DLL );
   }
   
   npEntry = (NPDLLLISTENTRY)WBLocalAlloc( LPTR, sizeof(DLLLISTENTRY) );
   if ( npEntry == NULL ) {
      FreeLibrary( hinstDLL );
      WBRuntimeError( WBERR_OUTOFMEMORY );
   }

   GetModuleFileName( hinstDLL, szFileName, sizeof( szFileName ) );
   lpch = strrchr( szFileName, '\\' );
   assert( lpch != NULL );
   
   npEntry->hinstDLL = hinstDLL;
   npEntry->hmodDLL = hmodDLL;
   strcpy( npEntry->szFileName, lpch+1 );
   
   npEntry->npNext = g_npTask->npDLLList;
   g_npTask->npDLLList = npEntry;
   
   lpEnumFunctions( (LPENUMCALLBACK)WBRegisterFunction );   
   
   if ( errEnumFunctions != 0 ) 
      WBRuntimeError( errEnumFunctions );
      
   WBCheckEndOfLine();
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
int WINAPI __export
WBRegisterFunction( LPSTR lpszName, LPSTR lpszTemplate, LPFUNCTION lpfn )
{
   NPWBHASHENTRY npHash;
   UINT fuDeclType;
   NPWBDECLARATION npDecl;
   NPWBINTFUNC npIntFunc;
   NPFUNCTIONINFO npFuncInfo;
   UINT uSize;

   npHash.VoidPtr = WBLookupDeclaration( lpszName, &fuDeclType );
   if ( fuDeclType != DECL_NOTFOUND ) {
      errEnumFunctions = WBERR_EXTFUNCCONFLICT;
      return FALSE;
   }
   
   uSize = sizeof(WBDECLARATION) + strlen(lpszName);
            
   // Adjust size for word alignment of extra data
   if ( uSize & 0x0001 )
      uSize++;
               
   npDecl = (NPWBDECLARATION)WBLocalAlloc( LPTR, 
                                           uSize + sizeof(NPWBINTFUNC)
                                                 + sizeof(FUNCTIONINFO) );
   if ( npDecl == NULL ) {
      errEnumFunctions = WBERR_OUTOFMEMORY;
      return FALSE;
   }
      
   npIntFunc = (NPWBINTFUNC)(((NPSTR)npDecl) + uSize);
   npDecl->npExtra.IntFunc = npIntFunc;
   npFuncInfo = (NPFUNCTIONINFO)(((NPSTR)npIntFunc) + sizeof(NPWBINTFUNC));
   npIntFunc->npIntFunc = npFuncInfo;
   
   strcpy( npFuncInfo->szName, lpszName );
   strcpy( npFuncInfo->szTemplate, lpszTemplate );
   npFuncInfo->lpfn = lpfn;
   
   WBAddDeclaration( npDecl, lpszName, DECL_INTFUNC );
   
   return TRUE;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBUnloadExplicitLinks( void )
{
   NPDLLLISTENTRY npNext, npTemp;
   
   // Free loaded DLL's and DLL list entries
   npNext = g_npTask->npDLLList;
   
   while ( npNext != NULL ) {
      FreeLibrary( npNext->hinstDLL );
      npTemp = npNext->npNext;
      WBLocalFree( (HLOCAL)npNext );
      npNext = npTemp;
   }
}