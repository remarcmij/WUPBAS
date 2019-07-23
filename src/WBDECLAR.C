#include <windows.h>
#include <string.h>
#include "wupbas.h"

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
NPVOID WBLookupDeclaration( LPSTR lpszName, UINT FAR* lpfuDeclType )
{
   NPWBDECLARATION npEntry;
   
   for ( npEntry = g_npTask->DeclarationTable[WBHashName(lpszName, DECL_TABLE_SIZE)];
         npEntry != NULL; 
         npEntry = npEntry->npNext ) {
               
      if ( strcmp( npEntry->szName, lpszName ) == 0 ) {
         *lpfuDeclType = npEntry->fuDeclType;
         return npEntry->npExtra.VoidPtr;
      }
   }

   *lpfuDeclType = DECL_NOTFOUND;   
   return NULL;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBAddDeclaration( NPWBDECLARATION npDecl, LPSTR lpszName, UINT fuDeclType )
{
   UINT uHash;
              
   strcpy( npDecl->szName, lpszName );
   npDecl->fuDeclType = fuDeclType;
                 
   uHash = WBHashName( npDecl->szName, DECL_TABLE_SIZE );
   npDecl->npNext = g_npTask->DeclarationTable[uHash];
   g_npTask->DeclarationTable[uHash] = npDecl;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBDestroyDeclarationTable( void )
{
   NPWBDECLARATION npDecl1, npDecl2;
   NPWBTYPEMEMBER npMember1, npMember2;
   NPWBARGATTRIB npArg1, npArg2;
   int i;
   
   for ( i = 0; i < DECL_TABLE_SIZE; i++ ) {
   
      npDecl1 = g_npTask->DeclarationTable[i];
      
      while ( npDecl1 != NULL ) {
         
         npDecl2 = npDecl1->npNext;
         
         switch ( npDecl1->fuDeclType ) {
         
            case DECL_USERTYPE:
               npMember1 = npDecl1->npExtra.UserType->npMemberList;
               while ( npMember1 != NULL ) {
                  npMember2 = npMember1->npNext;
                  WBLocalFree( (HLOCAL)npMember1 );
                  npMember1 = npMember2;
               }
               break;
               
            case DECL_DECLARE:
               npArg1 = npDecl1->npExtra.Declare->npArgList;
               while ( npArg1 != NULL ) {
                  npArg2 = npArg1->npNext;
                  WBLocalFree( (HLOCAL)npArg1 );
                  npArg1 = npArg2;
               }
               break;
         }
         
         WBLocalFree( (HLOCAL)npDecl1 );
         npDecl1 = npDecl2;
      }
      
      g_npTask->DeclarationTable[i] = NULL;
   }
}
