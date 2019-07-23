#define STRICT
#include <windows.h>
#include <string.h>
#include <limits.h>
#include "wupbas.h"
#include "wbtype.h"
#include "wbmem.h"

static void WBInitMemberList( LPBYTE lpByte, NPWBTYPEMEMBER npMember );

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBTypeStatement( void )
{
   TOKEN tok;
   char szTypeName[MAXTOKEN];
   NPWBDECLARATION npDecl;
   NPWBUSERTYPE npUserType;
   NPWBHASHENTRY npHash;
   UINT fuDeclType;
   NPWBTYPEMEMBER npMember, npTail;
   DWORD dwTotalSize = 0L;
   UINT uOffset = 0;
   UINT uSize;
   
   tok = LastToken();
   if ( tok != TOKEN_IDENT ) 
      return WBERR_EXPECT_IDENT;

   strcpy( szTypeName, GetTokenText() );

   npHash.VoidPtr = WBLookupDeclaration( szTypeName, &fuDeclType ); 
   if ( fuDeclType != DECL_NOTFOUND )
      return WBERR_DUPLICATEDEFINITION;

   uSize = sizeof(WBDECLARATION) + strlen( szTypeName );
   
   // Adjust size for alignment of extra data
   if ( uSize & 0x0001 )
      uSize++;
      
   npDecl = (NPWBDECLARATION)WBLocalAlloc( LPTR, uSize + sizeof(WBUSERTYPE) );
   if ( npDecl == NULL )
      return WBERR_OUTOFMEMORY;
      
   npUserType = (NPWBUSERTYPE)(((NPSTR)npDecl) + uSize);
   npDecl->npExtra.UserType = npUserType;
      
   tok = NextToken();
   if ( tok != TOKEN_EOL ) 
      return WBERR_GARBAGE;

   npTail = NULL;
         
   tok = NextToken();
   
   while ( tok != TOKEN_EOF && tok != TOKEN_ENDTYPE ) {
      
      if ( tok != TOKEN_IDENT ) 
         return WBERR_EXPECT_IDENT;
         
      npMember = (NPWBTYPEMEMBER)WBLocalAlloc( LPTR, sizeof(WBTYPEMEMBER) + strlen(GetTokenText()) );
      if ( npMember == NULL )
         return WBERR_OUTOFMEMORY;

      strcpy( npMember->szMemberName, GetTokenText() );

      tok = NextToken();
      if ( tok != TOKEN_AS )
         return WBERR_EXPECT_AS;
         
      tok = NextToken();
      if ( tok == TOKEN_IDENT && strcmp( GetTokenText(), "STRING" ) == 0 )
         tok = TOKEN_STRING;
         
      switch ( tok ) {
         
         case TOKEN_INTEGER:
            npMember->fMemberType = VT_I2;
            npMember->uSize = 2;
            npMember->uOffset = uOffset;
            uOffset += 2;
            dwTotalSize += 2L;
            break;
            
         case TOKEN_LONG:
            npMember->fMemberType = VT_I4;
            npMember->uSize = 4;
            npMember->uOffset = uOffset;
            uOffset += 4;
            dwTotalSize += 4L;
            break;
            
         case TOKEN_STRING:
            tok = NextToken();
            if ( tok != TOKEN_MULT ) 
               return WBERR_EXPECT_MULT;
            tok = NextToken();
            if ( tok != TOKEN_LONGINT )
               return WBERR_EXPECT_SIZE;
            if ( LastInteger() > UINT_MAX )
               return WBERR_NUMBERTOOBIG;
            uSize = (UINT)LastInteger();
            npMember->fMemberType = VT_STRING;
            npMember->uSize = uSize;
            npMember->uOffset = uOffset;
            uOffset += uSize;
            dwTotalSize += (DWORD)uSize;
            break;
            
         case TOKEN_IDENT:
            {
               NPWBUSERTYPE npChild;
               npHash.VoidPtr = WBLookupDeclaration( GetTokenText(), &fuDeclType );
               if ( fuDeclType != DECL_USERTYPE )
                  return WBERR_UNDEFINEDTYPE;
                  
               npChild = npHash.UserType;
               npMember->fMemberType = VT_USERTYPE;
               npMember->uSize = uSize = npChild->uSize;
               npMember->uOffset = uOffset;
               npMember->npChild = npChild;               
               uOffset += uSize;                                           
               dwTotalSize += (DWORD)uSize;
            }
            break;
            
         default:
            return WBERR_EXPECTTYPE;
      }
      
      if ( npTail != NULL )
         npTail->npNext = npMember;
      npTail = npMember;
      if ( npUserType->npMemberList == NULL )
         npUserType->npMemberList = npMember;
      
      tok = NextToken();
      if ( tok != TOKEN_EOL )
         return WBERR_GARBAGE;
         
      tok = NextToken();
   }
    
   if ( tok == TOKEN_EOF )
      return WBERR_UNEXPECTED_EOF;
      
   tok = NextToken();
   if ( tok != TOKEN_EOL )
      return WBERR_GARBAGE;
   
   if ( dwTotalSize > (DWORD)UINT_MAX )
      return WBERR_TYPETOOLARGE;
      
   npUserType->uSize = (UINT)dwTotalSize;

   WBAddDeclaration( npDecl, szTypeName, DECL_USERTYPE );   
   
   return 0;
}   

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
NPWBTYPEMEMBER WBFindTypeMember( NPWBUSERTYPE npUserType, LPSTR lpszName )
{
   NPWBTYPEMEMBER npMember;
   
   npMember = npUserType->npMemberList;
   
   while ( npMember != NULL ) {
      if ( strcmp( npMember->szMemberName, lpszName ) == 0 )
         break;
      npMember = npMember->npNext;
   }
   
   return npMember;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
LPWBTYPEMEM WBAllocTypeMem( NPWBUSERTYPE npUserType, BOOL fInitData )
{
   HMEM hmem;
   LPWBTYPEMEM lpTypeMem;
   LPBYTE lpByte;
   
   hmem = WBSubAlloc( g_npTask->hTypePool, 
                      LMEM_MOVEABLE, 
                      (USHORT)(sizeof(WBTYPEMEM) - 1 + npUserType->uSize) );
   if ( hmem == NULL )
      WBRuntimeError( WBERR_OUTOFMEMORY );
      
   lpTypeMem = (LPWBTYPEMEM)WBSubLock( hmem );
   lpTypeMem->npUserType = npUserType;
   lpTypeMem->hmem = hmem;

   if ( fInitData ) {   
      lpByte = lpTypeMem->data;
      memset( lpByte, 0, npUserType->uSize );
      WBInitMemberList( lpByte, npUserType->npMemberList );
   }
   
   return lpTypeMem;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
static void WBInitMemberList( LPBYTE lpByte, NPWBTYPEMEMBER npMember )
{
   while ( npMember != NULL ) {
      
      switch ( npMember->fMemberType ) {
      
         case VT_STRING:
            memset( lpByte, ' ', npMember->uSize );
            break;
            
         case VT_USERTYPE:
            assert( npMember->npChild != NULL );
            WBInitMemberList( lpByte, npMember->npChild->npMemberList );
            break;            
      }

      lpByte += npMember->uSize;      
      npMember = npMember->npNext;
   }
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBFreeTypeMem( LPWBTYPEMEM lpTypeMem )
{
   WBSubUnlock( lpTypeMem->hmem );
   WBSubFree( lpTypeMem->hmem );
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBGetUserTypeData( LPWBTYPEMEM lpTypeMem, 
                       NPWBUSERTYPE FAR* lpnpUserType, 
                       UINT FAR* lpfMemberType,
                       LPBYTE FAR* lplpData,
                       UINT FAR* lpuSize )
{  
   NPWBUSERTYPE npUserType;
   UINT fMemberType;
   NPWBTYPEMEMBER npMember;
   LPBYTE lpData;
   UINT uSize;
   UINT uOffset;

   npUserType = lpTypeMem->npUserType;
   fMemberType = VT_USERTYPE;
   lpData = lpTypeMem->data;
   uOffset = 0;
   uSize = npUserType->uSize;
   
   if ( LastToken() == TOKEN_PERIOD ) {
   
      if ( NextToken() != TOKEN_IDENT ) 
         return WBERR_EXPECT_IDENT;
         
      npMember = WBFindTypeMember( npUserType, GetTokenText() );
      if ( npMember == NULL )
         return WBERR_UNDEFINEDELEMENT;
         
      while ( NextToken() == TOKEN_PERIOD ) {
      
         if ( npMember->fMemberType != VT_USERTYPE )
            return WBERR_UNDEFINEDELEMENT;
            
         npUserType = npMember->npChild;
         assert( npUserType != NULL );
         uOffset += npMember->uOffset;
         
         if ( NextToken() != TOKEN_IDENT )
            return WBERR_EXPECT_IDENT;
   
         npMember = WBFindTypeMember( npUserType, GetTokenText() );
         if ( npMember == NULL )
            return WBERR_UNDEFINEDELEMENT;
      }
      
      npUserType = npMember->npChild;
      fMemberType = npMember->fMemberType;
      uOffset += npMember->uOffset;
      uSize = npMember->uSize;
   }
   
   *lpnpUserType = npUserType;
   *lpfMemberType = fMemberType;
   *lplpData = lpData + uOffset;
   *lpuSize = uSize;
   
   return 0;
}
                            

