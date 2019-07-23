#define STRICT
#include <windows.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include "wupbas.h"
#include "wbtype.h"
#include "wbmem.h"

#define ARGTYPE_NULL      0
#define ARGTYPE_VOID      1
#define ARGTYPE_INTEGER   2
#define ARGTYPE_LONG      3
#define ARGTYPE_STRING    4
#define ARGTYPE_USER      5
#define ARGTYPE_ANY       6

#define MAXARGLIST        20

union ARGUNION {
   WORD tWord;
   DWORD tDWord;
};

typedef struct tagARGLIST {
   BOOL fWord;
   HLSTR hlstr;
   union ARGUNION arg;
} ARGLIST;

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBDeclareStatement( void )
{
   TOKEN tok;
   char szFuncName[MAXTOKEN];
   char szTrueName[MAXTOKEN];
   char szPathName[_MAX_PATH];
   BOOL fGotLib;
   BOOL fIsFunction;
   UINT fuSysLib;
   UINT fReturnType;
   BOOL fByVal;
   UINT fArgType; 
   NPWBHASHENTRY npHash;
   UINT fuDeclType;
   NPWBDECLARATION npDecl;
   NPWBDECLARE npDeclare;
   NPWBARGATTRIB nparg, nptail;
   LPSTR lpsz, lpch1, lpch2;
   UINT uSize;
   int cArg;
   
   tok = LastToken();

   fReturnType = ARGTYPE_VOID;
   
   if ( tok == TOKEN_FUNCTION )
      fIsFunction = TRUE;
   else if ( tok == TOKEN_SUB ) 
      fIsFunction = FALSE;
   else 
      return WBERR_EXPECTSUBORFUNCTION;     
      
   if ( NextToken() != TOKEN_IDENT ) 
      WBRuntimeError( WBERR_EXPECTPROCNAME );

   strcpy( szFuncName, GetTokenText() ); 
   strcpy( szTrueName, GetTokenText() );
   
   npHash.VoidPtr = WBLookupDeclaration( szFuncName, &fuDeclType );
   if ( fuDeclType != DECL_NOTFOUND )
      return WBERR_DUPLICATEDEFINITION;

   tok = NextToken();
   
   switch ( tok ) {
   
      case TOKEN_CONCAT:
         fReturnType = ARGTYPE_LONG;
         tok = NextToken();
         break;
            
      case TOKEN_PERCENT:
         fReturnType = ARGTYPE_INTEGER;
         tok = NextToken();
         break;
   
      case TOKEN_DOLLAR:
         fReturnType = ARGTYPE_STRING;
         tok = NextToken();
         break;
   }
   
   fGotLib = FALSE;
   fuSysLib = 0;
   
   while ( tok == TOKEN_LIB || tok == TOKEN_ALIAS ) {
   
      if ( tok == TOKEN_ALIAS ) {
      
         if ( NextToken() != TOKEN_STRINGLIT )
            return WBERR_EXPECT_STRINGCONST;
            
         lpsz = GetTokenText();
         while ( isspace( *lpsz ) )
            lpsz++;
            
         uSize = min( strlen( lpsz ), sizeof(szTrueName) - 1 );
         memcpy( szTrueName, lpsz, uSize );
         szTrueName[uSize] = '\0';
         _strupr( szTrueName );
         
         tok = NextToken();
      }
      else {
      
         if ( NextToken() != TOKEN_STRINGLIT )
            return WBERR_EXPECT_STRINGCONST;
            
         lpsz = GetTokenText();
         while ( isspace( *lpsz ) )
            lpsz++;
            
         uSize = min( strlen( lpsz ), sizeof(szPathName) - 1 );
         memcpy( szPathName, lpsz, uSize );
         szPathName[uSize] = '\0';
         _strupr( szPathName );

         // Look for USER, KERNEL and GDI. If one of these three,
         // strip off path information and file extension if present.
         lpch1 = strrchr( szPathName, '\\');
         if ( lpch1 != NULL )
            lpch1++;
         else
            lpch1 = szPathName;
            
         lpch2 = strchr( lpch1, '.');
         if ( lpch2 == NULL )
            lpch2 = strchr( lpch1, '\0');

         uSize = lpch2 - lpch1;
         
         if ( uSize == 4 && memcmp( lpch1, "USER", 4 ) == 0 ) {
            strcpy( szPathName, "USER" );
            fuSysLib = SYSLIB_USER;
         }
         else if ( uSize == 6 && memcmp( lpch1, "KERNEL", 6 ) == 0 ) {
            strcpy( szPathName, "KERNEL" );
            fuSysLib = SYSLIB_KERNEL;
         }
         else if ( uSize == 3 && memcmp( lpch1, "GDI", 3 ) == 0 ) {
            strcpy( szPathName, "GDI" );
            fuSysLib = SYSLIB_GDI;
         }
         
         fGotLib = TRUE;
         tok = NextToken();
      }
   }
   
   if ( !fGotLib )
      return WBERR_EXPECTLIB;      

   uSize = sizeof(WBDECLARATION) + + strlen( szFuncName );
   
   // Adjust size for word aligment of extra data
   if ( uSize & 0x0001 )
      uSize++;
      
   // Allocate memory for Declare Definition structure
   npDecl = (NPWBDECLARATION)WBLocalAlloc( LPTR, uSize + sizeof(WBDECLARE) + strlen(szPathName) );
   if ( npDecl == NULL )
      return WBERR_OUTOFMEMORY;

   npDeclare = (NPWBDECLARE)(((NPSTR)npDecl) + uSize);
   npDecl->npExtra.Declare = npDeclare;
   
   strcpy( npDeclare->szTrueName, szTrueName );
   strcpy( npDeclare->szPathName, szPathName );
   npDeclare->fuSysLib = fuSysLib;
   
   if ( tok != TOKEN_LPAREN )
      return WBERR_EXPECT_LPAREN;
      
   tok = NextToken();
   cArg = 0;
   nptail = NULL;
   
   while ( tok != TOKEN_RPAREN && tok != TOKEN_EOL ) {

      if ( tok == TOKEN_BYVAL ) {
         fByVal = TRUE;
         tok = NextToken();
      }
      else {
         fByVal = FALSE;
      }
               
      if ( tok != TOKEN_IDENT ) 
         return WBERR_EXPECT_PARAMETER;

      tok = NextToken();
      
      switch ( tok ) {
      
         case TOKEN_AS:         
            tok = NextToken();
            if ( tok == TOKEN_IDENT && strcmp( GetTokenText(), "STRING" ) == 0 )
               tok = TOKEN_STRING;
            break;
            
         case TOKEN_CONCAT:
            tok = TOKEN_LONG;
            break;
            
         case TOKEN_DOLLAR:
            tok = TOKEN_STRING;
            break;
            
         case TOKEN_PERCENT:
            tok = TOKEN_INTEGER;
            break;
            
         default:
            return WBERR_EXPECTTYPE;
      }
      
      switch ( tok ) {
      
         case TOKEN_INTEGER:
            fArgType = ARGTYPE_INTEGER;
            break;
            
         case TOKEN_LONG:
            fArgType = ARGTYPE_LONG;
            break;
               
         case TOKEN_STRING:
            fArgType = ARGTYPE_STRING;
            break;
            
         case TOKEN_IDENT:
            npHash.VoidPtr = WBLookupDeclaration( GetTokenText(), &fuDeclType );
            if ( fuDeclType != DECL_USERTYPE )
               return WBERR_UNDEFINEDTYPE;
            fArgType = ARGTYPE_USER;
            break;
                                                      
         case TOKEN_ANY:
            fArgType = ARGTYPE_ANY;
            break;
            
         default:
            //fArgType = 0;   // To avoid compiler warning
            return WBERR_EXPECT_TYPENAME2;
      }
      
      nparg = (NPWBARGATTRIB)WBLocalAlloc( LPTR, sizeof(WBARGATTRIB) );
      if ( nparg == NULL )
         return WBERR_OUTOFMEMORY;

      // Copy argument attribs               
      nparg->fByVal = fByVal;
      nparg->fArgType = fArgType;
      
      if ( fArgType == ARGTYPE_USER ) {
         nparg->npUserType = npHash.UserType;
         if ( fByVal )
            return WBERR_BYREFONLY;
      }
         
      // Insert argument attrib structure into argument list
      if ( npDeclare->npArgList == NULL )
         npDeclare->npArgList = nparg;
               
      if ( nptail != NULL )
         nptail->npNext = nparg;
         
      nptail = nparg;
         
      cArg++;
      
      tok = NextToken();
      
      if ( tok == TOKEN_COMMA ) {
         tok = NextToken();
         if ( tok == TOKEN_EOL )
            tok = NextToken();
      }
      else if ( tok == TOKEN_AS )
         return WBERR_DUPLICATETYPEDEF;
      else if ( tok != TOKEN_RPAREN )
         return WBERR_EXPECT_RPAREN;
   }
   
   npDeclare->cArg = cArg;

   tok = NextToken();
      
   if ( tok == TOKEN_AS ) {
   
      if ( fReturnType != ARGTYPE_VOID )
         return WBERR_DUPLICATETYPEDEF;
   
      tok = NextToken();
      if ( tok == TOKEN_IDENT && strcmp( GetTokenText(), "STRING" ) == 0 )
         tok = TOKEN_STRING;
         
      switch ( tok ) {
      
         case TOKEN_INTEGER:
            fReturnType = ARGTYPE_INTEGER;
            break;
            
         case TOKEN_LONG:
            fReturnType = ARGTYPE_LONG;
            break;
               
         case TOKEN_STRING:
            fReturnType = ARGTYPE_STRING;
            break;
            
         case TOKEN_IDENT:
            npHash.VoidPtr = WBLookupDeclaration( GetTokenText(), &fuDeclType );
            if ( fuDeclType != DECL_USERTYPE )
               return WBERR_UNDEFINEDTYPE;
            else
               return WBERR_CANTRETURNUSERTYPE;
            break;
            
         default:
            return WBERR_EXPECT_TYPENAME2;
      }
      
      tok = NextToken();
   }

   if ( fIsFunction && fReturnType == ARGTYPE_VOID )
      return WBERR_NORETURNTYPE;
   else if ( !fIsFunction && fReturnType != ARGTYPE_VOID )
      return WBERR_CANTRETURNVALUE;
      
   npDeclare->fReturnType = fReturnType;
   
   if ( tok != TOKEN_EOL )
      return WBERR_EXPECT_EOL;
      
   WBAddDeclaration( npDecl, szFuncName, DECL_DECLARE );
   
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBCallDeclaredFunction( NPWBDECLARE npfun, LPVAR lpResult )
{
   ARGLIST ArgList[MAXARGLIST];
   int i;
   HMODULE hmod;
   HINSTANCE hinst;
   FARPROC lpfn;
   TOKEN tok;
   int cArg;
   BOOL fByVal;
   VARIANT vResult;
   ERR err;
   NPWBARGATTRIB nparg;
   WORD wArg;
   DWORD dwArg;
   HLSTR hlstr;
   LPSTR lpsz;
   NPVARDEF npvar;
   LPWBTYPEMEM lpTypeMem;
   NPWBUSERTYPE npUserType;
   UINT fMemberType;
   LPBYTE lpData;
   UINT uSize;
   NPDLLLISTENTRY npEntry;
   char szFileName[_MAX_PATH];
   UINT uPrevErrorMode;
   static WORD wSPBefore;
   static WORD wSPAfter;

   // If procedure already resolved previously, use it directly
   if ( npfun->lpfn != NULL ) {
      lpfn = npfun->lpfn;
   }
   // USER, GDI and KERNEL are always loaded. We can use GetModule Handle
   // directly without doing a LoadLib.
   else if ( npfun->fuSysLib ) {
      hmod = GetModuleHandle( npfun->szPathName );
      if ( hmod == NULL )
         return WBERR_LIBNOTFOUND;
      lpfn = npfun->lpfn = GetProcAddress( hmod, npfun->szTrueName );
      if ( lpfn == NULL )
         return WBERR_PROCADDRESS;
   }
   else {
      // Check whether module  was explicitly loaded previously
      lpsz = strrchr( npfun->szPathName, '\\');
      if ( lpsz != NULL )
         lpsz++;
      else
         lpsz = npfun->szPathName;
         
      npEntry = g_npTask->npDLLList;
      while ( npEntry != NULL ) {
         if ( strcmp( lpsz, npEntry->szFileName ) == 0 )
            break;
         npEntry = npEntry->npNext;
      }

      // If module is in list, use the stored HMODULE       
      if ( npEntry != NULL ) {
         hmod = npEntry->hmodDLL;
         lpfn = npfun->lpfn = GetProcAddress( hmod, npfun->szTrueName );
      }
      // If not found, load the module
      else {
         uPrevErrorMode = SetErrorMode( SEM_NOOPENFILEERRORBOX );
         hinst = LoadLibrary( npfun->szPathName );
         SetErrorMode( uPrevErrorMode );
         
         if ( hinst < HINSTANCE_ERROR )
            return WBERR_LIBNOTFOUND;
            
         hmod = GetModuleHandle( (LPSTR)MAKELONG(hinst, NULL) );
         npfun->lpfn = lpfn = GetProcAddress( hmod, npfun->szTrueName );
      
         // Add module to list of explicitly loaded DLLs so it gets unloaded
         // at deinitialization time
         npEntry = (NPDLLLISTENTRY)WBLocalAlloc( LPTR, sizeof(DLLLISTENTRY) );
         if ( npEntry == NULL ) {
            FreeLibrary( hinst );
            return WBERR_OUTOFMEMORY;
         }
   
         GetModuleFileName( hinst, szFileName, sizeof( szFileName ) );
         lpsz = strrchr( szFileName, '\\' );
         if ( lpsz != NULL )
            lpsz++;
         else
            lpsz = szFileName;
         
         npEntry->hinstDLL = hinst;
         npEntry->hmodDLL  = hmod;
         strcpy( npEntry->szFileName, lpsz );
         
         npEntry->npNext = g_npTask->npDLLList;
         g_npTask->npDLLList = npEntry;
         
         if ( lpfn == NULL )
            return WBERR_PROCADDRESS;
      }
   }
      
   tok = LastToken();
   cArg = 0;
   nparg = npfun->npArgList;
   
   while ( tok != TOKEN_RPAREN && tok != TOKEN_EOL ) {

      if ( cArg >= npfun->cArg || cArg >= MAXARGLIST )
         return WBERR_ARGTOOMANY;
            
      if ( tok == TOKEN_BYVAL ) {
         fByVal = TRUE;
         tok = NextToken();
      }
      else {
         fByVal = FALSE;
      }

      switch ( nparg->fArgType ) {
      
         case ARGTYPE_INTEGER:
         case ARGTYPE_LONG:
            if ( fByVal || nparg->fByVal ) {
               if ( (err = WBExpression( &vResult )) != 0 )
                  return err;
               switch ( vResult.type ) {
                  case VT_I4:
                     break;
                  case VT_STRING:
                     if ( (err = WBMakeLongVariant( &vResult )) != 0 )
                        return err;
                     break;
                  default:
                     return WBERR_TYPEMISMATCH;
               } 
               if ( nparg->fArgType == ARGTYPE_INTEGER ) {
                  ArgList[cArg].fWord = TRUE;
                  ArgList[cArg].arg.tWord = (WORD)(int)vResult.var.tLong;
               }
               else {
                  ArgList[cArg].fWord = FALSE;
                  ArgList[cArg].arg.tDWord = (DWORD)vResult.var.tLong;
               }
               ArgList[cArg].hlstr = NULL;
            }
            else {
               if ( LastToken() != TOKEN_IDENT )
                  return WBERR_EXPECT_VARIABLE;
               npvar = WBLookupVariable( GetTokenText(), FALSE );
               if ( npvar == NULL )
                  return WBERR_VARUNDEFINED;
               if ( npvar->value.type == VT_USERTYPE ) {
                  lpTypeMem = (LPWBTYPEMEM)npvar->value.var.tlpVoid;
                  NextToken();
                  err = WBGetUserTypeData( lpTypeMem,    &npUserType, 
                                           &fMemberType, &lpData, &uSize );
                  if ( err != 0 )
                     return err;

                  if ( nparg->fArgType == ARGTYPE_INTEGER &&
                       fMemberType != VT_I2 ) 
                     return WBERR_TYPEMISMATCH;
                  else if ( nparg->fArgType == ARGTYPE_LONG &&
                       fMemberType != VT_I4 ) 
                     return WBERR_TYPEMISMATCH;
                        
                  ArgList[cArg].fWord = FALSE;
                  ArgList[cArg].arg.tDWord = (DWORD)(LPVOID)lpData;
                  ArgList[cArg].hlstr = NULL;
               }
               else {
                  switch ( npvar->value.type ) {
                     case VT_I4:
                        break;
                     case VT_STRING:
                        if ( (err = WBMakeLongVariant( &(npvar->value) )) != 0 )
                           return err;
                        break;
                     default:     
                        return WBERR_TYPEMISMATCH;
                  }
                  if ( nparg->fArgType == ARGTYPE_INTEGER ) 
                     npvar->value.var.tLong = (int)npvar->value.var.tLong;
                     
                  ArgList[cArg].fWord = FALSE;
                  ArgList[cArg].arg.tDWord = (DWORD)(LPVOID)(&(npvar->value.var.tLong));
                  ArgList[cArg].hlstr = NULL;
                  NextToken();
               }
            }
            break;
            
         case ARGTYPE_STRING:
            if ( (err = WBExpression( &vResult )) != 0 )
               return err;
               
            switch ( vResult.type ) {
               case VT_STRING:
                  break;
               case VT_I4:
               case VT_DATE:
                  if ( (err = WBVariantToString( &vResult )) != 0 )
                     return err;
                  break;
               default:
                  return WBERR_TYPEMISMATCH;
            }
            
            hlstr = vResult.var.tString;
            if ( fByVal || nparg->fByVal )
               lpsz = WBDerefZeroTermHlstr( hlstr );
            else
               lpsz = WBDerefHlstr( hlstr );
            ArgList[cArg].fWord = FALSE;
            ArgList[cArg].arg.tDWord = (DWORD)(LPVOID)lpsz;
            ArgList[cArg].hlstr = hlstr;
            break;
            
         case ARGTYPE_USER:
            if ( LastToken() != TOKEN_IDENT )
               return WBERR_EXPECT_VARIABLE;
                  
            npvar = WBLookupVariable( GetTokenText(), FALSE );
            if ( npvar == NULL )
               return WBERR_VARUNDEFINED;
            if ( npvar->value.type != VT_USERTYPE )
               return WBERR_TYPEMISMATCH;
            if ( fByVal ) 
               return WBERR_BYREFONLY;
                  
            lpTypeMem = (LPWBTYPEMEM)npvar->value.var.tlpVoid;
            NextToken();
            err = WBGetUserTypeData( lpTypeMem,    &npUserType, 
                                     &fMemberType, &lpData, &uSize );
            if ( err != 0 )
               return err;
               
            if ( npUserType != nparg->npUserType )
               return WBERR_TYPEMISMATCH;
                  
            ArgList[cArg].fWord = FALSE;
            ArgList[cArg].arg.tDWord = (DWORD)(LPVOID)lpData;
            ArgList[cArg].hlstr = NULL;
            break;
            
         case ARGTYPE_ANY:
            if ( fByVal || nparg->fByVal ) {
            
               if ( (err = WBExpression( &vResult )) != 0 )
                  return err;
                  
               switch ( vResult.type ) {
                  
                  case VT_I4:
                     ArgList[cArg].fWord = FALSE;
                     ArgList[cArg].arg.tDWord = (DWORD)vResult.var.tLong;
                     ArgList[cArg].hlstr = NULL;
                     break;
   
                  case VT_STRING:                  
                     hlstr = vResult.var.tString;
                     lpsz = WBDerefZeroTermHlstr( hlstr );
                     ArgList[cArg].fWord = FALSE;
                     ArgList[cArg].arg.tDWord = (DWORD)(LPVOID)lpsz;
                     ArgList[cArg].hlstr = hlstr;
                     break;
                     
                  default:
                     return WBERR_TYPEMISMATCH;
               }
            }
            else {
               if ( LastToken() != TOKEN_IDENT )
                  return WBERR_EXPECT_VARIABLE;
                  
               npvar = WBLookupVariable( GetTokenText(), FALSE );
               if ( npvar == NULL )
                  return WBERR_VARUNDEFINED;
                  
               switch ( npvar->value.type ) {
               
                  case VT_I4:
                     ArgList[cArg].fWord = FALSE;
                     ArgList[cArg].arg.tDWord = (DWORD)(LPVOID)(&(npvar->value.var.tLong));
                     ArgList[cArg].hlstr = NULL;
                     NextToken();
                     break;
                     
                  case VT_STRING:
                     hlstr = vResult.var.tString;
                     lpsz = WBDerefHlstr( hlstr );
                     ArgList[cArg].fWord = FALSE;
                     ArgList[cArg].arg.tDWord = (DWORD)(LPVOID)lpsz;
                     ArgList[cArg].hlstr = hlstr;
                     NextToken();
                     break;
                     
                  case VT_USERTYPE:
                     lpTypeMem = (LPWBTYPEMEM)npvar->value.var.tlpVoid;
                     NextToken();
                     err = WBGetUserTypeData( lpTypeMem,    &npUserType, 
                                              &fMemberType, &lpData, &uSize );
                     if ( err != 0 )
                        return err;
                     ArgList[cArg].fWord = FALSE;
                     ArgList[cArg].arg.tDWord = (DWORD)(LPVOID)lpData;
                     ArgList[cArg].hlstr = NULL;
                     break;
                     
                  default:
                     return WBERR_TYPEMISMATCH;
               }
            }
            break;
            
         default:
            assert( FALSE );
            return WBERR_INTERNAL_ERROR;
      }
      
      cArg++;
      nparg = nparg->npNext;
      
      tok = LastToken();
      if ( tok == TOKEN_COMMA ) {
         tok = NextToken();
         if ( tok == TOKEN_EOL )
            tok = NextToken();
      }
      else if ( tok != TOKEN_RPAREN && tok != TOKEN_EOL )
         return WBERR_SYNTAX;
   }
   
   if ( cArg < npfun->cArg )
      return WBERR_ARGTOOFEW;

   _asm {
         mov   wSPBefore,sp
   }
   
   for ( i = 0; i < cArg; i++ ) {
      if ( ArgList[i].fWord ) {
         wArg = ArgList[i].arg.tWord;
         _asm {
               push  wArg
         }
      }
      else {
         dwArg = ArgList[i].arg.tDWord;
         _asm {
               push  WORD PTR dwArg+2
               push  WORD PTR dwArg
          }
      }
   }
   
   _asm {
         call  lpfn
         mov   wSPAfter,sp
         mov   WORD PTR dwArg,ax
         mov   WORD PTR dwArg+2,dx
   }
   
   if ( wSPBefore != wSPAfter ) {
      _asm  mov   sp,WORD PTR wSPBefore
      return WBERR_BADCALLINGCONVENTION;
   }
   
   for ( i = 0; i < cArg; i++ ) {
      if ( ArgList[i].hlstr )
         WBDestroyHlstrIfTemp( ArgList[i].hlstr );
   }

   switch ( npfun->fReturnType ) {
      
      case ARGTYPE_VOID:
         lpResult->var.tLong = WB_TRUE;
         lpResult->type = VT_I4;
         break;
         
      case ARGTYPE_INTEGER:
         lpResult->var.tLong = (long)(int)LOWORD(dwArg);
         lpResult->type = VT_I4;
         break;
         
      case ARGTYPE_LONG:
         lpResult->var.tLong = (long)dwArg;
         lpResult->type = VT_I4;
         break;

      case ARGTYPE_STRING:
         lpsz = (LPSTR)dwArg;
         lpResult->var.tString = WBCreateTempHlstr( lpsz, (USHORT)strlen( lpsz ) );
         if ( lpResult->var.tString == NULL )
            return WBERR_STRINGSPACE;
         lpResult->type = VT_STRING;
         break;
                  
      default:
         return WBERR_TYPEMISMATCH;
   }
   
   return 0;
}
               
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR FAR PASCAL WBFnTruncString( int narg, LPVAR lparg, LPVAR lpResult )
{
   LPSTR lpsz1, lpsz2;
   USHORT usLen;
   
   if ( narg < 1 )
      return WBERR_ARGTOOFEW;
   else if ( narg > 1 )
      return WBERR_ARGTOOMANY;
   
   lpsz1 = WBLockHlstrLen( lparg[0].var.tString, &usLen );
   lpsz2 = memchr( lpsz1, 0, (size_t)usLen );
   if ( lpsz2 == NULL )
      lpsz2 = lpsz1 + usLen;
      
   usLen = (USHORT)(lpsz2 - lpsz1);
   lpResult->var.tString = WBCreateTempHlstr( lpsz1, usLen );
   WBUnlockHlstr( lparg[0].var.tString );
   
   if ( lpResult->var.tString == NULL )
      return WBERR_STRINGSPACE;
      
   lpResult->type = VT_STRING;
   return 0;
}

