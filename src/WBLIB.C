#define STRICT
#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include "wupbas.h"
#include "wbmem.h"
#include "wbfndos.h"
#include "wbtype.h"
#include "wbdllcal.h"

extern char NEAR g_szDLLPath[];

#define PROCTYPE_SUB       0
#define PROCTYPE_FUNCTION  1

ERR WBLoadLib( LPCSTR lpszFileName, LPLIBLINE FAR* lplpFirstLine );
LPVOID CALLBACK WBLibNextLine( LPSTR lpszBuffer, UINT cb );
void CALLBACK WBLibGotoLine( LPVOID lpLine );

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBLoadLibStatement( void )
{
   VARIANT vResult;
   LPSTR lpszFileName;
   NPWBHASHENTRY npHash;
   UINT fuDeclType;
   ERR err;
   char szName[MAXTOKEN];
   LPSTR lpch;

   if ( (err = WBExpression( &vResult )) != 0 )
      WBRuntimeError( err );
      
   if ( (err = WBVariantToString( &vResult )) != 0 )
      WBRuntimeError( err );

   lpszFileName = WBDerefZeroTermHlstr( vResult.var.tString );
   
   if ( (err = WBLoadProgram( lpszFileName )) != 0 )
      WBRuntimeError( err );

   lpszFileName = WBDerefZeroTermHlstr( vResult.var.tString );

   // Prefix name LIBMAIN with base file name   
   lpch = strrchr( lpszFileName, '\\' );
   if ( lpch != NULL )
      lpch++;
   else
      lpch = lpszFileName;
   strncpy( szName, lpch, sizeof(szName)-1 );
   szName[sizeof(szName)-1] = '\0';
   lpch = strchr( szName, '.' );
   if ( lpch != NULL )
      *lpch = NULL;
   strcat( szName, "LIBMAIN" );

   WBDestroyHlstrIfTemp( vResult.var.tString );   

   // Check for a Sub LibMain and execute it if it exists
   npHash.VoidPtr = WBLookupDeclaration( szName, &fuDeclType );
   if ( fuDeclType == DECL_USERFUNC ) {
      if ( (err = WBCallUserFunction( npHash.UserFunc, 0, NULL, &vResult )) != 0 )
         WBRuntimeError( err );
   }
   
   WBCheckEndOfLine();
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBLoadProgram( LPSTR lpszFileName )
{
   char szFileName[_MAX_PATH];
   char szFileName2[_MAX_PATH];
   NPWBLOADFILEENTRY nplfe;
   char szName[MAXTOKEN];
   NPWBHASHENTRY npHash;
   UINT fuDeclType;
   NPWBDECLARATION npDecl;
   NPWBUSERFUNC npUserFunc;
   UINT uSize;
   LPLIBLINE lpFirstLine;
   LPVOID lpLine;
   LPSTR lpch;
   TOKEN tok;
   int cArg;
   ERR err;
   BOOL fFileScope = TRUE;
   
   // If not directory specified, use the directory in from which the
   // DLL was loaded
   if ( strpbrk( lpszFileName, ":\\" ) == NULL ) {
      strcpy( szFileName, g_szDLLPath );
      strcat( szFileName, lpszFileName );
   }
   else {
      strcpy( szFileName, lpszFileName );
   }

   _fullpath( szFileName2, szFileName, sizeof(szFileName2) );
   
   if( (err = WBLoadLib( szFileName2, &lpFirstLine )) != 0 )
      return err;
      
   // Allocate new execution context
   if ( WBPushContext() == NULL )
      return WBERR_OUTOFMEMORY;
      
      
   nplfe = (NPWBLOADFILEENTRY)WBLocalAlloc( LPTR, 
                                            sizeof(WBLOADFILEENTRY) +
                                            strlen( szFileName2 ) );
   if ( nplfe == NULL )
      return WBERR_OUTOFMEMORY;
      
   _strlwr( strcpy( nplfe->szLibName, szFileName2 ) );
   nplfe->npNext = g_npTask->npLoadFileList;
   g_npTask->npLoadFileList = nplfe;
   
   // Set current file name
   g_npContext->npszLibName = nplfe->szLibName;
      
   // Install NextLine and GotoLine callback's for library
   g_npContext->lpfnNextLine = WBLibNextLine;
   g_npContext->lpfnGotoLine = WBLibGotoLine;


   // Build table of user functions and subroutines
            
   g_npContext->lpLibCurrent = lpFirstLine;
   
   SkipLine();
   lpLine = (LPVOID)g_npContext->lpLibCurrent;
   
   tok = NextToken();
         
   while (  tok != TOKEN_EOF ) {

      switch ( tok ) {

         case TOKEN_SUB:
         case TOKEN_FUNCTION:
            fFileScope = FALSE;         
      
            if ( (tok = NextToken()) != TOKEN_IDENT ) 
                  return WBERR_EXPECT_PROCNAME;
            
            strcpy( szName, GetTokenText() ); 
   
            // If Sub name is LibMain prefix the name with the
            // base file name of the program file to make it
            // unique.
            if ( strcmp ( szName, "LIBMAIN" ) == 0 ) {
               lpch = strrchr( szFileName, '\\' );
               if ( lpch != NULL )
                  lpch++;
               else
                  lpch = szFileName;
               strncpy( szName, lpch, sizeof(szName)-1 );
               lpch = strchr( szName, '.' );
               if ( lpch != NULL )
                  *lpch = NULL;
               szName[sizeof(szName)-1] = '\0';
               strcat( szName, GetTokenText() );
            }
   
            npHash.VoidPtr = WBLookupDeclaration( szName, &fuDeclType );
            if ( fuDeclType != DECL_NOTFOUND )   
               WBRuntimeError( WBERR_DUPLICATEDEFINITION );
               
            // Count parameter in formal parameter list
            cArg = 0;
            tok = NextToken();
            if ( tok != TOKEN_LPAREN )
               return WBERR_EXPECT_LPAREN;
               
            tok = NextToken();
            
            while ( tok != TOKEN_RPAREN && tok != TOKEN_EOL ) {
            
               if ( tok != TOKEN_IDENT ) 
                  return WBERR_EXPECT_PARAMETER;
                  
               cArg++;
               
               tok = NextToken();
               
               if ( tok == TOKEN_COMMA ) {
                  tok = NextToken();
                  if ( tok == TOKEN_EOL )
                     tok = NextToken();
               }
            }
            
            if ( tok != TOKEN_RPAREN )
               return WBERR_EXPECT_RPAREN;
            
            if ( cArg > MAXARGS )
               return WBERR_ARGTOOMANY;
                  
            tok = NextToken();
            if ( tok != TOKEN_EOL )
               return WBERR_EXPECT_EOL;
   
            uSize = sizeof(WBDECLARATION) + strlen(szName);
            
            // Adjust size for word alignment of extra data
            if ( uSize & 0x0001 )
               uSize++;
               
            npDecl = (NPWBDECLARATION)WBLocalAlloc( LPTR, uSize + sizeof(WBUSERFUNC) );
            if ( npDecl == NULL )
               return WBERR_OUTOFMEMORY;
               
            npUserFunc = (NPWBUSERFUNC)(((NPSTR)npDecl) + uSize);
            npDecl->npExtra.UserFunc = npUserFunc;
            npUserFunc->cArg = cArg;
            npUserFunc->lpLine = lpLine;
            
            // Record the name of the file for error reporting
            npUserFunc->npszLibName = nplfe->szLibName;
            
            // Point back from extra data to name field
            npUserFunc->lpszName = npDecl->szName;
            
            WBAddDeclaration( npDecl, szName, DECL_USERFUNC );
            break;
            
         case TOKEN_TYPE:
            if ( !fFileScope )
               return WBERR_ILLEGALPLACEMENT;
               
            tok = NextToken();
            if ( (err = WBTypeStatement()) != 0 )
               return err;
            break;
            
         case TOKEN_GLOBAL:
            if ( !fFileScope )
               return WBERR_ILLEGALPLACEMENT;
               
            tok = NextToken();
            if ( (err = WBGlobalStatement()) != 0 )
               return err;
            break;
         
         case TOKEN_DECLARE:
            if ( !fFileScope )
               return WBERR_ILLEGALPLACEMENT;
               
            tok = NextToken();
            if ( (err = WBDeclareStatement()) != 0 )
               return err;
            break;
      }
      
      SkipLine();
      lpLine = g_npContext->lpLibCurrent;
      tok = NextToken();
   }

   WBPopContext();
   
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBCallUserFunction( NPWBUSERFUNC npUserFunc, int narg, LPVAR lparg, LPVAR lpResult )
{
   NPVARDEF npvar;
   NPVARDEF npvarArg[MAXARGS];
   TOKEN tok, endTok1, endTok2;
   UINT fuProcType;
   NPVARDEF NEAR* npvtLocal;
   int i;
   ERR err;
   
   if ( narg < npUserFunc->cArg )
      return WBERR_ARGTOOFEW;
   else if ( narg > npUserFunc->cArg )
      return WBERR_ARGTOOMANY;
   
   if ( WBPushContext() == NULL )
      return WBERR_OUTOFMEMORY;

   // Mark new context with library file name      
   g_npContext->npszLibName = npUserFunc->npszLibName;
      
   // Install NextLine and GotoLine callback's for library
   g_npContext->lpfnNextLine = WBLibNextLine;
   g_npContext->lpfnGotoLine = WBLibGotoLine;
   
   WBGotoLine( npUserFunc->lpLine );
   SkipLine();
   tok = NextToken();

   switch ( tok ) {
   
      case TOKEN_SUB:   
         fuProcType = PROCTYPE_SUB;
         endTok1 = TOKEN_ENDSUB;
         endTok2 = TOKEN_EXITSUB;
         break;
         
      case TOKEN_FUNCTION:
         fuProcType = PROCTYPE_FUNCTION;
         endTok1 = TOKEN_ENDFUNCTION;
         endTok2 = TOKEN_EXITFUNCTION;
         break;
         
      default:  
         assert( FALSE );
         return WBERR_INTERNAL_ERROR;
         
   }
   
   // Skip function name
   tok = NextToken();
   
   tok = NextToken();
   if ( tok != TOKEN_LPAREN )
      return WBERR_EXPECT_LPAREN;
      
   tok = NextToken();

   for ( i = 0; 
         i < narg && tok != TOKEN_RPAREN && tok != TOKEN_EOL;
         i++ ) {
         
      if ( tok != TOKEN_IDENT ) 
         return WBERR_SYNTAX;
      
      npvar = WBInsertLocalVariable( GetTokenText() );
      npvarArg[i] = npvar;
      
      switch ( lparg[i].type ) {
      
         case VT_I4:
         case VT_DATE:
            npvar->value = lparg[i];
            npvar->fuAttribs = VA_VARIANT;
            break;
            
         case VT_USERTYPE:
            npvar->value = lparg[i];
            npvar->fuAttribs = VA_USERTYPE;
            break;
            
         case VT_STRING:
            npvar->value.var.tString = WBCreateHlstrFromTemp( lparg[i].var.tString );
            if ( npvar->value.var.tString == NULL )
               return WBERR_STRINGSPACE;
            npvar->value.type = VT_STRING;
            npvar->fuAttribs = VA_VARIANT;
            break;
            
         default:
            if ( lparg[i].type > VT_OBJECT ) {
               npvar->value = lparg[i];
               npvar->fuAttribs = VA_VARIANT;
               WBWireObject( npvar->value.var.tCtl );
            }
            else 
               return WBERR_TYPEMISMATCH;
      }
      
      tok = NextToken();
      if ( tok == TOKEN_COMMA ) {
         tok = NextToken();
         if ( tok == TOKEN_EOL )
            tok = NextToken();
      }
      else if ( tok != TOKEN_RPAREN && tok != TOKEN_EOL ) 
         return WBERR_SYNTAX;
   }
   
   if ( tok != TOKEN_RPAREN ) 
      return WBERR_EXPECT_RPAREN;
      
   assert( i == narg );
      
   tok = NextToken();
   
   if ( tok != TOKEN_EOL ) 
      return WBERR_EXPECT_EOL;

   // Insert the Function name into the local var table as
   // a VT_EMPTY variant.
   if ( fuProcType == PROCTYPE_FUNCTION ) {
      npvar = WBInsertLocalVariable( npUserFunc->lpszName );
      npvar->fuAttribs = VA_VARIANT;
   }
      
   tok = NextToken();
   
   while ( (tok = LastToken()) != endTok1 && tok != endTok2 && tok != TOKEN_EOF) {
      WBStatement();
   }

   if ( tok == TOKEN_EOF ) {
      WBRuntimeError( WBERR_UNEXPECTED_EOF );
   }
   else {
      err = 0;
         
      if ( fuProcType == PROCTYPE_FUNCTION ) {
         npvar = WBLookupVariable( npUserFunc->lpszName, FALSE );
         assert ( npvar != NULL );
         if ( npvar->value.type == VT_EMPTY )
            WBRuntimeError( WBERR_NO_RETURN_VALUE );
         else {
            if (npvar->value.type == VT_STRING ) {
               lpResult->var.tString = WBCreateTempHlstrFromTemp( 
                                                   npvar->value.var.tString );
               lpResult->type = VT_STRING;
            }
            else if (npvar->value.type > VT_OBJECT ) {
               *lpResult = npvar->value;
               // Increment ref count to returned object so
               // it doesn't get destroyed along with the
               // local variable table.
               WBObjectRefIncrement( lpResult->var.tCtl );
            }
            else {
               *lpResult = npvar->value;
            }
         }
      }
      else {
         lpResult->var.tLong = WB_TRUE;
         lpResult->type = VT_I4;
      }
   }
   
   // Variables of type VT_USERTYPE are always passed by reference.
   // To prevent that the memory associated with the VT_TYPE is
   // destroyed along with the local var table we will change
   // the variant type from VT_USERTYPE to VT_EMPTY.   
   for ( i = 0; i < narg; i++ ) {
      if ( npvarArg[i]->value.type == VT_USERTYPE )
         npvarArg[i]->value.type = VT_EMPTY;
   }
   
   npvtLocal = g_npContext->npvtLocal;
   
   // Restore previous context
   WBPopContext();
   
   // Destroy local variable table if it exists
   if ( npvtLocal != NULL ) {      
      WBDestroyVarTable( npvtLocal );
      WBLocalFree( (HLOCAL)npvtLocal );
   }
   
   if ( lpResult->type > VT_OBJECT ) {
      // Undo the temporary ref count increment
      WBObjectRefDecrement( lpResult->var.tCtl );
   }
   
   return err;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
LPVOID CALLBACK WBLibNextLine( LPSTR lpszBuffer, UINT cb )
{  
   LPVOID lpLine;
   LPSTR lpszText;
   USHORT usSize;
   
   // If current line pointer points to head node we have reached
   // the end of the list and we return NULL to indicate EOF
   if ( g_npContext->lpLibCurrent == g_npTask->lpLibHead )
      return NULL;

   // Otherwise copy up to cb bytes (including null terminator) into the 
   // buffer provided.
   lpszText = g_npContext->lpLibCurrent->szText;
   usSize = min( (USHORT)strlen( lpszText ), (USHORT)cb-1  );
   if ( usSize > 0 )
      _fmemcpy( lpszBuffer, lpszText, usSize );
   lpszBuffer[usSize] = '\0';

   g_npContext->uLineNum = g_npContext->lpLibCurrent->uLineNum;
      
   // Advance the current line pointer to the next line.      
   lpLine = (LPVOID)g_npContext->lpLibCurrent;
   g_npContext->lpLibCurrent = g_npContext->lpLibCurrent->lpNext;
   
   return lpLine;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void CALLBACK WBLibGotoLine( LPVOID lpLine )
{
   // Reset the current line pointer to the one provided
   g_npContext->lpLibCurrent = (LPLIBLINE)lpLine;
}   

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBLoadLib( LPCSTR lpszFileName, LPLIBLINE FAR* lplpFirstLine )
{
   FILE *pfile;
   HMEM hmem;
   LPLIBLINE lpHead;
   LPLIBLINE lpNew;
   char szLine[256];
   LPSTR lpch1, lpch2;
   UINT uLineNum = 0;
   ERR err = 0;
   

   *lplpFirstLine = NULL;
      
   pfile = fopen( lpszFileName, "r" );
   if ( pfile == NULL )
      return WBERR_FILENOTFOUND;
      
   if ( g_npTask->lpLibHead == NULL ) {
      // Allocate head node for doubly linked list   
      hmem = WBSubAlloc( g_npTask->hLibPool, LMEM_MOVEABLE, sizeof( LIBLINE ));
      if ( hmem == NULL )
         return WBERR_OUTOFMEMORY;
         
      lpHead = (LPLIBLINE)WBSubLock( hmem );
      assert( lpHead != NULL );
      lpHead->hmem = hmem;
         
      // Initialize next and prev pointer to point to the head node itself
      lpHead->lpNext = lpHead->lpPrev = lpHead;
      g_npTask->lpLibHead = lpHead;
   }
   else {
      lpHead = g_npTask->lpLibHead;
   }

   // Read file and append to list
   while ( fgets( szLine, sizeof( szLine ), pfile ) != NULL ) {

      uLineNum++;
                     
      // Trim leading blanks
      lpch1 = szLine;
      while ( isspace( *lpch1 ) )
         lpch1++;

      // Discard comments
      if ( *lpch1 == '\'' || *lpch1 == ';' )
         *lpch1 = '\0';
  
      // Trim trailing white space 
      lpch2 = lpch1 + strlen( lpch1 ) - 1;
      while ( lpch2 >= lpch1 && (isspace( *lpch2 ) || iscntrl( *lpch2 )) )
         lpch2--;
      *(lpch2+1) = '\0';
      
      if ( *lpch1 == '\0' )
         continue;
      
      // Allocate memory for new node
      hmem = WBSubAlloc( g_npTask->hLibPool, LMEM_MOVEABLE,
                       (USHORT)sizeof( LIBLINE ) + (USHORT)strlen( lpch1 ) );
      if ( hmem == NULL ) {
         err = WBERR_OUTOFMEMORY;
         break;
      }
                               
      lpNew = (LPLIBLINE)WBSubLock( hmem );

      // Copy test into new node
      strcpy( lpNew->szText, lpch1 );
      lpNew->hmem = hmem;
      lpNew->uLineNum = uLineNum;

      // Insert at end of list
      lpNew->lpNext = lpHead;
      lpNew->lpPrev = lpHead->lpPrev;
      lpHead->lpPrev->lpNext = lpNew;
      lpHead->lpPrev = lpNew;
      
      if ( *lplpFirstLine == NULL )
         *lplpFirstLine = lpNew;
      
   }
   
   fclose( pfile );
   
   return err;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBDestroyLoadLibList( void )
{
   NPWBLOADFILEENTRY nplfe, nplfe2;
   
   nplfe = g_npTask->npLoadFileList;
   
   while ( nplfe != NULL ) {
      nplfe2 = nplfe->npNext;
      WBLocalFree( (HLOCAL)nplfe );
      nplfe = nplfe2;
   }
}
   
// Don't bother freeing individual lines in release version. In the debug
// version we want to free individual lines so that memory leaks can be
// reported.
#ifdef _DEBUG
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBDestroyLib( LPLIBLINE lpHead )
{
   LPLIBLINE lpLine, lpTemp;
   HMEM hmem;

   lpLine = lpHead->lpNext;
   
   while ( lpLine != lpHead ) {
      lpTemp = lpLine->lpNext;
      hmem = lpLine->hmem;
      WBSubUnlock( hmem );
      WBSubFree( hmem );
      lpLine = lpTemp;
   }

   // Free head node   
   hmem = lpHead->hmem;
   WBSubUnlock( hmem );
   WBSubFree( hmem );
}
#endif
