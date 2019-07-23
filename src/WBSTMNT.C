#define STRICT
#include <windows.h>
#include <string.h>
#include <ctype.h> 
#include <time.h>
#include <limits.h>
#include <assert.h>
#include "wupbas.h"
#include "wbfunc.h"
#include "wbfileio.h"
#include "wbdllcal.h"
#include "wbtype.h"

void WBIfStatement( void );
void WBSkipIfStatement( void );
void WBWhileStatement( void );
void WBDoStatement( void );
void WBForStatement( void );
void WBOptionStatement( void );
void WBRunStatementBlock( WORD wTerm );
void WBSkipStatementBlock( WORD wTerm );
void WBSimpleAssigment( void );
void WBConstStatement( void );
void WBDimStatement( void );

void WBSubIntrinsicFunction( NPFUNCTIONINFO npIntFunc );
void WBSubUserDefinedFunction( NPWBUSERFUNC npUserFunc );
void WBSubDeclaredFunction( NPWBDECLARE npDeclare );

void WBUserTypeAssignment( LPWBTYPEMEM lpTypeMem );
void WBObjectOperation( HCTL hctl );
void WBExecCommand( void );

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBStatement( void )
{
   TOKEN tok;
   NPWBHASHENTRY npHash;
   UINT fuDeclType;
   NPVARDEF npvar;
   
   tok = LastToken();
   
   switch ( tok ) {
         
      case TOKEN_IDENT:
         if ( PeekToken() == TOKEN_EQU ) {
            WBSimpleAssigment();
            break;
         }
         
         npHash.VoidPtr = WBLookupDeclaration( GetTokenText(), &fuDeclType );
         switch ( fuDeclType ) {
            case DECL_INTFUNC:
               WBSubIntrinsicFunction( npHash.IntFunc->npIntFunc );
               break;
            case DECL_USERFUNC:
               WBSubUserDefinedFunction( npHash.UserFunc );
               break;
            case DECL_DECLARE:
               WBSubDeclaredFunction( npHash.Declare );
               break;
            case DECL_USERTYPE:
               WBRuntimeError( WBERR_EXPECT_PROCNAME );
               break;
               
            default:
               npvar = WBLookupVariable( GetTokenText(), FALSE );
               if ( npvar == NULL ) 
                  WBExecCommand();
               else if ( npvar->value.type == VT_USERTYPE ) 
                  WBUserTypeAssignment( (LPWBTYPEMEM)npvar->value.var.tlpVoid );
               else if ( npvar->value.type > VT_OBJECT ) 
                  WBObjectOperation( npvar->value.var.tCtl );
               else
                  WBRuntimeError( WBERR_SYNTAX );
         }
         break;
            
      case TOKEN_IF:
         NextToken();
         WBIfStatement();
         break;
         
      case TOKEN_WHILE:
         NextToken();
         WBWhileStatement();
         break;
         
      case TOKEN_DO:
         NextToken();
         WBDoStatement();
         break;
         
      case TOKEN_EXITDO:
         if ( g_npTask->npContext->nDoLoopLevel > 0 ) {
            g_npTask->npContext->fExitBlock = TRUE;
            NextToken();
            WBCheckEndOfLine();
         }
         else 
            WBRuntimeError( WBERR_EXITDOWITHOUTDO );
         break;
         
      case TOKEN_FOR:
         NextToken();
         WBForStatement();
         break;
         
      case TOKEN_EXITFOR:
         if ( g_npTask->npContext->nForNextLevel > 0 ) {
            g_npTask->npContext->fExitBlock = TRUE;
            NextToken();
            WBCheckEndOfLine();
         }
         else 
            WBRuntimeError( WBERR_EXITFORWITHOUTFOR );
         break;

      case TOKEN_EOL:
         NextToken();
         break;
               
      case TOKEN_LOADLIB:
         NextToken();
         WBLoadLibStatement();
         break;
         
      case TOKEN_LOADDLL:
         NextToken();
         WBLoadDLLStatement();
         break;
         
      case TOKEN_END:
         NextToken();
         WBCheckEndOfLine();
         WBStopExecution( 0 );
         break;

      case TOKEN_STOP:
         NextToken();
         WBCheckEndOfLine();
         WBStopExecution( WBERR_STOP );
         break;

      case TOKEN_EXIT:
         NextToken();
         WBCheckEndOfLine();
         WBStopExecution( WBERR_EXIT );
         break;

      case TOKEN_STRINGLIT:
         WBExecCommand();
         break;

      case TOKEN_OPTION:
         NextToken();
         WBOptionStatement();
         break;

      case TOKEN_DIM:
      case TOKEN_LOCAL:
         NextToken();
         WBDimStatement();
         break;

      case TOKEN_CONST:
         NextToken();
         WBConstStatement();
         break;

      case TOKEN_ELSE:
         WBRuntimeError( WBERR_ELSEWITHOUTIF );
         break;
         
      case TOKEN_ELSEIF:
         WBRuntimeError( WBERR_ELSEIFWITHOUTIF );
         break;
         
      case TOKEN_ENDIF:
         WBRuntimeError( WBERR_ENDIFWITHOUTIF);
         break;
         
      case TOKEN_LOOP:
         WBRuntimeError( WBERR_LOOPWITHOUTDO );
         break;
         
      case TOKEN_NEXT:
         WBRuntimeError( WBERR_NEXTWITHOUTFOR );
         break;
         
      case TOKEN_WEND:
         WBRuntimeError( WBERR_WENDWITHOUTWHILE );
         break;
         
      case TOKEN_ENDSUB:
         WBRuntimeError( WBERR_ENDSUBWITHOUTSUB );
         break;
         
      case TOKEN_EXITSUB:
         WBRuntimeError( WBERR_EXITSUBWITHOUTSUB );
         break;
         
      case TOKEN_ENDFUNCTION:
         WBRuntimeError( WBERR_ENDFUNCTIONWITHOUTFUNCTION );
         break;
         
      case TOKEN_EXITFUNCTION:
         WBRuntimeError( WBERR_EXITFUNCTIONWITHOUTFUNCTION );
         break;
         
      case TOKEN_ENDTYPE:
         WBRuntimeError( WBERR_ENDTYPEWITHOUTTYPE );
         break;
         
      default:
         if ( LastTokenClass() == TC_KEYWORD ) {
            tok = PeekToken();
            if ( tok == TOKEN_EQU || tok == TOKEN_PERIOD )
               WBRuntimeError( WBERR_RESERVEDWORD );
         }
         WBExecCommand();
   }
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBRunStatementBlock( WORD wTerm )
{
   TOKEN tok;
   
   g_npTask->npContext->fExitBlock = FALSE;
   
   tok = LastToken();
   
   while ( tok != TOKEN_EOF && !IsTermToken( tok, wTerm ) ) {
      WBStatement();
         
      if ( g_npTask->npContext->fExitBlock ) {
         WBSkipStatementBlock( wTerm );
         break;
      }
      tok = LastToken();
   }
   
   if ( tok == TOKEN_EOF ) 
      WBRuntimeError( WBERR_UNEXPECTED_EOF );
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBSkipStatementBlock( WORD wTerm )
{
   TOKEN tok;
   
   tok = LastToken();
   
   while ( !IsTermToken( tok, wTerm ) && tok != TOKEN_EOF ) {
   
      switch ( tok ) {
      
         case TOKEN_IF:
             WBSkipIfStatement();
             break;
         
         case TOKEN_FOR:
            SkipLine();
            NextToken();
            WBSkipStatementBlock( TERM_NEXT );
            tok = NextToken();
            if ( tok == TOKEN_IDENT ) 
               NextToken();
            WBCheckEndOfLine();
            break;
               
         case TOKEN_WHILE:
            SkipLine();
            NextToken();
            WBSkipStatementBlock( TERM_WEND );
            NextToken();
            WBCheckEndOfLine();
            break;
            
         case TOKEN_DO:
            SkipLine();
            NextToken();
            WBSkipStatementBlock( TERM_LOOP );
            SkipLine();
            NextToken();
            break;
            
         default:
            SkipLine();
            NextToken();
      }
      
      tok = LastToken();
   }
   
   if ( tok == TOKEN_EOF ) 
      WBRuntimeError( WBERR_UNEXPECTED_EOF );
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBIfStatement( void )
{
   TOKEN tok;
   VARIANT vResult;
   BOOL fIsElseIf;
   ERR err;
   
   do {   
   
      fIsElseIf = FALSE;

      if ( ( err = WBExpression( &vResult ) ) != 0 )
         WBRuntimeError( err );
         
      if ( (err = WBMakeLongVariant( &vResult )) != 0 ) 
         WBRuntimeError( err );
         
      tok = LastToken();
      if ( tok != TOKEN_THEN ) {
         if ( tok == TOKEN_EOL )
            WBRuntimeError( WBERR_MISSING_THEN );
         else
            WBRuntimeError( WBERR_INVALID_CONDITION );
      }
      
      tok = NextToken();
      WBCheckEndOfLine();

      if ( vResult.var.tLong != 0 ) {
      
         WBRunStatementBlock( TERM_ENDIF | TERM_ELSE | TERM_ELSEIF );
            
         tok = LastToken();
         if ( tok == TOKEN_ELSEIF ) {
            while ( tok == TOKEN_ELSEIF ) {
               SkipLine();
               tok = NextToken();
               WBSkipStatementBlock( TERM_ENDIF | TERM_ELSE | TERM_ELSEIF );
          tok = LastToken();
            }
         }
         if ( tok == TOKEN_ENDIF ) {
            NextToken();
            WBCheckEndOfLine();
         }
         else if ( tok == TOKEN_ELSE ) {
            NextToken();
            WBCheckEndOfLine();
            WBSkipStatementBlock( TERM_ENDIF | TERM_ELSE | TERM_ELSEIF );
            tok = LastToken();
            if ( TOKEN_ELSE == tok ) 
               WBRuntimeError( WBERR_ELSEWITHOUTIF );
            else if ( TOKEN_ELSEIF == tok ) 
               WBRuntimeError( WBERR_ELSEIFWITHOUTIF );
            tok = NextToken();
            WBCheckEndOfLine();
         }
      }
      else {
      
         WBSkipStatementBlock( TERM_ENDIF | TERM_ELSE | TERM_ELSEIF );
         tok = LastToken();
   
         if ( tok == TOKEN_ENDIF ) {
            tok = NextToken();
            WBCheckEndOfLine();
         }
         else if ( tok == TOKEN_ELSE ) {
         
            tok = NextToken();
            WBCheckEndOfLine();
            
            WBRunStatementBlock( TERM_ENDIF | TERM_ELSE | TERM_ELSEIF );
            
            tok = LastToken();
            if ( TOKEN_ELSE == tok ) 
               WBRuntimeError( WBERR_ELSEWITHOUTIF );
            else if ( TOKEN_ELSEIF == tok ) 
               WBRuntimeError( WBERR_ELSEIFWITHOUTIF );
               
            tok = NextToken();
            WBCheckEndOfLine();
         }
         else if ( tok == TOKEN_ELSEIF ) {
            NextToken();
            fIsElseIf = TRUE;
         }
      }
   
   } while ( fIsElseIf );
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBSkipIfStatement( void )
{
   TOKEN tok;

   do {
   
      // Skip condition. End symbols are THEN (good) or end-of-line (bad)
      tok = NextToken();   
      while ( tok != TOKEN_THEN && tok != TOKEN_EOL ) {
         tok = NextToken();
      }
   
      // Check for end-of-line token
      if ( tok == TOKEN_EOL ) 
         WBRuntimeError( WBERR_MISSING_THEN );
          
      // Next symbol must be end-of-line
      tok = NextToken();
      WBCheckEndOfLine();
      
      WBSkipStatementBlock( TERM_ENDIF | TERM_ELSE | TERM_ELSEIF );
         
      tok = LastToken();
      
   } while ( tok == TOKEN_ELSEIF );
   
   if ( tok == TOKEN_ELSE ) {
      
      tok = NextToken();
      WBCheckEndOfLine();
      WBSkipStatementBlock( TERM_ENDIF );
      
      tok = NextToken();
      WBCheckEndOfLine();
   }
   else if ( tok == TOKEN_ENDIF ) {
      tok = NextToken();
      WBCheckEndOfLine();
   }
   else {
      assert( FALSE );
      WBRuntimeError( WBERR_INTERNAL_ERROR );
   }
}   

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBDoStatement( void )
{
   TOKEN tok;
   VARIANT vResult;
   LPVOID lpLine;
   BOOL fHaveDoCond;
   BOOL fDoCond;
   ERR err;

   g_npTask->npContext->nDoLoopLevel++;
   lpLine = GetCodePointer();
   tok = LastToken();
   
   for ( ;; ) {
   
      if ( tok == TOKEN_WHILE ) {
         fHaveDoCond = TRUE;
         NextToken();
         if ( (err = WBExpression( &vResult )) != 0 )
            WBRuntimeError( err );
         if ( (err = WBMakeLongVariant( &vResult )) != 0 ) 
            WBRuntimeError( err );
         if ( vResult.var.tLong == 0 ) 
            fDoCond = FALSE;
         else 
            fDoCond = TRUE;
         WBCheckEndOfLine();
      }
      else if ( tok == TOKEN_UNTIL ) {
         fHaveDoCond = TRUE;
         NextToken();
         if ( (err = WBExpression( &vResult )) != 0 )
            WBRuntimeError( err );
         if ( (err = WBMakeLongVariant( &vResult )) != 0 ) 
            WBRuntimeError( err );
         if ( vResult.var.tLong == 0 ) 
            fDoCond = TRUE;
         else 
            fDoCond = FALSE;
         WBCheckEndOfLine();
      }
      else if ( tok == TOKEN_EOL ) {
         fDoCond = TRUE;
         fHaveDoCond = FALSE;
         NextToken();
      }
      else {
         WBRuntimeError( WBERR_EXPECT_WHILE_UNTIL_EOL );
         return; // prevent C4701 on variables fHaveDoCond and fDoCond 
      }
      
      if ( !fDoCond ) {
         WBSkipStatementBlock( TERM_LOOP );
         tok = NextToken();
         WBCheckEndOfLine();
         break;
      }
      
      WBRunStatementBlock( TERM_LOOP );
      if ( g_npTask->npContext->fExitBlock ) {
         SkipLine();
         NextToken();
         break;
      }
      
      tok = NextToken();

      if ( tok == TOKEN_WHILE ) {
         if ( fHaveDoCond )
            WBRuntimeError( WBERR_LOOPWITHOUTDO ); 
         NextToken();
         if ( (err = WBExpression( &vResult )) != 0 )
            WBRuntimeError( err );
         if ( (err = WBMakeLongVariant( &vResult )) != 0 ) 
            WBRuntimeError( err );
         WBCheckEndOfLine();
         if ( vResult.var.tLong == 0 )
            break;
      }
      else if ( tok == TOKEN_UNTIL ) {
         if ( fHaveDoCond )
            WBRuntimeError( WBERR_LOOPWITHOUTDO ); 
         NextToken();
         if ( (err = WBExpression( &vResult )) != 0 )
            WBRuntimeError( err );
         if ( (err = WBMakeLongVariant( &vResult )) != 0 ) 
            WBRuntimeError( err );
         WBCheckEndOfLine();
         if ( vResult.var.tLong != 0 )
            break;
      }
      else if ( tok != TOKEN_EOL ) {
         WBRuntimeError( WBERR_EXPECT_WHILE_UNTIL_EOL );
      }
      
      WBGotoLine( lpLine );
      SkipLine();
      NextToken();   // Bump past DO token
      tok = NextToken();
   }
   
   g_npTask->npContext->nDoLoopLevel--;
   g_npTask->npContext->fExitBlock = FALSE;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBWhileStatement( void )
{
   VARIANT vResult;
   LPVOID lpLine;
   ERR err;

   lpLine = GetCodePointer();

   for ( ;; ) {      

      if ( ( err = WBExpression( &vResult ) ) != 0 )
         WBRuntimeError( err );
      if ( (err = WBMakeLongVariant( &vResult )) != 0 ) 
         WBRuntimeError( err );
            
      WBCheckEndOfLine();
      
      if ( vResult.var.tLong != 0 ) {
      
         WBRunStatementBlock( TERM_WEND );
            
         NextToken();
         WBCheckEndOfLine();
         WBGotoLine( lpLine );
         SkipLine();
         NextToken();   // Skip WHILE token
         NextToken();
      }
      else {
         WBSkipStatementBlock( TERM_WEND );
         NextToken();
         WBCheckEndOfLine();
         break;
      }
   }
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBForStatement( void )
{
   TOKEN tok;
   char szIndexVarName[MAXTOKEN];
   VARIANT vResult;
   NPVARDEF npvar;
   long lStop;
   long lStep = 1;
   LPVOID lpLine;
   ERR err;

   if ( LastToken() != TOKEN_IDENT )
      WBRuntimeError( WBERR_EXPECT_VARIABLE );
      
   strcpy( szIndexVarName, GetTokenText() );
   
   if ( NextToken() != TOKEN_EQU )
      WBRuntimeError( WBERR_EXPECT_EQU );
   NextToken();

   if ( (err = WBExpression( &vResult )) != 0 )
      WBRuntimeError( err );
   if ( (err = WBMakeLongVariant( &vResult )) != 0 ) 
      WBRuntimeError( err );
      
   npvar = WBLookupVariable( szIndexVarName, TRUE );
   if ( !(npvar->fuAttribs & VA_VARIANT) )
      WBRuntimeError( WBERR_TYPEMISMATCH );
   
   WBSetVariable( npvar, &vResult );
      
   if ( LastToken() != TOKEN_TO )
      WBRuntimeError( WBERR_EXPECT_TO );
      
   NextToken();
  
   if ( (err = WBExpression( &vResult )) != 0 )
      WBRuntimeError( err );
   if ( (err = WBMakeLongVariant( &vResult )) != 0 ) 
      WBRuntimeError( err );
      
   lStop = vResult.var.tLong;

   tok = LastToken();
      
   if ( tok == TOKEN_STEP ) {
      NextToken();
      if ( (err = WBExpression( &vResult )) != 0 )
         WBRuntimeError( err );
      if ( (err = WBMakeLongVariant( &vResult )) != 0 ) 
         WBRuntimeError( err );
      lStep = vResult.var.tLong;
   }
   else if ( tok != TOKEN_EOL ) 
      WBRuntimeError( WBERR_EXPECT_STEP_OR_EOL );

   NextToken();
       
   g_npTask->npContext->nForNextLevel++;
   lpLine = GetCodePointer();
   
   for ( ;; ) {
   
      if ( lStep >= 0 ) {
         if ( npvar->value.var.tLong > lStop ) {
            WBSkipStatementBlock( TERM_NEXT );
            break;
         }
      }
      else {
         if ( npvar->value.var.tLong < lStop ) {
            WBSkipStatementBlock( TERM_NEXT );
            break;
         }
      }

      WBRunStatementBlock( TERM_NEXT );
      if ( g_npTask->npContext->fExitBlock )
         break;
      
      if ( npvar->value.type != VT_I4 )
         WBRuntimeError( WBERR_TYPEMISMATCH );
      
      npvar->value.var.tLong += lStep;
      WBGotoLine( lpLine );      
      SkipLine();
      NextToken();
   }                                   
   
   tok = NextToken();
   
   if ( tok == TOKEN_IDENT ) {
      if ( strcmp( GetTokenText(), szIndexVarName ) != 0 ) 
         WBRuntimeError( WBERR_NEXTWITHOUTFOR );
      NextToken();
   }

   WBCheckEndOfLine();
   
   g_npTask->npContext->nForNextLevel--;
   g_npTask->npContext->fExitBlock = FALSE;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBGlobalStatement( void )
{
   TOKEN tok;
   NPVARDEF npvar;
   VARIANT vResult;
   HLSTR hlstr;
   NPWBHASHENTRY npHash;
   UINT fuDeclType;
   LPWBTYPEMEM lpTypeMem;
   ERR err;
   
   tok = LastToken();
   
   if ( tok == TOKEN_CONST ) {
   
      tok = NextToken();
      if ( tok == TOKEN_EOL )
         return WBERR_EXPECT_IDENT;

      while ( tok != TOKEN_EOL ) {
      
         if ( tok != TOKEN_IDENT )
            return WBERR_EXPECT_IDENT;
   
         WBLookupDeclaration( GetTokenText(), &fuDeclType );
         if ( fuDeclType != DECL_NOTFOUND )
            return WBERR_DUPLICATEDEFINITION;
            
         npvar = WBInsertGlobalVariable( GetTokenText() );
         npvar->fuAttribs = VA_CONST | VA_VARIANT;
            
         if ( NextToken() != TOKEN_EQU )
            return WBERR_EXPECT_EQU;
            
         NextToken();

         if ( (err = WBExpression( &vResult )) != 0 )
            return err;
            
         tok = LastToken();
            
         switch ( vResult.type ) {
         
            case VT_I4:
               npvar->value.var.tLong = vResult.var.tLong;
               npvar->value.type = VT_I4;
               break;
               
            case VT_STRING:
               // Create permanent copy of string
               hlstr = WBCreateHlstrFromTemp( vResult.var.tString );
               if ( hlstr == NULL )
                  return WBERR_STRINGSPACE;
               WBDestroyHlstrIfTemp( vResult.var.tString );
               npvar->value.var.tString = hlstr;
               npvar->value.type = VT_STRING;
               break;
               
            default:
               return WBERR_TYPEMISMATCH;
         }

         if ( tok != TOKEN_EOL ) {
            if ( tok == TOKEN_COMMA ) {
               tok = NextToken();
               if (tok == TOKEN_EOL)
                  tok = NextToken();
            }
            else
               return WBERR_EXPECTCOMMAOREOL;
         }
      }
   }
   else {
   
      if ( tok == TOKEN_EOL )
         return WBERR_EXPECT_IDENT;
         
      while ( tok != TOKEN_EOL ) {
      
         if ( tok != TOKEN_IDENT )
            return WBERR_EXPECT_IDENT;
         
         WBLookupDeclaration( GetTokenText(), &fuDeclType );
         if ( fuDeclType != DECL_NOTFOUND )
            return WBERR_DUPLICATEDEFINITION;
            
         npvar = WBInsertGlobalVariable( GetTokenText() );
         npvar->fuAttribs = VA_VARIANT;
         
         tok = NextToken();   
         
         if ( tok == TOKEN_AS ) {
         
            tok = NextToken();
            
            switch ( tok ) {
            
               case TOKEN_IDENT:
                  npHash.VoidPtr = WBLookupDeclaration( GetTokenText(), &fuDeclType );
                  if ( fuDeclType != DECL_USERTYPE )
                     return WBERR_UNDEFINED_TYPE;
         
                  lpTypeMem = WBAllocTypeMem( npHash.UserType, TRUE );
                  npvar->fuAttribs = VA_USERTYPE;
                  npvar->value.type = VT_USERTYPE;
                  npvar->value.var.tlpVoid = lpTypeMem;
                  break;
                  
               case TOKEN_VARIANT:
                  npvar->value.type = VT_I4;
                  npvar->value.var.tLong = 0L;
                  break;
               
               default:
                  return WBERR_EXPECT_TYPENAME;
            }
            
            tok = NextToken();
         }
         
         if ( tok != TOKEN_EOL ) {
            if ( tok == TOKEN_COMMA ) {
               tok = NextToken();
               if ( tok == TOKEN_EOL )
                  tok = NextToken();
            }
            else
               return WBERR_EXPECTCOMMAOREOL;
         }
      }
   }
   
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBConstStatement( void )
{
   TOKEN tok;
   NPVARDEF npvar;
   VARIANT vResult;
   HLSTR hlstr;
   UINT fuDeclType;
   ERR err;
   
   tok = LastToken();
   
   if ( tok == TOKEN_EOL )
      WBRuntimeError( WBERR_EXPECT_IDENT );

   while ( tok != TOKEN_EOL ) {
      
      if ( tok != TOKEN_IDENT ) 
         WBRuntimeError( WBERR_EXPECT_IDENT );
   
      WBLookupDeclaration( GetTokenText(), &fuDeclType );
      if ( fuDeclType != DECL_NOTFOUND )
         WBRuntimeError( WBERR_DUPLICATEDEFINITION );
            
      npvar = WBInsertLocalVariable( GetTokenText() );
      npvar->fuAttribs = VA_CONST | VA_VARIANT;
            
      if ( NextToken() != TOKEN_EQU )
         WBRuntimeError( WBERR_EXPECT_EQU );
            
      NextToken();

      if ( (err = WBExpression( &vResult )) != 0 )
         WBRuntimeError( err );
            
      tok = LastToken();
            
      switch ( vResult.type ) {
         
         case VT_I4:
            npvar->value.var.tLong = vResult.var.tLong;
            npvar->value.type = VT_I4;
            break;
               
         case VT_STRING:
            // Create permanent copy of string
            hlstr = WBCreateHlstrFromTemp( vResult.var.tString );
            if ( hlstr == NULL )
               WBRuntimeError( WBERR_STRINGSPACE );
            WBDestroyHlstrIfTemp( vResult.var.tString );
            npvar->value.var.tString = hlstr;
            npvar->value.type = VT_STRING;
            break;
               
         default:
            WBRuntimeError( WBERR_TYPEMISMATCH );
      }

      if ( tok != TOKEN_EOL ) {
         if ( tok == TOKEN_COMMA ) {
            tok = NextToken();
            if ( tok == TOKEN_EOL )
               tok = NextToken();
         }
         else
            WBRuntimeError( WBERR_EXPECTCOMMAOREOL );
      }
   }
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBDimStatement( void )
{
   TOKEN tok;
   NPWBHASHENTRY npHash;
   UINT fuDeclType;
   NPVARDEF npvar;
   LPWBTYPEMEM lpTypeMem;
   
   tok = LastToken();
   if ( tok == TOKEN_EOL )
      WBRuntimeError( WBERR_EXPECT_IDENT );
   
   while ( tok != TOKEN_EOL ) {
   
      if ( tok != TOKEN_IDENT ) 
         WBRuntimeError( WBERR_EXPECT_IDENT );
      
      WBLookupDeclaration( GetTokenText(), &fuDeclType );
      if ( fuDeclType != DECL_NOTFOUND )
         WBRuntimeError( WBERR_DUPLICATEDEFINITION );
         
      npvar = WBInsertLocalVariable( GetTokenText() );
      npvar->fuAttribs = VA_VARIANT;
      
      tok = NextToken();   
      
      if ( tok == TOKEN_AS ) {
      
         tok = NextToken();
         
         switch ( tok ) {
         
            case TOKEN_IDENT:
               npHash.VoidPtr = WBLookupDeclaration( GetTokenText(), &fuDeclType );
               if ( fuDeclType != DECL_USERTYPE )
                  WBRuntimeError( WBERR_UNDEFINED_TYPE );
      
               lpTypeMem = WBAllocTypeMem( npHash.UserType, TRUE );
               npvar->fuAttribs = VA_USERTYPE;
               npvar->value.type = VT_USERTYPE;
               npvar->value.var.tlpVoid = lpTypeMem;
               break;
               
            case TOKEN_VARIANT:
               npvar->value.type = VT_I4;
               npvar->value.var.tLong = 0L;
               break;
            
            default:
               WBRuntimeError( WBERR_EXPECT_TYPENAME );
         }
         
         tok = NextToken();
      }
      
      if ( tok != TOKEN_EOL ) {
         if ( tok == TOKEN_COMMA ) {
            tok = NextToken();
            if ( tok == TOKEN_EOL )
               tok = NextToken();
         }
         else
            WBRuntimeError( WBERR_EXPECTCOMMAOREOL );
      }
   }
   
   NextToken();
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBOptionStatement( void )
{
   TOKEN tok;
   VARIANT vTemp;
   ERR err;
   
   tok = LastToken();
   if ( tok == TOKEN_IDENT ) {
      if ( strcmp( GetTokenText(), "COMPARE" ) == 0 ) {
         if ( NextToken() == TOKEN_IDENT ) {
            if ( strcmp( GetTokenText(), "TEXT" ) == 0 )
               g_npTask->fCompareText = TRUE;
            else if ( strcmp( GetTokenText(), "BINARY" ) == 0 )
               g_npTask->fCompareText = FALSE;
            else
               WBRuntimeError( WBERR_EXPECT_TEXT_OR_BINARY );
               
            NextToken();
            WBCheckEndOfLine();
            return;
         }
      }
      else if ( strcmp( GetTokenText(), "DATETIME" ) == 0 ) {
         if ( NextToken() == TOKEN_STRINGLIT ) 
            strcpy( g_npTask->szDatimFmt, GetTokenText() );
         else
            WBRuntimeError( WBERR_EXPECT_STRINGCONST );
         // Validate the date/format string using current date/time
         vTemp.var.tLong = (long)time( NULL );
         vTemp.type = VT_DATE;
         g_npTask->fuDateOrder = DO_RESET;
         if ( (err = WBVariantDateToString( &vTemp )) != 0)
            WBRuntimeError( err );
         WBDestroyHlstrIfTemp( vTemp.var.tString );
         NextToken();
         WBCheckEndOfLine();
         return;
      }
   }
   
   WBRuntimeError( WBERR_EXPECT_IDENT );
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBSimpleAssigment( void )
{
   TOKEN tok;
   NPVARDEF npvar;
   VARIANT vResult;
   HLSTR hlstr;
   ERR err;

   // Search the local and global var tables. If the variable is not
   // found, create a new variable, initially with attribute VA_VARIANT
   // and variant type VT_EMPTY
   npvar = WBLookupVariable( GetTokenText(), TRUE );
   if ( npvar->fuAttribs & VA_CONST )
      WBRuntimeError( WBERR_CONSTASSIGNMENT );
   
   // Skip '=' sign (we have already established that the '=' sign is
   // present)
   NextToken();
   
   // Get token following '=' sign
   tok = NextToken();
   
   if ( tok == TOKEN_NEW ) {

      // The NEW keyword can only be used when assigning to a variable
      // with the VA_VARIANT attribute, i.e. an existing variable with
      // that attribute, or a variable that has just been created as
      // a result of calling WBLookupVariable earlier in this function.
      if ( !(npvar->fuAttribs & VA_VARIANT) ) 
         WBRuntimeError( WBERR_TYPEMISMATCH );
         
      tok = NextToken();
      if ( tok != TOKEN_IDENT ) 
         WBRuntimeError( WBERR_EXPECT_IDENT );

      // Try an create an object of the requested type
      if ( (err = WBCreateObject( GetTokenText(), &vResult )) != 0 )
         WBRuntimeError( err );
         
      WBWireObject( vResult.var.tCtl );
      WBSetVariable( npvar, &vResult );
      NextToken();
   }
   else if ( npvar->fuAttribs & VA_VARIANT ) {
   
      if ( (err = WBExpression( &vResult )) != 0 )
         WBRuntimeError( err );
      
      switch ( vResult.type ) {
      
         case VT_I4:
         case VT_DATE:
            break;
            
         case VT_STRING:
            // Create permanent copy of string
            hlstr = WBCreateHlstrFromTemp( vResult.var.tString );
            if ( hlstr == NULL )
               WBRuntimeError( WBERR_STRINGSPACE );
            WBDestroyHlstrIfTemp( vResult.var.tString );
            vResult.var.tString = hlstr;
            break;
            
         case VT_USERTYPE:
            WBRuntimeError( WBERR_TYPEMISMATCH );
            break;
            
         default:
            if ( vResult.type > VT_OBJECT ) 
               WBWireObject( vResult.var.tCtl );
            else {
               assert( FALSE );
               WBRuntimeError( WBERR_INTERNAL_ERROR );
            }
      }
      
      WBSetVariable( npvar, &vResult );
   }
      
   else if ( npvar->fuAttribs & VA_USERTYPE ) {
      ResetLine();
      NextToken();
      WBUserTypeAssignment( (LPWBTYPEMEM)npvar->value.var.tlpVoid );
   }
   else {
      assert( FALSE );
      WBRuntimeError( WBERR_INTERNAL_ERROR );
   }
   
   WBCheckEndOfLine();
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
// **** TODO: REWORK THIS FUNCTION TO MAKE USE OF WBGetUserTypeData FUNCTION
// DEFINED IN TYPE.C
void WBUserTypeAssignment( LPWBTYPEMEM lpTypeMem )
{
   TOKEN tok;
   NPVARDEF npvar;
   NPWBUSERTYPE npUserType1, npUserType2;
   NPWBTYPEMEMBER npMember;
   LPWBTYPEMEM lpTypeMem2;
   UINT uOffset1, uOffset2;
   LPBYTE lpData;
   VARIANT vResult;
   LPSTR lpsz;
   UINT uLen;
   ERR err;
   
   npUserType1 = lpTypeMem->npUserType;
   uOffset1 = 0;

   tok = NextToken();
   
   while ( tok == TOKEN_PERIOD ) {
      
      if ( NextToken() != TOKEN_IDENT ) 
         WBRuntimeError( WBERR_EXPECT_IDENT );
                           
      npMember = WBFindTypeMember( npUserType1, GetTokenText() );
      if ( npMember == NULL )
         WBRuntimeError( WBERR_UNDEFINEDELEMENT );

      uOffset1 += npMember->uOffset;
      
      // If the leaf member is not a user defined type, call the
      // expression evaluator to get a variant value and assign
      // the result to the leaf member         
      if ( npMember->fMemberType != VT_USERTYPE ) {
      
         if ( NextToken() != TOKEN_EQU )
            WBRuntimeError( WBERR_EXPECT_EQU );
            
         NextToken();

         if ( (err = WBExpression( &vResult )) != NULL )
            WBRuntimeError( err );
         
         if ( LastToken() != TOKEN_EOL )
            WBRuntimeError( WBERR_EXPECT_EOL );
               
         lpData = lpTypeMem->data + uOffset1;
                           
         switch ( npMember->fMemberType ) {
                              
            case VT_I2:
               {
                  LPINT lpint;
                                    
                  if ( (err = WBMakeLongVariant( &vResult )) != 0 ) 
                     WBRuntimeError( err );
                                       
                  if ( vResult.var.tLong < INT_MIN ||
                       vResult.var.tLong > INT_MAX ) 
                     WBRuntimeError( WBERR_VALUEOUTOFRANGE );
                                    
                  lpint = (LPINT)lpData;   
                  *lpint = (int)vResult.var.tLong;
               }
               break;
                              
            case VT_I4:
               {
                  LPLONG lplong;
                  if ( (err = WBMakeLongVariant( &vResult )) != 0 ) 
                     WBRuntimeError( err );
                  lplong = (LPLONG)lpData;
                  *lplong = vResult.var.tLong;
               }
               break;
                                 
            case VT_STRING:
               if ( (err = WBVariantToString( &vResult )) != 0 ) 
                  WBRuntimeError( err );
               lpsz = WBDerefHlstrLen( vResult.var.tString, &((USHORT)uLen) );
               uLen = min( uLen, npMember->uSize );
               if ( uLen > 0 )
                  memcpy( lpData, lpsz, uLen );
               if ( uLen < npMember->uSize )
                  memset( lpData+uLen, ' ', npMember->uSize - uLen );
               WBDestroyHlstrIfTemp( vResult.var.tString );
               break;
                              
            default:
               WBRuntimeError( WBERR_TYPEMISMATCH );
         }
                        
         NextToken();
         return;
      }

      npUserType1 = npMember->npChild;      
      tok = NextToken();
   }
   
   if ( tok != TOKEN_EQU )
      WBRuntimeError( WBERR_EXPECT_EQU );
      
   tok = NextToken();

   if ( tok != TOKEN_IDENT ) {
      if ( LastTokenClass() == TC_KEYWORD )
         WBRuntimeError( WBERR_RESERVEDWORD );
      else
         WBRuntimeError( WBERR_EXPECT_IDENT );
   }
   
   npvar = WBLookupVariable( GetTokenText(), FALSE );
   if ( npvar == NULL )
      WBRuntimeError( WBERR_VARUNDEFINED );

   if ( npvar->value.type != VT_USERTYPE )
      WBRuntimeError( WBERR_TYPEMISMATCH );
            
   lpTypeMem2 = (LPWBTYPEMEM)npvar->value.var.tlpVoid;
   npUserType2 = lpTypeMem2->npUserType;
   uOffset2 = 0;
   
   tok = NextToken();
   
   while ( tok == TOKEN_PERIOD ) {
      
      if ( NextToken() != TOKEN_IDENT ) 
         WBRuntimeError( WBERR_EXPECT_IDENT );
                           
      npMember = WBFindTypeMember( npUserType2, GetTokenText() );
      if ( npMember == NULL )
         WBRuntimeError( WBERR_UNDEFINEDELEMENT );

      if ( npMember->fMemberType != VT_USERTYPE )
         WBRuntimeError( WBERR_TYPEMISMATCH );

      npUserType2 = npMember->npChild;      
      uOffset2 += npMember->uOffset;
      
      tok = NextToken();
   }
   
   if ( tok != TOKEN_EOL )
      WBRuntimeError( WBERR_EXPECT_EOL );
      
   if ( npUserType1 != npUserType2 ) 
      WBRuntimeError( WBERR_TYPEMISMATCH );

   // No-op if assigning to itself
   if ( lpTypeMem->data != lpTypeMem2->data ) {
      memcpy( lpTypeMem->data + uOffset1,
              lpTypeMem2->data + uOffset2,
              npUserType1->uSize );
   }              
}            

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBObjectOperation( HCTL hctl )
{
   TOKEN tok;
   VARIANT vResult;
   long lIndex;
   USHORT iProp;
   LPMETHODINFO lpMethodInfo;
   int narg;
   LPVAR lparg;
   ERR err;
   
   tok = NextToken();
   
   if ( tok != TOKEN_PERIOD )
      WBRuntimeError( WBERR_EXPECT_PERIOD );
      
   if ( NextToken() != TOKEN_IDENT ) 
      WBRuntimeError( WBERR_EXPECT_IDENT );
   
   if ( (iProp = WBLookupProperty( hctl, GetTokenText())) != (USHORT)-1 ) {
         
      if ( NextToken() == TOKEN_LPAREN ) {
         NextToken();
   
         if ( (err = WBExpression( &vResult )) != 0 )
            WBRuntimeError( err );
                  
         if ( (err = WBMakeLongVariant( &vResult )) != 0 )
            WBRuntimeError( err );
                              
         if ( LastToken() != TOKEN_RPAREN )
            WBRuntimeError( WBERR_EXPECT_RPAREN );
                           
         lIndex = vResult.var.tLong;
         if ( lIndex == -1 ) 
            WBRuntimeError( WBERR_INDEXRANGE );
         NextToken();
      }
      else {
         lIndex = -1L;
      }
            
      if ( LastToken() != TOKEN_EQU )
         WBRuntimeError( WBERR_EXPECT_EQU );
                              
      tok = NextToken();
      if ( (err = WBExpression( &vResult )) != 0 )
         WBRuntimeError( err );
                           
      if ( (err = WBSetPropertyVariant( hctl, iProp, lIndex, &vResult )) != 0 ) 
         WBRuntimeError( err );
                              
      WBCheckEndOfLine();
   }
   else if ( (lpMethodInfo = WBLookupMethod( hctl, GetTokenText() )) != NULL ) {    
         
      if ( NextToken() == TOKEN_EOL ) {
         narg = 0;
         lparg = NULL;
      }
      else {
         err = WBBuildArgList( lpMethodInfo->npszTemplate, &narg, &lparg );
         if ( err != 0 )
            WBRuntimeError( err );
      }
      
      err = lpMethodInfo->lpfn( hctl, narg, lparg, &vResult );
      WBFreeArgList( narg, lparg );
      if ( err != 0 )
         WBRuntimeError( err );
                  
      // Discard return value                  
      if ( vResult.type == VT_STRING )
         WBDestroyHlstrIfTemp( vResult.var.tString );
      else if ( vResult.type > VT_OBJECT )
         WBUnWireObject( vResult.var.tCtl );            
         
      WBCheckEndOfLine();
   }
   else {
      WBRuntimeError( WBERR_UNKNOWN_METHODORPROP );
   }
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBSubIntrinsicFunction( NPFUNCTIONINFO npIntFunc )
{
   VARIANT vResult;
   int narg;
   LPVAR lparg;
   ERR err;
   
   if ( NextToken() == TOKEN_EOL ) {
      narg = 0;
      lparg = NULL;
   }
   else {
      err = WBBuildArgList( npIntFunc->szTemplate, &narg, &lparg );
      if ( err != 0 )
         WBRuntimeError( err );

      if ( LastToken() != TOKEN_EOL )
         WBRuntimeError( WBERR_EXPECT_EOL );
   }
                                    
   err = npIntFunc->lpfn( narg, lparg, &vResult );
   WBFreeArgList( narg, lparg );
   if ( err != 0 )
      WBRuntimeError( err );
      
   // Discard return value                  
   if ( vResult.type == VT_STRING )
      WBDestroyHlstrIfTemp( vResult.var.tString );
   else if ( vResult.type > VT_OBJECT )
      WBUnWireObject( vResult.var.tCtl );            
                  
   NextToken();
   return;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBSubUserDefinedFunction( NPWBUSERFUNC npUserFunc )
{
   VARIANT vResult;
   int narg;
   LPVAR lparg;
   ERR err;
   
   if ( NextToken() == TOKEN_EOL ) {
      narg = 0;
      lparg = NULL;
   }
   else {
      if ( (err = WBBuildArgList( NULL, &narg, &lparg )) != 0 )
         WBRuntimeError( err );
                     
      if ( LastToken() != TOKEN_EOL )
         WBRuntimeError( WBERR_EXPECT_EOL );
   }         
                  
   err = WBCallUserFunction( npUserFunc, narg, lparg, &vResult );
   WBFreeArgList( narg, lparg );
   if ( err != 0 )
      WBRuntimeError( err );
   
   // Discard return value                  
   if ( vResult.type == VT_STRING )
      WBDestroyHlstrIfTemp( vResult.var.tString );
   else if ( vResult.type > VT_OBJECT )
      WBUnWireObject( vResult.var.tCtl );            
                  
   NextToken();
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBSubDeclaredFunction( NPWBDECLARE npDeclare )
{
   VARIANT vResult;
   ERR err;
   
   NextToken();
                  
   if ( (err = WBCallDeclaredFunction( npDeclare, &vResult )) != 0 )
      WBRuntimeError( err );
                  
   // Discard return value                  
   if ( vResult.type == VT_STRING )
      WBDestroyHlstrIfTemp( vResult.var.tString );
   else if ( vResult.type > VT_OBJECT )
      WBUnWireObject( vResult.var.tCtl ); 
      
   WBCheckEndOfLine();
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBExecCommand( void )
{
   TOKEN tok;
   VARIANT vResult;
   LPSTR lpsz;
   ERR err;

   tok = LastToken();
   
   switch ( tok ) {
   
      case TOKEN_IDENT:         
         if ( strcmp( GetTokenText(), "EVAL") == 0 ) {
            tok = NextToken();
            goto strexpr_cmd;
         }
         else {
            goto verbatim;
         }

      case TOKEN_STRINGLIT:
strexpr_cmd:      
         // Evaluate command expression and pass to host
         if ( (err = WBExpression( &vResult)) != 0 )
            WBRuntimeError( err );
            
         if ( vResult.type != VT_STRING ) {
            if ( vResult.type > VT_OBJECT )
               WBDestroyObjectIfNoRef( vResult.var.tCtl );
            WBRuntimeError( WBERR_TYPEMISMATCH );
         }
            
         WBCheckEndOfLine();
         
         lpsz = WBDerefZeroTermHlstr( vResult.var.tString );
         if ( lpsz != NULL ) {
            while ( isspace( *lpsz ) )
               lpsz++;
            if ( *lpsz )
               err = WBExtCmd( lpsz );
         }
         
         WBDestroyHlstrIfTemp( vResult.var.tString );
         
         if ( err != 0 )
            WBRuntimeError( -err );
         return;
         
     default:
verbatim:     
         lpsz = g_npTask->npContext->szProgText;
         while ( isspace( *lpsz ) )
            lpsz++;
         if ( *lpsz ) {
            err = WBExtCmd( lpsz );
            if ( err != 0 )
               WBRuntimeError( -err );
         }
         SkipLine();
         NextToken();
         return;
   }
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBCheckEndOfLine( void )
{
   if ( LastToken() != TOKEN_EOL ) 
      WBRuntimeError( WBERR_GARBAGE );
   
   NextToken();
}
   
   
