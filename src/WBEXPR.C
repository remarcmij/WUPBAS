#define STRICT
#include <windows.h>
#include <stdlib.h>
#include <string.h>
#include "wupbas.h"
#include "wbdllcal.h"
#include "wbtype.h"
#include "wbmem.h"

#define FASTCALL __fastcall

static ERR NEAR FASTCALL WBLogicalExpr( LPVAR lpResult );
static ERR NEAR FASTCALL WBShiftExpr( LPVAR lpResult );
static ERR NEAR FASTCALL WBAdditiveExpr( LPVAR lpResult );
static ERR NEAR FASTCALL WBTerm( LPVAR lpResult );
static ERR NEAR FASTCALL WBFactor( LPVAR lpResult );

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBExpression( LPVAR lpResult )
{
   TOKEN tok, op;
   VARIANT vTemp;
   ERR err;

   if ( LastToken() == TOKEN_EOL ) 
      return WBERR_EXPECT_EXPRESSION;
   
   if ( (err = WBLogicalExpr( lpResult ) ) != 0 )
      return err;
         
   while ( LastTokenClass() == TC_LOP ) {
   
      op = LastToken();
      tok = NextToken();
      
      if ( tok == TOKEN_EOL )
         tok = NextToken();
         
      if ( (err = WBLogicalExpr( &vTemp )) != 0 ) 
         return err;
            
      if ( (err = WBMakeLongVariant( lpResult )) != 0 )
         return err;
            
      if ( (err = WBMakeLongVariant( &vTemp )) != 0 )
         return err;

      switch ( op ) {
         case TOKEN_AND:         
            lpResult->var.tLong &= vTemp.var.tLong;
            break;
            
         case TOKEN_OR:
            lpResult->var.tLong |= vTemp.var.tLong;
            break;
            
         case TOKEN_XOR:
            lpResult->var.tLong ^= vTemp.var.tLong;
            break;
            
         default:  
            assert( FALSE );
            return WBERR_INTERNAL_ERROR;
      }
   }
      
   return 0;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*                                                                          */
/* [NOT] subsubexpr1 [ { = | <> | > | >= | < | <= } subsubexpr2 ]           */
/*                                                                          */
/*--------------------------------------------------------------------------*/
static ERR NEAR FASTCALL WBLogicalExpr( LPVAR lpResult )
{
   TOKEN tok, op;
   BOOL fNot;
   int nExprType;
   VARIANT vTemp, vResult;
   ERR err;

   if ( LastToken() == TOKEN_NOT ) {
      fNot = TRUE;
      NextToken();
   }
   else {
      fNot = FALSE;
   }
      
   if ( (err = WBShiftExpr( lpResult ) ) != 0 )
      return err;
   
   if ( LastTokenClass() == TC_RELOP ) {

      op = LastToken();
      tok = NextToken();
      
      if ( tok == TOKEN_EOL )
         tok = NextToken();
         
      if ( ( err = WBShiftExpr( &vTemp )) != 0 )
         return err;
   
      if ( lpResult->type == vTemp.type ) {
         nExprType = lpResult->type;
      }
      else {
         if ( lpResult->type == VT_STRING ) {
            if ( WBVariantStringToOther( lpResult, vTemp.type ) != 0 ) 
               nExprType = VT_STRING;
            else 
               nExprType = vTemp.type;
         }
         else if ( vTemp.type == VT_STRING ) {
            if ( WBVariantStringToOther( &vTemp, lpResult->type ) != 0 )
               nExprType = VT_STRING;
            else
               nExprType = lpResult->type;
         }
         else {
            return WBERR_TYPEMISMATCH;
         }
      }
      
      // Result of logical expression is an Int    
      vResult.type = VT_I4;
      
      // all expression types except of type VT_STRING and of 
      // type > VT_OBJECT are treated as long integer expressions
      if ( nExprType != VT_STRING && nExprType < VT_OBJECT ) {   
         long lComp;

         lComp = lpResult->var.tLong - vTemp.var.tLong;
      
         switch ( op ) {
         
            case TOKEN_EQU:
               vResult.var.tLong = ( lComp == 0 ) ? WB_TRUE : WB_FALSE;
               break;
            
            case TOKEN_NE:
               vResult.var.tLong = ( lComp != 0 ) ? WB_TRUE : WB_FALSE;
               break;
            
            case TOKEN_GT:
               vResult.var.tLong = ( lComp > 0 ) ? WB_TRUE : WB_FALSE;
               break;
            
            case TOKEN_GE:
               vResult.var.tLong = ( lComp >= 0 ) ? WB_TRUE : WB_FALSE;
               break;
            
            case TOKEN_LT:
               vResult.var.tLong = ( lComp < 0 ) ? WB_TRUE : WB_FALSE;
               break;
            
            case TOKEN_LE:
               vResult.var.tLong = ( lComp <= 0 ) ? WB_TRUE : WB_FALSE;
               break;
               
            default:
               assert( FALSE );
               return WBERR_INTERNAL_ERROR;
         }
      }
      
      else if ( nExprType == VT_STRING ) {
         LPSTR lpsz1, lpsz2;
         USHORT usLen, usLen1, usLen2;
         int nComp;
   
         if ( (err = WBVariantToString( lpResult )) != 0 ||
              (err = WBVariantToString( &vTemp )) != 0 ) {
            return err;
         }
            
         lpsz1 = WBDerefHlstrLen( lpResult->var.tString, &usLen1 );
         lpsz2 = WBDerefHlstrLen( vTemp.var.tString, &usLen2 );
            
         usLen = min( usLen1, usLen2 );
         
         if ( usLen > 0 ) {
            assert( lpsz1 != NULL );
            assert( lpsz2 != NULL );
            if ( g_npTask->fCompareText )
               nComp = memicmp( lpsz1, lpsz2, usLen );
            else
               nComp = memcmp( lpsz1, lpsz2, usLen );
         }
         else
            nComp = 0;
         
         if ( nComp == 0 ) {
            if ( usLen1 > usLen2 ) 
               nComp = 1;
            else if ( usLen1 < usLen2 )
               nComp = -1;
         }
                           
         switch ( op ) {
         
            case TOKEN_EQU:
               vResult.var.tLong = ( nComp == 0 ) ? WB_TRUE : WB_FALSE;
               break;
            
            case TOKEN_NE:
               vResult.var.tLong = ( nComp != 0 ) ? WB_TRUE : WB_FALSE;
               break;
            
            case TOKEN_GT:
               vResult.var.tLong = ( nComp > 0 ) ? WB_TRUE : WB_FALSE;
               break;
            
            case TOKEN_GE:
               vResult.var.tLong = ( nComp >= 0 ) ? WB_TRUE : WB_FALSE;
               break;
            
            case TOKEN_LT:
               vResult.var.tLong = ( nComp < 0 ) ? WB_TRUE : WB_FALSE;
               break;
                        
            case TOKEN_LE:
               vResult.var.tLong = ( nComp <= 0 ) ? WB_TRUE : WB_FALSE;
               break;
               
            default:
               assert( FALSE );
               return WBERR_INTERNAL_ERROR;
         }
         
         WBDestroyHlstrIfTemp( lpResult->var.tString );
         WBDestroyHlstrIfTemp( vTemp.var.tString );
      }
      else {
         // No comparisons of object variables
         return WBERR_TYPEMISMATCH;
      }
      
      *lpResult = vResult;
   }
   
   
   if ( fNot ) {
      if ( (err = WBMakeLongVariant( lpResult )) != 0 )
         return err;
      lpResult->var.tLong = ~(lpResult->var.tLong);
   }
   
   return 0;         
}            

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
static ERR NEAR FASTCALL WBShiftExpr( LPVAR lpResult )
{
   TOKEN tok, op;
   VARIANT vTemp;
   ERR err;
   
   if ( (err = WBAdditiveExpr( lpResult )) != 0 )
      return err;
      
   if ( LastTokenClass() == TC_SHIFTOP ) {
   
      op = LastToken();
      tok = NextToken();
      
      if ( tok == TOKEN_EOL )
         tok = NextToken();
         
      if ( (err = WBAdditiveExpr( &vTemp )) != 0 )
         return err;
         
      if ( (err = WBMakeLongVariant( lpResult )) != 0 )
         return err;
         
      if ( (err = WBMakeLongVariant( &vTemp )) != 0 )
         return err;
      
      if ( vTemp.var.tLong < 0 || vTemp.var.tLong > sizeof(long) * 8 )
         return WBERR_ARGRANGE;
         
      if ( op == TOKEN_SHIFTLEFT )
         lpResult->var.tLong = (DWORD)lpResult->var.tLong << 
                                 (WORD)vTemp.var.tLong;      
      else
         lpResult->var.tLong = (DWORD)lpResult->var.tLong >>
                                 (WORD)vTemp.var.tLong;      
   }
   
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
static ERR NEAR FASTCALL WBAdditiveExpr( LPVAR lpResult )
{
   TOKEN tok, op;
   BOOL fGotSign;
   BOOL fNegate;
   VARIANT vTemp;
   HLSTR hlstr;
   ERR err;

   if ( LastTokenClass() == TC_ADDOP ) {   
      switch ( LastToken() ) {
         case TOKEN_PLUS:
            tok = NextToken();
            fGotSign = TRUE;
            fNegate = FALSE;
            break;
         case TOKEN_MINUS:
            tok = NextToken();
            fGotSign = TRUE;
            fNegate = TRUE;
            break;
         default:
            assert( FALSE );
            return WBERR_INTERNAL_ERROR;
      }
   }
   else {
      fGotSign = FALSE;
      fNegate = FALSE;
   }
   
   if ( ( err = WBTerm( lpResult) ) != 0 )
      return err;
   
   if ( fGotSign ) {
      if ( lpResult->type != VT_I4 ) {
         return WBERR_BADEXPRESSION;
      }
      if ( fNegate )
         lpResult->var.tLong = -lpResult->var.tLong;
   }   
      
   while ( LastTokenClass() == TC_ADDOP ) {
           
      op = LastToken();
      tok = NextToken();

      // Allow continuation of expression on next line if  
      // operator was last token on previous line
      if ( tok == TOKEN_EOL )
         NextToken();
         
      if ( (err = WBTerm( &vTemp ) ) != 0 )
         return err;
      
      switch ( op ) {
      
         case TOKEN_MINUS: 
            if ( (err = WBMakeLongVariant( lpResult )) != 0 )
               return err;
            if ( (err = WBMakeLongVariant( &vTemp )) != 0 )
               return err;
            lpResult->var.tLong -= vTemp.var.tLong;
            break;
            
         case TOKEN_PLUS:
            if ( (err = WBMakeLongVariant( lpResult )) != 0 )
               return err;
            if ( (err = WBMakeLongVariant( &vTemp )) != 0 )
               return err;
            lpResult->var.tLong += vTemp.var.tLong;
            break;
            
         case TOKEN_CONCAT:
            {
               LPSTR lpsz1, lpsz2;
               USHORT usLen, usLen1, usLen2;
               
               if ( (err = WBVariantToString( lpResult )) != 0 )
                  return err;
               if ( (err = WBVariantToString( &vTemp )) != 0 )
                  return err;
                  
               usLen1 = WBGetHlstrLen( lpResult->var.tString );   
               usLen2 = WBGetHlstrLen( vTemp.var.tString );   

               if ( (DWORD)usLen1 + (DWORD)usLen2 > MAXSTRING ) {
                  return WBERR_STRINGOVERFLOW;
               }
                
               usLen = usLen1 + usLen2;

               // return null string if concatenation gives null string
               if ( usLen == 0 ) {
                  // Note: lpResult already contains null string
                  // vTemp also contains null string and therefore
                  // does not need cleaning up
                  break;
               }
                                 
               hlstr = WBCreateTempHlstr( NULL, usLen );
               if ( hlstr == NULL ) {
                  return WBERR_STRINGSPACE;
               }
                  

               if ( usLen1 > 0 ) {
                  lpsz1 = WBDerefHlstr( hlstr );
                  assert( lpsz1 != NULL );
                  assert( lpResult->type == VT_STRING );               
                  lpsz2 = WBDerefHlstr( lpResult->var.tString );
                  assert( lpsz2 != NULL );
                  memcpy( lpsz1, lpsz2, usLen1 );
               }
               WBDestroyHlstrIfTemp( lpResult->var.tString );
               
               if ( usLen2 > 0 ) {
                  lpsz1 = WBDerefHlstr( hlstr );
                  assert( lpsz1 != NULL );
                  assert( vTemp.type == VT_STRING );               
                  lpsz2 = WBDerefHlstr( vTemp.var.tString );
                  assert( lpsz2 != NULL );
                  memcpy( lpsz1 + usLen1, lpsz2, usLen2 );
               }
               WBDestroyHlstrIfTemp( vTemp.var.tString );
               
               lpResult->var.tString = hlstr;
               break;
            }                  

         default:
            assert( FALSE );
            return WBERR_INTERNAL_ERROR;
      }
  }
  
  return 0;
}
      
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
static ERR NEAR FASTCALL WBTerm( LPVAR lpResult )
{
   TOKEN tok, op;
   VARIANT vTemp;
   ERR err;
   
   if ( (err = WBFactor( lpResult ) ) != 0 )
      return err;
      
   while ( LastTokenClass() == TC_MULTOP ) {   
   
      op = LastToken();
      tok = NextToken();
      
      if ( tok == TOKEN_EOL )
         tok = NextToken();

      if ( ( err = WBFactor( &vTemp ) ) != 0 )
         return err;

      if ( (err = WBMakeLongVariant( lpResult )) != 0 ||
           (err = WBMakeLongVariant( &vTemp )) != 0 ) {
         return err;
      }

      switch ( op ) {
      
         case TOKEN_MULT:         
            lpResult->var.tLong *= vTemp.var.tLong;
            break;
            
         case TOKEN_DIV:
            lpResult->var.tLong /= vTemp.var.tLong;
            break;
            
         case TOKEN_MOD:
            lpResult->var.tLong %= vTemp.var.tLong;
            break;
            
         default:
            assert( FALSE );
            return WBERR_INTERNAL_ERROR;
      }
   }
   
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
static ERR NEAR FASTCALL WBFactor( LPVAR lpResult )
{
   TOKEN tok;
   ERR err;
   NPWBHASHENTRY npHash;
   UINT fuDeclType;
   NPWBUSERTYPE npUserType;
   UINT fMemberType;
   LPBYTE lpData;
   UINT uSize;
   LPMETHODINFO lpMethodInfo;
   NPVARDEF npvar;
   HCTL hctl;
   VARIANT vTemp;
   long lIndex;
   int narg;
   LPVAR lparg;
   USHORT iProp;

   g_npTask->fTrueVar = FALSE;
   
   tok = LastToken();
   
   switch ( tok ) {
   
      case TOKEN_LONGINT:
         lpResult->var.tLong = LastInteger();
         lpResult->type = VT_I4;
         NextToken();
         return 0;
         
      case TOKEN_IDENT:
         npvar = WBLookupVariable( GetTokenText(), FALSE );
         if ( npvar != NULL ) {
            
            switch ( npvar->value.type ) {
            
               case VT_I4:
               case VT_STRING:
               case VT_DATE:
                  err = WBGetVariableValue( npvar, lpResult );
                  if ( err == 0 ) {
                     g_npTask->fTrueVar = TRUE;
                     NextToken();
                  }
                  return err;
            
               case VT_USERTYPE:
                  if ( NextToken() == TOKEN_PERIOD ) {
                     err = WBGetUserTypeData( (LPWBTYPEMEM)npvar->value.var.tlpVoid,
                                              &npUserType,
                                              &fMemberType,
                                              &lpData,
                                              &uSize );
                     if ( err != 0 )
                        return err;
                        
                     switch ( fMemberType ) {
                     
                        case VT_I2:
                           lpResult->var.tLong = (long)*((LPINT)lpData);
                           lpResult->type = VT_I4;
                           return 0;
                           
                        case VT_I4:
                           lpResult->var.tLong = *((LPLONG)lpData);
                           lpResult->type = VT_I4;
                           return 0;
                           
                        case VT_STRING:
                           lpResult->var.tString = WBCreateTempHlstr( (LPSTR)lpData, (USHORT)uSize );
                           if ( lpResult->var.tString == NULL )
                              return WBERR_STRINGSPACE;
                           lpResult->type = VT_STRING;
                           return 0;
                           
                        case VT_USERTYPE:
                           return WBERR_TYPEMISMATCH;
                           
                        default:
                           assert( FALSE );
                           return WBERR_INTERNAL_ERROR;
                     }
                  }
                  else {
                     lpResult->type = VT_USERTYPE;
                     lpResult->var.tlpVoid = npvar->value.var.tlpVoid;
                     return 0;
                  }
                              
               case VT_EMPTY:
                  return WBERR_VAR_UNINITIALIZED;
                  
               default:
                  if ( npvar->value.type > VT_OBJECT ) {
                  
                     // If not followed by a period, return the
                     // object itself
                     if ( NextToken() != TOKEN_PERIOD ) {
                        *lpResult = npvar->value;
                        return 0;
                     }
                     
                     // Else go look for a property
                     hctl = npvar->value.var.tCtl;
                      
                     if ( NextToken() != TOKEN_IDENT ) 
                        return WBERR_EXPECT_IDENT;
                     
                     if ( (iProp = WBLookupProperty(hctl, GetTokenText() )) != (USHORT)-1 ) {
                        tok = NextToken();
                        if ( tok == TOKEN_LPAREN ) {
                           NextToken();
                           if ( (err = WBExpression( &vTemp )) != 0 )
                              return err;
                           
                           if ( (err = WBMakeLongVariant( &vTemp )) != 0 )
                              return err;
                           
                           lIndex = vTemp.var.tLong;
                           if ( lIndex == -1 ) 
                              return WBERR_INDEXRANGE;
                           
                           if ( LastToken() != TOKEN_RPAREN )
                              return WBERR_EXPECT_RPAREN;
                           NextToken();
                        }
                        else { 
                           lIndex = -1;
                        }
                              
                        return WBGetPropertyVariant( hctl, iProp, lIndex, lpResult );
                     }
                     else if ( (lpMethodInfo = WBLookupMethod( hctl, GetTokenText())) != NULL ) {
                        if ( NextToken() != TOKEN_LPAREN )
                           return WBERR_EXPECT_LPAREN;
                           
                        if ( NextToken() == TOKEN_RPAREN ) {
                           narg = 0;
                           lparg = 0;
                        }
                        else {
                           err = WBBuildArgList( lpMethodInfo->npszTemplate, &narg, &lparg );
                           if ( err != 0 )
                              return err;
                        }
                        err = lpMethodInfo->lpfn( hctl, narg, lparg, lpResult );
                        WBFreeArgList( narg, lparg );
                        if ( err != 0 )
                           return err;
                        if ( LastToken() != TOKEN_RPAREN )
                           return WBERR_EXPECT_RPAREN;
                        NextToken();
                        return 0;
                     }
                     else {
                        return WBERR_UNKNOWN_PROP_OR_METHOD;
                     }
                  }
                  else {
                     assert( FALSE );
                     return WBERR_INTERNAL_ERROR;
                  }
            }
         }
         else {
            npHash.VoidPtr = WBLookupDeclaration( GetTokenText(), &fuDeclType );
            
            switch ( fuDeclType ) {
            
               case DECL_INTFUNC:
                  if ( NextToken() != TOKEN_LPAREN )
                     WBRuntimeError( WBERR_EXPECT_LPAREN );
                     
                  if ( NextToken() == TOKEN_RPAREN ) {
                     narg = 0;
                     lparg = NULL;
                  }
                  else if ( (err = WBBuildArgList( npHash.IntFunc->npIntFunc->szTemplate, &narg, &lparg )) != 0 )
                     return err;
                  err = npHash.IntFunc->npIntFunc->lpfn( narg, lparg, lpResult );
                  WBFreeArgList( narg, lparg );
                  if ( err != 0 )
                     return err;
                     
                  tok = LastToken();
                  if ( tok != TOKEN_RPAREN )
                     return WBERR_EXPECT_RPAREN;
                     
                  NextToken();
                  return 0;
                  
               case DECL_USERFUNC:
                  if ( NextToken() != TOKEN_LPAREN )
                     WBRuntimeError( WBERR_EXPECT_LPAREN );
                     
                  if ( NextToken() == TOKEN_RPAREN ) {
                     narg = 0;
                     lparg = NULL;
                  }
                  else if ( (err = WBBuildArgList( NULL, &narg, &lparg )) != 0 )
                     return err;
                  err = WBCallUserFunction( npHash.UserFunc, narg, lparg, lpResult );
                  WBFreeArgList( narg, lparg );
                  if ( err != 0 )
                     return err;
                  
                  tok = LastToken();
                  if ( tok != TOKEN_RPAREN )
                     return WBERR_EXPECT_RPAREN;
                  
                  NextToken();
                  return 0;
                  
               case DECL_DECLARE:
                  if ( NextToken() != TOKEN_LPAREN )
                     WBRuntimeError( WBERR_EXPECT_LPAREN );
                  NextToken();
                  if ( (err = WBCallDeclaredFunction( npHash.Declare, lpResult )) != 0 )
                     return err;
                  tok = LastToken();
                  if ( tok != TOKEN_RPAREN )
                     return WBERR_EXPECT_RPAREN;
                  NextToken();
                  return 0;
                  
               default:
                  switch( PeekToken() ) {
                     case TOKEN_LPAREN:
                        return WBERR_UNDEFINED_FUNCTION;
                     case TOKEN_PERIOD:
                        return WBERR_OBJECTUNDEFINED;
                     default:
                        return WBERR_VARUNDEFINED;
                  }
            }
         }
         
      case TOKEN_FALSE:
         lpResult->var.tLong = WB_FALSE;
         lpResult->type = VT_I4;
         NextToken();
         return 0;
         
      case TOKEN_TRUE:
         lpResult->var.tLong = WB_TRUE;
         lpResult->type = VT_I4;
         NextToken();
         return 0;
         
      case TOKEN_LPAREN:
         NextToken();                   
         
         if ( (err = WBExpression( lpResult ) ) != 0 )
            return err;
            
         tok = LastToken();
         if ( tok != TOKEN_RPAREN ) 
            return WBERR_EXPECT_RPAREN;
            
         NextToken();
         return 0;
         
      case TOKEN_STRINGLIT:
         lpResult->var.tString = WBCreateTempHlstr( 
                                    GetTokenText(), 
                                    (USHORT)strlen( GetTokenText() ) );
         if ( lpResult->var.tString == NULL )
            return WBERR_STRINGSPACE;
                                                
         lpResult->type = VT_STRING;                                    
         NextToken();
         return 0;

      default:
         if ( LastTokenClass() == TC_KEYWORD )
            return WBERR_RESERVEDWORD;
         else
            return WBERR_BADEXPRESSION;
   }
}
