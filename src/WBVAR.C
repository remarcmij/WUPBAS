#define STRICT
#include <windows.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <time.h>
#include "wupbas.h"
#include "wbtype.h"

NPSTR anpszDateSpec[] = {
   "D",
   "DD",
   "DDD",
   "DDDD",
   "H",
   "HH",
   "M",
   "MM",
   "MMM",
   "MMMM",
   "S",
   "SS",
   "YY",
   "YYYY"
};

enum WBDATESPEC {
   DATESPEC_D,
   DATESPEC_DD,
   DATESPEC_DDD,
   DATESPEC_DDDD,
   DATESPEC_H,
   DATESPEC_HH,
   DATESPEC_M,
   DATESPEC_MM,
   DATESPEC_MMM,
   DATESPEC_MMMM,
   DATESPEC_S,
   DATESPEC_SS,
   DATESPEC_YY,
   DATESPEC_YYYY
};

NPSTR anpszDayName[] = {
   "Sunday",
   "Monday",
   "Tuesday",
   "Wednesday",
   "Thursday",
   "Friday",
   "Saturday"
};

PSTR apszMonthName[] = {
   "January",
   "February",
   "March",
   "April",
   "May",
   "June",
   "July",
   "August",
   "September",
   "October",
   "November",
   "December"
};
   
void WBCreateLocalVarTable( void );

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
NPVARDEF WBInsertLocalVariable( LPCSTR lpszName )
{
   NPVARDEF npvar;
   UINT uHash;

   assert( g_npContext != NULL );

   // If a local variable table has not been created yet, create
   // one now
   if ( g_npContext->npvtLocal == NULL )
      WBCreateLocalVarTable();
      
   uHash = WBHashName( lpszName, VAR_TABLE_SIZE );

   for ( npvar = g_npContext->npvtLocal[uHash];
         npvar != NULL; 
         npvar = npvar->npNext ) {
            
      if ( strcmp( npvar->szName, lpszName ) == 0 ) 
         return npvar;
   }

   // Request zero-initialized local memory   
   npvar = (NPVARDEF)WBLocalAlloc(LPTR, sizeof(VARDEF)+strlen(lpszName));
   if ( npvar == NULL )
      WBRuntimeError( WBERR_OUTOFMEMORY );
      
   // Note: variable is implicitly initialized as a variant of
   // type VT_EMPTY
      
   strcpy( npvar->szName, lpszName );
      
   // store in local var table
   npvar->npNext = g_npContext->npvtLocal[uHash];
   g_npContext->npvtLocal[uHash] = npvar;
   
   return npvar;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
NPVARDEF WBInsertGlobalVariable( LPCSTR lpszName )
{
   NPVARDEF npvar;
   UINT uHash;
   
   uHash = WBHashName( lpszName, VAR_TABLE_SIZE );

   for ( npvar = g_npTask->npvtGlobal[uHash];
         npvar != NULL; 
         npvar = npvar->npNext ) {
            
      if ( strcmp( npvar->szName, lpszName ) == 0 ) 
         return npvar;
   }

   // Request zero-initialized local memory   
   npvar = (NPVARDEF)WBLocalAlloc(LPTR, sizeof(VARDEF)+strlen(lpszName));
   if ( npvar == NULL )
      WBRuntimeError( WBERR_OUTOFMEMORY );
      
   // Note: variable is implicitly initialized as a variant of
   // type VT_EMPTY
      
   strcpy( npvar->szName, lpszName );
      
   // store in global var table
   npvar->npNext = g_npTask->npvtGlobal[uHash];
   g_npTask->npvtGlobal[uHash] = npvar;
   
   return npvar;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
NPVARDEF WBLookupVariable( LPSTR lpszName, BOOL fCreate )
{
   NPVARDEF npvar;
   UINT uHash;
   UINT fuDeclType;
   
   // If a local variable table has not been created yet, create
   // one now
   if ( g_npContext->npvtLocal == NULL )
      WBCreateLocalVarTable();
      
   uHash = WBHashName( lpszName, VAR_TABLE_SIZE );

   // Search local var table
   for ( npvar = g_npContext->npvtLocal[uHash];
         npvar != NULL; 
         npvar = npvar->npNext ) {
            
      if ( strcmp( npvar->szName, lpszName ) == 0 ) 
         return npvar;
   }
   
   // Search global var table
   for ( npvar = g_npTask->npvtGlobal[uHash];
         npvar != NULL; 
         npvar = npvar->npNext ) {
            
      if ( strcmp( npvar->szName, lpszName ) == 0 ) 
         return npvar;
   }
   
   // If not in tables already, create a new local variable if so requested
   if ( fCreate ) {

      WBLookupDeclaration( lpszName, &fuDeclType );
      if ( fuDeclType != DECL_NOTFOUND )
         WBRuntimeError( WBERR_DUPLICATEDEFINITION );
         
      // Request zero-initialized local memory   
      npvar = (NPVARDEF)WBLocalAlloc(LPTR, sizeof(VARDEF)+strlen(lpszName));
      if ( npvar == NULL )
         WBRuntimeError( WBERR_OUTOFMEMORY );
      
      // Note: variable is implicitly initialized as a variant of
      // type VT_EMPTY
      strcpy( npvar->szName, lpszName );
      npvar->fuAttribs = VA_VARIANT;
      
      // store in local var table
      npvar->npNext = g_npContext->npvtLocal[uHash];
      g_npContext->npvtLocal[uHash] = npvar;
   
      return npvar;
   }
   
   return NULL;            
   
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBGetVariableValue( NPVARDEF npvar, LPVAR lpVariant )
{

   if ( npvar->value.type > VT_OBJECT ) {
      if ( WBGetControlModel( npvar->value.var.tCtl ) == NULL )
         return WBERR_OBJECTDESTROYED;
   }
   else if ( npvar->value.type == VT_STRING ) {
      assert( npvar->value.var.tString != NULL );
   }
   else if ( npvar->value.type == VT_EMPTY )
      return WBERR_VAR_UNINITIALIZED;
                  
   *lpVariant = npvar->value;
   g_npTask->npTrueVar = npvar;
   
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBSetVariable( NPVARDEF npvar, LPVAR lpVariant )
{
   WBDestroyVariantValue( &(npvar->value) );
   npvar->value = *lpVariant;   
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
UINT WBHashName( LPCSTR lpszName, UINT uTableSize )
{
   register UINT uHash = 0;
   
   while ( *lpszName )
      uHash = (uHash<<5) + uHash + *lpszName++;
      
   return ( uHash % uTableSize );
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBMakeLongVariant( LPVAR lpvar )
{
   switch ( lpvar->type ) {
   
      case VT_I4:
         return 0;
         
      case VT_STRING:   
         return WBVariantStringToOther( lpvar, VT_I4 );
         
      case VT_USERTYPE:
         return WBERR_TYPEMISMATCH;
         
      default:
         return WBERR_NUM_CONVERSION;
   }
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBVariantToString( LPVAR lpvar )
{
   char szBuffer[MAXTEXT];
   ERR err;

   err = 0;

   switch ( lpvar->type ) {
   
      case VT_STRING:
         break;
         
      case VT_I4:   
         _ltoa( lpvar->var.tLong, szBuffer, 10 );
         lpvar->var.tString = WBCreateTempHlstr( szBuffer, (USHORT)strlen( szBuffer ) );
         if ( lpvar->var.tString == NULL )
            err = WBERR_STRINGSPACE;
         lpvar->type = VT_STRING;
         break;
         
      case VT_DATE:
         err = WBVariantDateToString( lpvar );
         break;
         
       default:
         // Only variant types are VT_STRING, VT_I4 and VT_DATE are
         // supported.
         err = WBERR_TYPEMISMATCH;
   }
   
   return err;
}
      
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBVariantStringToOther( LPVAR lpvar, int nVarType )
{
   static char szBuffer[64];
   VARIANT vTemp;
   LPSTR lpsz;
   USHORT usLen;
   PSTR lpch;
   long lValue;
   ERR err;

   assert( lpvar->type == VT_STRING );   
   
   usLen = WBGetHlstrLen( lpvar->var.tString );
   if ( usLen < 1 || usLen > sizeof( szBuffer )-1 )
      return WBERR_TYPEMISMATCH;

   // Begin note: following lines are equivalent to WNGetHlstr except
   // that it does NOT destroy temporary string
   lpsz = WBDerefHlstrLen( lpvar->var.tString, &usLen );
   if ( usLen > 0 ) {
      // If the string doesn't fit in the buffer, it can be anything
      // that can be represented in a long
      if ( usLen > sizeof(szBuffer) - 1 )
         return WBERR_TYPEMISMATCH;
         
      memcpy( szBuffer, lpsz, usLen );
   }
   // End note
   
   szBuffer[usLen] = '\0';
   
   switch ( nVarType ) {

      case VT_I4:   
         lpch = szBuffer;
         if ( *lpch == '+' || *lpch == '-' ) {
            lpch++;
            usLen--;
         }
         if ( usLen == 0 || strspn( lpch, "0123456789" ) != usLen )
            return WBERR_NUM_CONVERSION;
         lValue = atol( szBuffer );
         break;
         
      case VT_DATE:
         err = WBParseDateTimeString( szBuffer, &vTemp );
         if ( err != 0 )
            return WBERR_DATE_CONVERSION;
         assert( vTemp.type == VT_DATE );
         lValue = vTemp.var.tLong;
         break;
         
      default:
         return WBERR_TYPEMISMATCH;
   }                            

   // Get rid of string
   WBDestroyHlstrIfTemp( lpvar->var.tString );
   
   // set new value and type for variant
   lpvar->var.tLong = lValue;
   lpvar->type = nVarType;      
   return 0;
}
      
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBVariantDateToString( LPVAR lpVariant )
{
   TOKEN tok;
   long lTemp;
   LPSTR lpch;
   struct tm *ptm;
   int i, m;
   enum WBDATESPEC n;
   char szBuffer[MAXTEXT];
   char szDateOrder[7];
   ERR err;
   
   if ( lpVariant->type != VT_DATE )
      return WBERR_TYPEMISMATCH;
      
   // Create a new context
   if ( WBPushContext() == NULL )
      return WBERR_OUTOFMEMORY;
      
   strcpy( g_npContext->szProgText, g_npTask->szDatimFmt );
   g_npContext->npchNextPos = g_npContext->szProgText;
      
   // Break up the date/time value into its individual components
   lTemp = lpVariant->var.tLong;
   ptm = localtime( (time_t *)&lTemp );
   
   lpch = szBuffer;
   m = 0;
   err = 0;
   
   // Get the first token
   tok = NextToken();
   
   // Format the date/time string until the end-on-line token is
   // found in the format picture.
   while ( err == 0 && tok != TOKEN_EOL ) {     

      // Insert intervening space
      for ( i = 0; i < g_npContext->nLeadingSpaces; i++ )
         *lpch++ = ' ';
               
      if ( tok == TOKEN_IDENT ) {

         for ( n = 0; n < sizeof( anpszDateSpec ) / sizeof( PSTR ); n++ ) {
            if ( _stricmp( g_npContext->szToken, anpszDateSpec[n] ) == 0 )
               break;
         }
         
         switch ( n ) {
         
            case DATESPEC_D:
               if ( m > 2 )
                  err = WBERR_DATIMFMT;
               else {
                  itoa( ptm->tm_mday, lpch, 10 );
                  lpch += strlen( lpch );
                  szDateOrder[m++] = 'D';
               }
               break;
               
            case DATESPEC_DD:
               if ( m > 2 )
                  err = WBERR_DATIMFMT;
               else {
                  if ( ptm->tm_mday < 10 ) 
                     *lpch++ = '0';
                  itoa( ptm->tm_mday, lpch, 10 );
                  lpch += strlen( lpch );
                  szDateOrder[m++] = 'D';
               }
               break;
               
            case DATESPEC_DDD:
               if ( m > 2 )
                  err = WBERR_DATIMFMT;
               else {
                  memcpy( lpch, anpszDayName[ptm->tm_wday], 3 );
                  lpch += 3;
                  szDateOrder[m++] = 'D';
               }
               break;
               
            case DATESPEC_DDDD:
               if ( m > 2 )
                  err = WBERR_DATIMFMT;
               else {
                  strcpy( lpch, anpszDayName[ptm->tm_wday] );
                  lpch += strlen( lpch );
                  szDateOrder[m++] = 'D';
               }
               break;
               
            case DATESPEC_M:
               if ( m > 6 )
                  err = WBERR_DATIMFMT;
               else {
                  if ( m > 2 ) 
                     itoa( ptm->tm_min, lpch, 10 );
                  else
                     itoa( ptm->tm_mon + 1, lpch, 10 );
                  lpch += strlen( lpch );
                  szDateOrder[m++] = 'M';
               }
               break;
               
            case DATESPEC_MM:
               if ( m > 6 )
                  err = WBERR_DATIMFMT;
               else {
                  if ( m > 2 ) {            
                     if ( ptm->tm_min < 10 )
                        *lpch++ = '0';
                     itoa( ptm->tm_min, lpch, 10 );
                  }
                  else {
                     if ( ptm->tm_mon + 1 < 10 )
                        *lpch++ = '0';
                     itoa( ptm->tm_mon + 1, lpch, 10 );
                  }
                  lpch += strlen( lpch );
                  szDateOrder[m++] = 'M';
               }
               break;
               
            case DATESPEC_MMM:
               if ( m > 2 )
                  err = WBERR_DATIMFMT;
               else {
                  memcpy( lpch, apszMonthName[ptm->tm_mon], 3 );
                  lpch += 3;
                  szDateOrder[m++] = 'M';
               }
               break;
               
            case DATESPEC_MMMM:
               if ( m > 2 )
                  err = WBERR_DATIMFMT;
               else {
                  strcpy( lpch, apszMonthName[ptm->tm_mon] );
                  lpch += strlen( lpch );
                  szDateOrder[m++] = 'M';
               }
               break;
               
            case DATESPEC_YY:
               if ( m > 2 )
                  err = WBERR_DATIMFMT;
               else {
                  int nYear = ptm->tm_year;
                  if ( nYear >= 100 )
                     nYear -= 100;
                  if ( nYear < 10 ) 
                     *lpch++ = '0';
                  itoa( nYear, lpch, 10 );
                  lpch += strlen( lpch );
                  szDateOrder[m++] = 'Y';
               }
               break;   
                  
            case DATESPEC_YYYY:
               if ( m > 2 )
                  err = WBERR_DATIMFMT;
               else {
                  int nYear = ptm->tm_year + 1900;
                  itoa( nYear, lpch, 10 );
                  lpch += strlen( lpch );
                  szDateOrder[m++] = 'Y';
               }
               break;   
               
            case DATESPEC_H:
               if ( m <= 2 || m > 6 )
                  err = WBERR_DATIMFMT;
               else {
                  itoa( ptm->tm_hour, lpch, 10 );
                  lpch += strlen( lpch );
                  szDateOrder[m++] = 'H';
               }
               break;
               
            case DATESPEC_HH:
               if ( m <= 2 || m > 6 )
                  err = WBERR_DATIMFMT;
               else {
                  if ( ptm->tm_hour < 10 )
                     *lpch++ = '0';
                  itoa( ptm->tm_hour, lpch, 10 );
                  lpch += strlen( lpch );
                  szDateOrder[m++] = 'H';
               }
               break;
               
            case DATESPEC_S:
               if ( m <= 2 || m > 6 )
                  err = WBERR_DATIMFMT;
               else {
                  itoa( ptm->tm_sec, lpch, 10 );
                  lpch += strlen( lpch );
                  szDateOrder[m++] = 'S';
               }
               break;
               
            case DATESPEC_SS:
               if ( m <= 2 || m > 6 )
                  err = WBERR_DATIMFMT;
               else {
                  if ( ptm->tm_sec < 10 )
                     *lpch++ = '0';
                  itoa( ptm->tm_sec, lpch, 10 );
                  lpch += strlen( lpch );
                  szDateOrder[m++] = 'S';
               }
               break;
               
            default:
               err = WBERR_DATIMFMT;
         }
      }
      else if ( tok == TOKEN_STRINGLIT ) {
         strcpy( lpch, g_npContext->szToken );
         lpch += strlen( lpch );
      }
      else {
      
         switch ( tok ) {
            case TOKEN_COLON:
               *lpch++ = ':';
               break;
               
            case TOKEN_DIV:
               *lpch++ = '/';
               break;
               
            case TOKEN_MINUS:
               *lpch++ = '-';
               break;
               
            default:
               err = WBERR_DATIMFMT;
         }
         
         if ( err == 0 ) {
            if ( m <= 2 ) {
               if ( tok != TOKEN_DIV && tok != TOKEN_MINUS ) 
                  err = WBERR_DATIMFMT;
            }
            else if ( tok != TOKEN_COLON ) 
               err = WBERR_DATIMFMT;
         }
      }

      tok = NextToken();
   }

   if ( err == 0 && g_npTask->fuDateOrder == DO_RESET ) {
      // If the date order of the current date/time picture format
      // (default or as the result of executing the Option DateTime
      // statement) has not yet been set, or has been reset, determined
      // the order now.
      szDateOrder[m] = '\0';
      if ( strcmp( szDateOrder, "DMYHMS" ) == 0  )
         g_npTask->fuDateOrder = DO_DMY;
      else if ( strcmp( szDateOrder, "MDYHMS" ) == 0 )
         g_npTask->fuDateOrder = DO_MDY;
      else if ( strcmp( szDateOrder, "YMDHMS" ) == 0 ) 
         g_npTask->fuDateOrder = DO_YMD;
      else
         err = WBERR_DATIMFMT;
   }
   
   if ( err == 0 ) {
      *lpch = '\0';         
      lpVariant->var.tString = WBCreateTempHlstr( szBuffer, 
                                                  (USHORT)strlen( szBuffer ));
      lpVariant->type = VT_STRING;
   }

   WBPopContext();   
   
   return err;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBCreateLocalVarTable( void )
{
   NPVARDEF NEAR* npvtLocal;
   
   npvtLocal = (NPVARDEF NEAR*)WBLocalAlloc( LPTR, 
                                     sizeof(NPVARDEF) * VAR_TABLE_SIZE );
   if ( npvtLocal == NULL )
      WBRuntimeError( WBERR_OUTOFMEMORY );
   
   g_npContext->npvtLocal = npvtLocal;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBDestroyVarTable( NPVARDEF NEAR* npvt )
{
   NPVARDEF npvar, npvarNext;
   int i;
   
   for ( i = 0; i < VAR_TABLE_SIZE; i++ ) {
      npvar = npvt[i];
      while ( npvar != NULL ) {
         WBDestroyVariantValue( &(npvar->value) );
         npvarNext = npvar->npNext;
         WBLocalFree( (HLOCAL)npvar );
         npvar = npvarNext;
      }
   }
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBDestroyVariantValue( LPVAR lpVar )
{
   if ( lpVar->type == VT_STRING )
      WBDestroyHlstr( lpVar->var.tString );
   else if ( lpVar->type == VT_USERTYPE )
      WBFreeTypeMem( (LPWBTYPEMEM)lpVar->var.tlpVoid );
   else if ( lpVar->type > VT_OBJECT )
      WBUnWireObject( lpVar->var.tCtl );
   
   lpVar->var.tLong = 0;
   lpVar->type = VT_EMPTY;
}

