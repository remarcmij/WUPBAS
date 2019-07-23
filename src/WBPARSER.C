#define STRICT
#include <windows.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <time.h>
#include "wupbas.h"

#define MAXHEXDIGITS       8
#define MAXLONGDIGITS      10
static char szMaxLong[MAXLONGDIGITS+1] = "2147483647";
static char szLong[MAXLONGDIGITS+1];

typedef struct tagRESERVEDWORD {
   NPSTR npszText;
   TOKEN tokID;
   TOKENCLASS tokClass;
} RESERVEDWORD;

RESERVEDWORD ReservedWord2[] = {
   "AS",          TOKEN_AS,      TC_KEYWORD,
   "DO",          TOKEN_DO,      TC_KEYWORD,
   "IF",          TOKEN_IF,      TC_KEYWORD,
   "OR",          TOKEN_OR,      TC_LOP,
   "TO",          TOKEN_TO,      TC_KEYWORD
};
   
RESERVEDWORD ReservedWord3[] = {
   "AND",         TOKEN_AND,     TC_LOP,
   "ANY",         TOKEN_ANY,     TC_KEYWORD,
   "DIM",         TOKEN_DIM,     TC_KEYWORD,
   "END",         TOKEN_END,     TC_KEYWORD,
   "FOR",         TOKEN_FOR,     TC_KEYWORD,
   "LIB",         TOKEN_LIB,     TC_KEYWORD,
   "MOD",         TOKEN_MOD,     TC_MULTOP,
   "NEW",         TOKEN_NEW,     TC_KEYWORD,
   "NOT",         TOKEN_NOT,     TC_KEYWORD,
   "SUB",         TOKEN_SUB,     TC_KEYWORD,
   "XOR",         TOKEN_XOR,     TC_LOP
};
   
RESERVEDWORD ReservedWord4[] = {
   "ELSE",        TOKEN_ELSE,    TC_KEYWORD,
   "EXIT",        TOKEN_EXIT,    TC_KEYWORD,
   "LONG",        TOKEN_LONG,    TC_KEYWORD,
   "LOOP",        TOKEN_LOOP,    TC_KEYWORD,
   "NEXT",        TOKEN_NEXT,    TC_KEYWORD,
   "STEP",        TOKEN_STEP,    TC_KEYWORD,
   "STOP",        TOKEN_STOP,    TC_KEYWORD,
   "THEN",        TOKEN_THEN,    TC_KEYWORD,
   "TRUE",        TOKEN_TRUE,    TC_KEYWORD,
   "TYPE",        TOKEN_TYPE,    TC_KEYWORD,
   "WEND",        TOKEN_WEND,    TC_KEYWORD
};
   
RESERVEDWORD ReservedWord5[] = {
   "ALIAS",       TOKEN_ALIAS,   TC_KEYWORD,
   "BYVAL",       TOKEN_BYVAL,   TC_KEYWORD,
   "CONST",       TOKEN_CONST,   TC_KEYWORD,
   "ENDIF",       TOKEN_ENDIF,   TC_KEYWORD,
   "FALSE",       TOKEN_FALSE,   TC_KEYWORD,
   "LOCAL",       TOKEN_LOCAL,   TC_KEYWORD,
   "UNTIL",       TOKEN_UNTIL,   TC_KEYWORD,
   "WHILE",       TOKEN_WHILE,   TC_KEYWORD
};

RESERVEDWORD ReservedWord6[] = {
   "DECLARE",     TOKEN_DECLARE,       TC_KEYWORD,
   "ELSEIF",      TOKEN_ELSEIF,        TC_KEYWORD,
   "ENDFUNCTION", TOKEN_ENDFUNCTION,   TC_KEYWORD,
   "ENDSUB",      TOKEN_ENDSUB,        TC_KEYWORD,
   "ENDTYPE",     TOKEN_ENDTYPE,       TC_KEYWORD,
   "EXITDO",      TOKEN_EXITDO,        TC_KEYWORD,
   "EXITFOR",     TOKEN_EXITFOR,       TC_KEYWORD,
   "EXITFUNCTION",TOKEN_EXITFUNCTION,  TC_KEYWORD,
   "EXITSUB",     TOKEN_EXITSUB,       TC_KEYWORD,
   "FUNCTION",    TOKEN_FUNCTION,      TC_KEYWORD,
   "GLOBAL",      TOKEN_GLOBAL,        TC_KEYWORD,
   "INTEGER",     TOKEN_INTEGER,       TC_KEYWORD,
   "LOADDLL",     TOKEN_LOADDLL,       TC_KEYWORD,
   "LOADLIB",     TOKEN_LOADLIB,       TC_KEYWORD,
   "OPTION",      TOKEN_OPTION,        TC_KEYWORD,
   "VARIANT",     TOKEN_VARIANT,       TC_KEYWORD
};

// Words that may follow END
RESERVEDWORD ReservedWord7[] = {
   "FUNCTION",    TOKEN_ENDFUNCTION,   TC_KEYWORD,
   "IF",          TOKEN_ENDIF,         TC_KEYWORD,
   "SUB",         TOKEN_ENDSUB,        TC_KEYWORD,
   "TYPE",        TOKEN_ENDTYPE,       TC_KEYWORD
};
   

typedef struct tagCONSTANTIDENT {
   NPSTR npszText;
   long lValue;
} CONSTANTIDENT;

CONSTANTIDENT ConstantTable[] = {
   "ABORT",                3,    // Abort button pressed
   "APPEND",               2,    // Open APPEND
   "ATTR_ARCHIVE",         32,
   "ATTR_DIRECTORY",       16,
   "ATTR_HIDDEN",          2,
   "ATTR_NORMAL",          0,
   "ATTR_READONLY",        1,
   "ATTR_SYSTEM",          4,
   "ATTR_VOLUME",          8,
   "BINARY",               3,    // Open BINARY
   "BIT0",                 0x00000001L,
   "BIT1",                 0x00000002L,
   "BIT2",                 0x00000004L,
   "BIT3",                 0x00000008L,
   "BIT4",                 0x00000010L,
   "BIT5",                 0x00000020L,
   "BIT6",                 0x00000040L,
   "BIT7",                 0x00000080L,
   "BIT8",                 0x00000100L,
   "BIT9",                 0x00000200L,
   "BIT10",                0x00000400L,
   "BIT11",                0x00000800L,
   "BIT12",                0x00001000L,
   "BIT13",                0x00002000L,
   "BIT14",                0x00004000L,
   "BIT15",                0x00008000L,
   "BIT16",                0x00010000L,
   "BIT17",                0x00020000L,
   "BIT18",                0x00040000L,
   "BIT19",                0x00080000L,
   "BIT20",                0x00100000L,
   "BIT21",                0x00200000L,
   "BIT22",                0x00400000L,
   "BIT23",                0x00800000L,
   "BIT24",                0x01000000L,
   "BIT25",                0x02000000L,
   "BIT26",                0x04000000L,
   "BIT27",                0x08000000L,
   "BIT28",                0x10000000L,
   "BIT29",                0x20000000L,
   "BIT30",                0x40000000L,
   "BIT31",                0x80000000L,
   "CANCEL",               2,    // Cancel button pressed
   "CPU286",               2,
   "CPU386",               4,
   "CPU486",               8,
   "DRIVE_FIXED",          3,
   "DRIVE_REMOTE",         4,
   "DRIVE_REMOVABLE",      2,
   "EXITWINDOWS",          0,
   "FALSE",                0,
   "IDABORT",              3,    // Abort button pressed
   "IDCANCEL",             2,    // Cancel button pressed
   "IDIGNORE",             5,    // Ignore button pressed
   "IDNO",                 7,    // No button pressed
   "IDOK",                 1,    // OK button pressed
   "IDRETRY",              4,    // Retry button pressed
   "IDYES",                6,    // Yes button pressed
   "IGNORE",               5,    // Ignore button pressed
   "INPUT",                0,    // Open INPUT
   "MB_ABORTRETRYIGNORE",  2,    // Abort, Retry, and Ignore buttons
   "MB_APPLMODAL",         0,    // Application Modal Message Box
   "MB_DEFBUTTON1",        0,    // First button is default
   "MB_DEFBUTTON2",        256,  // Second button is default
   "MB_DEFBUTTON3",        512,  // Third button is default
   "MB_ICONEXCLAMATION",   48,   // Warning message
   "MB_ICONINFORMATION",   64,   // Information message
   "MB_ICONQUESTION",      32,   // Warning query
   "MB_ICONSTOP",          16,   // Critical message
   "MB_OK",                0,    // OK button only
   "MB_OKCANCEL",          1,    // OK and Cancel buttons
   "MB_RETRYCANCEL",       5,    // Retry and Cancel buttons
   "MB_SYSTEMMODAL",       4096, // System Modal
   "MB_YESNO",             4,    // Yes and No buttons
   "MB_YESNOCANCEL",       3,    // Yes, No, and Cancel buttons
   "MINIMIZE",             6,
   "NO",                   7,    // No button pressed
   "NOWAIT",               0,
   "OK",                   1,    // OK button pressed
   "OUTPUT",               1,    // Open OUTPUT
   "REBOOTSYSTEM",         0x43, // EW_REBOOTSYSTEM (ExitWindows)
   "RESTARTWINDOWS",       0x42, // EW_RESTARTWINDOWS (ExitWindows)
   "RESTORE",              9,
   "RETRY",                4,    // Retry button pressed
   "SHOW",                 5,
   "SHOWMAXIMIZED",        3,
   "SHOWMINIMIZED",        2,
   "SHOWMINNOACTIVE",      7,
   "SHOWNA",               8,
   "SHOWNOACTIVATE",       4,
   "SHOWNORMAL",           1,
   
   "SM_CXBORDER",          5,    // GetSystemMetrics manifest constants
   "SM_CXCURSOR",          13,
   "SM_CXDLGFRAME",        7,
   "SM_CXDOUBLECLK",       36,
   "SM_CXFRAME",           32,
   "SM_CXFULLSCREEN",      16,
   "SM_CXHSCROLL",         21,
   "SM_CXHTHUMB",          10,
   "SM_CXICON",            11,
   "SM_CXICONSPACING",     38,
   "SM_CXMIN",             28,
   "SM_CXMINTRACK",        34,
   "SM_CXSCREEN",          0,
   "SM_CXSIZE",            30,
   "SM_CXVSCROLL",         2,
   "SM_CYBORDER",          6,
   "SM_CYCAPTION",         4,
   "SM_CYCURSOR",          14,
   "SM_CYDLGFRAME",        8,
   "SM_CYDOUBLECLK",       37,
   "SM_CYFRAME",           33,
   "SM_CYFULLSCREEN",      17,
   "SM_CYHSCROLL",         3,
   "SM_CYICON",            12,
   "SM_CYICONSPACING",     39,
   "SM_CYKANJIWINDOW",     18,
   "SM_CYMENU",            15,
   "SM_CYMIN",             29,
   "SM_CYMINTRACK",        35,
   "SM_CYSCREEN",          1,
   "SM_CYSIZE",            31,
   "SM_CYVSCROLL",         20,
   "SM_CYVTHUMB",          9,
   "SM_DBCSENABLED",       42,
   "SM_DEBUG",             22,
   "SM_MENUDROPALIGNMENT", 40,
   "SM_MOUSEPRESENT",      19,
   "SM_PENWINDOWS",        41,
   "SM_SWAPBUTTON",        23,
   
   "TRUE",                 -1,
   "V_DATE",               7,    // VarType
   "V_EMPTY",              0,    //   ""
   "V_FILE",               258,  //   ""
   "V_LONG",               3,    //   ""
   "V_OBJECT",             256,  //   ""
   "V_SLIST",              257,  //   ""
   "V_STRING",             8,    //   ""
   "WAIT",                 1,
   "YES",                  6     // Yes button pressed
};   

typedef struct tagWBMONTHINFO {
   NPSTR npszText;
   int nDays;
} WBMONTHINFO;

WBMONTHINFO aMonthTable[] = {
   "JANUARY",     31,
   "FEBRUARY",    29,
   "MARCH",       31,
   "APRIL",       30,
   "MAY",         31,
   "JUNE",        30,
   "JULY",        31,
   "AUGUST",      31,
   "SEPTEMBER",   30,
   "OCTOBER",     31,
   "NOVEMBER",    30,
   "DECEMBER",    31     
};

// The scanner supports two modes. In the normal mode, indicated by fBasicLang
// boolean (in scanner context structure) being TRUE the scanner produces 
// Basic language tokens (e.g. including Basic reserved words). When 
// fBasicLang is FALSE no consideration is given to Basic keywords. 


static int WBMonthNameToNumber( LPSTR lpszText );
      
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
TOKEN NextToken( void )
{
   int i, j, k, nComp;
   register NPSTR npchNextPos;

   npchNextPos = g_npContext->npchPrevPos = g_npContext->npchNextPos;
   if ( g_npContext->fEol ) {
      g_npContext->lpLine = WBNextLine( g_npContext->szProgText );
      if ( g_npContext->lpLine == NULL ) {
         g_npContext->tokType = TOKEN_EOF;
         return TOKEN_EOF;
      }
      g_npContext->fEol = FALSE;
      npchNextPos = g_npContext->npchPrevPos = g_npContext->szProgText;
   }
   
   // Skip white space
   g_npContext->nLeadingSpaces = 0;
   while ( isspace( *npchNextPos ) ) {
      npchNextPos++;   
      g_npContext->nLeadingSpaces++;
   }
      
   /*-----------------------------------------------------------------------*/   
   /* Identifier                                                            */
   /*-----------------------------------------------------------------------*/   
   if ( __iscsymf( *npchNextPos )  ) {
      register NPSTR npch;
      NPSTR npch1;
      int nComp, cItems;
      RESERVEDWORD NEAR* npTable;
   
      g_npContext->tokType = TOKEN_IDENT;
      g_npContext->tokClass = TC_USERIDENT;
      
      
      npch = npch1 = g_npContext->szToken;
      
      // Copy keyword
      while ( __iscsym( *npchNextPos) ) {
         *npch++ = (char)toupper( *npchNextPos );
         npchNextPos++;
      }
      // Zero terminate (and truncate if necessary)         
      *npch = *(npch1 + MAXTOKEN - 1) = '\0';
      
      npch = g_npContext->szToken;
  
      switch ( strlen( npch ) ) {
         case 1:
            cItems = 0;
            npTable = NULL;   // To avoid compiler warning
            break;
         case 2:
            cItems = sizeof(ReservedWord2) / sizeof(RESERVEDWORD);
            npTable = ReservedWord2;
            break;
         case 3:
            cItems = sizeof(ReservedWord3) / sizeof(RESERVEDWORD);
            npTable = ReservedWord3;
            break;
         case 4:
            cItems = sizeof(ReservedWord4) / sizeof(RESERVEDWORD);
            npTable = ReservedWord4;
            break;
         case 5:
            cItems = sizeof(ReservedWord5) / sizeof(RESERVEDWORD);
            npTable = ReservedWord5;
            break;
         default:
            cItems = sizeof(ReservedWord6) / sizeof(RESERVEDWORD);
            npTable = ReservedWord6;
      }
         
      // do binary search on reserved keyword table
      if ( cItems > 0 ) {
         register int i, j, k;
         i = 0; 
         j = cItems;
         do {
            k = ( i + j ) / 2;
            if ( (nComp = strcmp( npch, npTable[k].npszText )) > 0 )
               i = k + 1;
            else
               j = k - 1;
         } while ( nComp != 0 && i <= j );
                  
         if ( nComp == 0 ) {
         
            g_npContext->tokType = npTable[k].tokID;
            g_npContext->tokClass = npTable[k].tokClass;

            // For TOKEN_END, check if it ends a block structure            
            if ( g_npContext->tokType == TOKEN_END ) {
               NPSTR npch1, npch2;
               char szKeyword[16];
               int len;
               
               npch1 = npchNextPos;
               while ( isspace( *npch1 ) )
                  npch1++;
                  
               npch2 = npch1;
               while ( isalpha( *npch2 ) )
                  npch2++;
               
               len = npch2 - npch1;    
               
               if ( len > 0 && len < sizeof( szKeyword ) ) {
                  memcpy( szKeyword, npch1, len );
                  szKeyword[len] = '\0';
                  _strupr( szKeyword );
                  
                  for ( i = 0; 
                        i < sizeof(ReservedWord7) / sizeof(RESERVEDWORD);
                        i++ ) {
                     if ((j = strcmp(szKeyword, ReservedWord7[i].npszText)) == 0) {
                        g_npContext->tokType = ReservedWord7[i].tokID;
                        g_npContext->tokClass = ReservedWord7[i].tokClass;
                        npchNextPos = npch2;
                        break;
                     }
                     else if ( j < 0 )
                        break;
                  }
               }
            }
         }
      }
   }

   /*-----------------------------------------------------------------------*/   
   /* Integer constant                                                      */
   /*-----------------------------------------------------------------------*/   
   else if ( isdigit( *npchNextPos ) ) {
      BOOL fHexConstant = FALSE;
      LONG lValue = 0;
      
      g_npContext->tokType = TOKEN_LONGINT;
      g_npContext->tokClass = TC_NUMCONST;

      // test for hex constant of form '0x.....'
      if ( *npchNextPos == '0' ) {
         npchNextPos++;
         if ( *npchNextPos == 'x' || *npchNextPos == 'X' ) {
            fHexConstant = TRUE;
            npchNextPos++;
            lValue = 0;
            
            for ( i = 0; 
                  i < MAXHEXDIGITS && isxdigit( *npchNextPos );
                  i++, npchNextPos++ ) {
                  
               if ( *npchNextPos >= '0' && *npchNextPos <= '9' )
                  lValue = (lValue << 4) | (long)(*npchNextPos - '0');
               else
                  lValue = (lValue << 4) | 
                               (long)(toupper(*npchNextPos) - 'A' + 10 );
            }

            if ( isalnum( *npchNextPos ) ) {
               if ( isxdigit( *npchNextPos ) )
                  WBRuntimeError( WBERR_NUMBERTOOBIG );
               else 
                  WBRuntimeError( WBERR_HEXCONSTANT );
            }
         }
      }
      
      if ( !fHexConstant ) {
         // decimal constant
         for ( i = 0; i < MAXLONGDIGITS && isdigit( *npchNextPos ); i++ )
            szLong[i] = *npchNextPos++;
               
         if ( i == 1 )
            lValue = szLong[0] - '0';
         else {
            szLong[i] = '\0';
            if ( isdigit( *npchNextPos ) ||
                 (i == MAXLONGDIGITS && strcmp( szLong, szMaxLong ) > 0) ) {
                  WBRuntimeError( WBERR_NUMBERTOOBIG );
            }
            lValue = atol( szLong );
         }
      }
      
      g_npContext->lTokenVal = lValue;
   }
   
   /*-----------------------------------------------------------------------*/   
   /* Hex constant of form &Hxxxx                                           */
   /*-----------------------------------------------------------------------*/   
   else if ( *npchNextPos == '&' ) {
      npchNextPos++;
      
      if ( *npchNextPos == 'h'|| *npchNextPos == 'H' ) {
         LONG lValue = 0;
      
         g_npContext->tokType = TOKEN_LONGINT;
         g_npContext->tokClass = TC_NUMCONST;
         
         npchNextPos++;
         lValue = 0;
            
         for ( i = 0; 
               i < MAXHEXDIGITS && isxdigit( *npchNextPos );
               i++, npchNextPos++ ) {
                  
            if ( *npchNextPos >= '0' && *npchNextPos <= '9' )
               lValue = (lValue << 4) | (long)(*npchNextPos - '0');
            else
               lValue = (lValue << 4) | 
                            (long)(toupper(*npchNextPos) - 'A' + 10 );
         }

         if ( isalnum( *npchNextPos ) ) {
            if ( isxdigit( *npchNextPos ) )
               WBRuntimeError( WBERR_NUMBERTOOBIG );
            else 
               WBRuntimeError( WBERR_HEXCONSTANT );
         }
         
         g_npContext->lTokenVal = lValue;
      }
      else {
         g_npContext->tokType = TOKEN_CONCAT;
         g_npContext->tokClass = TC_ADDOP;
      }
   }
   
   /*-----------------------------------------------------------------------*/   
   /* String literal                                                        */
   /*-----------------------------------------------------------------------*/   
   else if ( *npchNextPos == '"' || *npchNextPos == '`' ) {
      NPSTR pch;
      char chQuote;

      chQuote = *npchNextPos++;      
      
      g_npContext->tokType = TOKEN_STRINGLIT;
      g_npContext->tokClass = TC_STRINGCONST;
      
      pch = g_npContext->szToken;
      for ( i = 0; i < MAXTEXT-1 && *npchNextPos && *npchNextPos != chQuote; i++ ) 
         *pch++ = *npchNextPos++;
      *pch = '\0';
      
      if ( *npchNextPos != chQuote  ) 
         WBRuntimeError( WBERR_STRINGLIT );
         
      npchNextPos++;
   }
   
   /*-----------------------------------------------------------------------*/   
   /* Predefined constant                                                   */
   /*-----------------------------------------------------------------------*/   
   else if ( *npchNextPos == '@' ) {
      NPSTR pch;
   
      g_npContext->tokType = TOKEN_LONGINT;
      g_npContext->tokClass = TC_SYMBOLCONST;
      
      npchNextPos++;
      pch = g_npContext->szToken;
      
      // Copy keyword (truncate if necessary)
      for ( i = 0; 
            i < MAXTOKEN-1 && ( isalnum( *npchNextPos) || *npchNextPos == '_');
            i++ ) {
         *pch++ = *npchNextPos++;
      }
         
      *pch = '\0';
      _strupr( g_npContext->szToken );
      
      // Eat up (and ignore) further eligible characters
      while ( isalnum( *npchNextPos ) || *npchNextPos == '_' )
         npchNextPos++;

      // do binary search on reserved keyword table
      i = 0; 
      j = sizeof(ConstantTable) / sizeof(CONSTANTIDENT) - 1;
      pch = g_npContext->szToken;
      do {
         k = ( i + j ) / 2;
         if ( (nComp = strcmp( pch, ConstantTable[k].npszText )) > 0 )
            i = k + 1;
         else
            j = k - 1;
      } while ( nComp != 0 && i <= j );
      
      if ( nComp == 0 ) 
         g_npContext->lTokenVal = ConstantTable[k].lValue;
      else 
         g_npContext->tokType = TOKEN_UNKNOWN;
   }
   
   /*-----------------------------------------------------------------------*/   
   /* Non-alphanumerics                                                     */
   /*-----------------------------------------------------------------------*/   
   else {
   
      g_npContext->tokClass = TC_OTHER;
      
      switch ( *npchNextPos ) {
      
         case ',':
            g_npContext->tokType = TOKEN_COMMA;
            npchNextPos++;
            break;
            
         case '=':
            g_npContext->tokType = TOKEN_EQU;
            g_npContext->tokClass = TC_RELOP;
            npchNextPos++;
            break;
            
         case '(':
            g_npContext->tokType = TOKEN_LPAREN;
            npchNextPos++;
            break;
            
         case ')':
            g_npContext->tokType = TOKEN_RPAREN;
            npchNextPos++;
            break;
            
         case '.':
            g_npContext->tokType = TOKEN_PERIOD;
            npchNextPos++;
            break;
            
         case '\0':  // physical end of line
            g_npContext->tokType = TOKEN_EOL;
            g_npContext->fEol = TRUE;
            break;
            
         case '\'':  // comment introducer
         case ';':   // comment introducer
            // start of comment is equivalent to end of line
            g_npContext->tokType = TOKEN_EOL;
            // move pointer to end of line
            npchNextPos = npchNextPos + strlen( npchNextPos );
            break;
            
         case '+':
            g_npContext->tokType = TOKEN_PLUS;
            g_npContext->tokClass = TC_ADDOP;
            npchNextPos++;
            break;
            
         case '-':
            g_npContext->tokType = TOKEN_MINUS;
            g_npContext->tokClass = TC_ADDOP;
            npchNextPos++;
            break;
            
         case '*':
            g_npContext->tokType = TOKEN_MULT;
            g_npContext->tokClass = TC_MULTOP;
            npchNextPos++;
            break;
            
         case '/':
            g_npContext->tokType = TOKEN_DIV;
            g_npContext->tokClass = TC_MULTOP;
            npchNextPos++;
            break;
            
         case '<':
            npchNextPos++;
            if ( *npchNextPos == '=' ) {
               g_npContext->tokType = TOKEN_LE;
               g_npContext->tokClass = TC_RELOP;
               npchNextPos++;
            }
            else if ( *npchNextPos == '>' ) {
               g_npContext->tokType = TOKEN_NE;
               g_npContext->tokClass = TC_RELOP;
               npchNextPos++;
            }
            else if ( *npchNextPos == '<' ) {
               g_npContext->tokType = TOKEN_SHIFTLEFT;
               g_npContext->tokClass = TC_SHIFTOP;
               npchNextPos++;
            }
            else {
               g_npContext->tokType = TOKEN_LT;
               g_npContext->tokClass = TC_RELOP;
            }
            break;
            
         case '>':
            npchNextPos++;
            if ( *npchNextPos == '=' ) {
               g_npContext->tokType = TOKEN_GE;
               g_npContext->tokClass = TC_RELOP;
               npchNextPos++;
            }
            else if ( *npchNextPos == '>' ) {
               g_npContext->tokType = TOKEN_SHIFTRIGHT;
               g_npContext->tokClass = TC_SHIFTOP;
               npchNextPos++;
            }
            else  {
               g_npContext->tokType = TOKEN_GT;
               g_npContext->tokClass = TC_RELOP;
            }
            break;
            
         case '#':
            g_npContext->tokType = TOKEN_HASHSIGN;
            npchNextPos++;
            break;
            
         case ':':
            g_npContext->tokType = TOKEN_COLON;
            npchNextPos++;
            break;
            
         case '%':
            g_npContext->tokType = TOKEN_PERCENT;
            npchNextPos++;
            break;
            
         case '$':
            g_npContext->tokType = TOKEN_DOLLAR;
            npchNextPos++;
            break;
            
         default:
            g_npContext->tokType = TOKEN_UNDEFINED;
            npchNextPos++;
      }
   }

   g_npContext->npchNextPos = npchNextPos;
   return g_npContext->tokType;
}
      
/*--------------------------------------------------------------------------*/
/* Looks for a subset of symbols only                                       */
/*--------------------------------------------------------------------------*/
TOKEN PeekToken( void )
{
   NPSTR npch;
   
   npch = g_npContext->npchNextPos;
   while ( isspace( *npch ) )
      npch++;
      
   switch ( *npch ) {
   
      case '=':
         return TOKEN_EQU;
         
      case '(':
         return TOKEN_LPAREN;
         
      case ',':
         return TOKEN_COMMA;
         
      case '.':
         return TOKEN_PERIOD;
         
      default:
         return TOKEN_UNKNOWN;
   }
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
LPVOID GetCodePointer( void )
{
   return g_npContext->lpLine;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void SkipLine( void )
{
   g_npContext->fEol = TRUE;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void ResetLine( void )
{
   g_npContext->npchNextPos = g_npContext->szProgText;
   g_npContext->fEol = FALSE;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBParseDateTimeString( LPSTR lpszDatimString, LPVAR lpResult )
{
   TOKEN tok;
   long lDay, lMonth, lYear;
   long lHour, lMinute, lSecond;
   long lTemp1, lTemp2, lTemp3;
   struct tm *ptm;
   struct tm tim;
   time_t ltime;
   ERR err;

   if ( WBPushContext() == NULL )
      return WBERR_OUTOFMEMORY;
      
   strcpy( g_npContext->szProgText, lpszDatimString );
   g_npContext->npchNextPos = g_npContext->szProgText;
      
   err = 0;
   
   // All date/times are relative to 1-Jan-1970 as dictated by the standard
   // C runtime library functions.
   lDay   = 1;
   lMonth = 1;
   lYear  = 70;
   lHour  = lMinute = lSecond = lTemp1 = lTemp2 = lTemp3 = 0;

   // Get first token   
   tok = NextToken();   
   
   if ( tok == TOKEN_LONGINT ) {
   
      lTemp1 = g_npContext->lTokenVal;
      tok = NextToken();
      if ( tok == TOKEN_COLON )
         goto parse_time;
         
      if ( tok == TOKEN_DIV || tok == TOKEN_MINUS || tok == TOKEN_COMMA )
         tok = NextToken();
         
      if ( tok == TOKEN_LONGINT ) {
         
         lTemp2 = g_npContext->lTokenVal;
         
         if ( tok == TOKEN_DIV || tok == TOKEN_MINUS || tok == TOKEN_COMMA )
            tok = NextToken();
            
         if ( tok != TOKEN_LONGINT ) {
            err = WBERR_DATETIMELIT;
         }
         else {
            
            lTemp3 = g_npContext->lTokenVal;
            
            switch ( g_npTask->fuDateOrder ) {
               
               case DO_DMY:
                  lDay = lTemp1;
                  lMonth = lTemp2;
                  lYear = lTemp3;
                  break;
                  
               case DO_MDY:
                  lMonth = lTemp1;
                  lDay = lTemp2;
                  lYear = lTemp3;
                  break;
                  
               case DO_YMD:
                  lYear = lTemp1;
                  lMonth = lTemp2;
                  lDay = lTemp3;
                  break;
                  
               default:
                  // Can only have one of the above three cases
                  assert( FALSE );
                  return WBERR_INTERNAL_ERROR;
            }
            
            tok = NextToken();
            if ( tok == TOKEN_LONGINT ) {
               lTemp1 = g_npContext->lTokenVal;
               tok = NextToken();
            }
            else if ( tok != TOKEN_EOL ) {
               err = WBERR_DATETIMELIT;
            }
         }
      }         
      else if ( tok == TOKEN_IDENT ) {
      
         // Since we have a word as the second item, the current date format
         // must be DMY or YMD
         
         if ( g_npTask->fuDateOrder != DO_DMY &&
              g_npTask->fuDateOrder != DO_YMD ) {
            err = WBERR_DATETIMELIT;
         }
         else {
            lMonth = (long)WBMonthNameToNumber( g_npContext->szToken );
            if ( lMonth == 0 ) {
               err = WBERR_DATETIMELIT;
            }
            else {
               if ( tok == TOKEN_DIV || tok == TOKEN_MINUS || tok == TOKEN_COMMA )
                  tok = NextToken();
                  
               if ( tok != TOKEN_LONGINT ) {
                  err = WBERR_DATETIMELIT;
               }
               else {
                  lTemp3 = g_npContext->lTokenVal;
                  
                  switch ( g_npTask->fuDateOrder ) {
                  
                     case DO_DMY:
                        lDay = lTemp1;
                        lYear = lTemp3;
                        break;
                        
                     case DO_YMD:
                        lDay = lTemp3;
                        lYear = lTemp1;
                        break;
                        
                     default:
                        assert( FALSE );
                        return WBERR_INTERNAL_ERROR;
                  }
         
                  tok = NextToken();
                  if ( tok == TOKEN_LONGINT ) {
                     lTemp1 = g_npContext->lTokenVal;
                     tok = NextToken();
                  }
                  else if ( tok != TOKEN_EOL ) {
                     err = WBERR_DATETIMELIT;
                  }
               }
            }   
         }
      }
      else {
         err = WBERR_DATETIMELIT;
      }
   }
   else if ( tok == TOKEN_IDENT ) {
   
      // Since we got a month name as the first item, the current date
      // format must be MDY for this to be valid
      if ( g_npTask->fuDateOrder != DO_MDY ) {
         err = WBERR_DATETIMELIT;
      }
      else {
         
         lMonth = (long)WBMonthNameToNumber( g_npContext->szToken );
         if ( lMonth == 0 ) {
            err = WBERR_DATETIMELIT;
         }
         else {
            if ( tok == TOKEN_DIV || tok == TOKEN_MINUS || tok == TOKEN_COMMA )
               tok = NextToken();
               
            if ( tok != TOKEN_LONGINT  ) {
               err = WBERR_DATETIMELIT;
            }
            else {
                     
               lDay = g_npContext->lTokenVal;
               
               if ( tok == TOKEN_DIV || tok == TOKEN_MINUS || tok == TOKEN_COMMA )
                  tok = NextToken();
                  
               if ( tok == TOKEN_DIV || tok == TOKEN_MINUS || tok == TOKEN_COMMA )
                  tok = NextToken();
                  
               if ( tok != TOKEN_LONGINT ) {
                  err = WBERR_DATETIMELIT;
               }
               else {
                  
                  lYear = g_npContext->lTokenVal;
                  tok = NextToken();
                  if ( tok == TOKEN_LONGINT ) {
                     lTemp1 = g_npContext->lTokenVal;
                     tok = NextToken();
                  }
                  else if ( tok != TOKEN_EOL ) {
                     err = WBERR_DATETIMELIT;
                  }
               }
            }
         }
      }
   }
   else {
      err = WBERR_DATETIMELIT;
   }

   // Validate day, year and month values
   if ( err == 0 ) {
      
      // If no century number given, add the current century number
      if ( lYear < 100L ) {
         // get current time
         time( &ltime );
         ptm = localtime( &ltime );
         // add current century number
         lYear += (((long)(ptm->tm_year) / 100L) * 100L) + 1900L;
      }
            
      // The year must be 1900 or later and no later than
      // 2014 (the retirement date of both the author and this program :-) 
      if ( lYear < 1900L && lYear > 2014L ) {
         err = WBERR_DATETIMELIT;
      }
      else {
         // Normalize back to base-1900            
         lYear -= 1900L;
                  
         if ( lMonth > 12L || lDay > (long)(aMonthTable[lMonth-1].nDays) ) {
            err = WBERR_DATETIMELIT;
         }
      }
   }
      

parse_time:

   if ( err == 0 && tok != TOKEN_EOL ) {
      // At this point we have a valid date. Now, parse the time part. We
      // already have the hour value in lTemp1 and the current token should
      // be the time separator character.
      lHour = lTemp1;

      if ( tok != TOKEN_COLON ) {
         err = WBERR_DATETIMELIT;
      }
      else {
         tok = NextToken();
         if ( tok != TOKEN_LONGINT ) {
            err = WBERR_DATETIMELIT;
         }
         else {
         
            lMinute = g_npContext->lTokenVal;
            tok = NextToken();
   
            // Seconds part is optional
            if ( tok != TOKEN_EOL ) {
               
               if ( tok != TOKEN_COLON ) {
                  err = WBERR_DATETIMELIT;
               }
               else {
                  
                  tok = NextToken();
                  if ( tok != TOKEN_LONGINT )
                     err = WBERR_DATETIMELIT;
                  
                  lSecond = g_npContext->lTokenVal;
                  tok = NextToken();
               }
      
               if ( tok != TOKEN_EOL ) {
                  err = WBERR_DATETIMELIT;
               }
            }
         }
      }
      
      // Validate time components
      if ( lHour > 23L || lMinute > 59L || lSecond > 59L ) {
         err = WBERR_DATETIMELIT;
      }
   }

   if ( err == 0 ) {
      
      // convert date/time components to a standard C time_t value
      memset( &tim, 0, sizeof( tim ) );
      tim.tm_mday = (int)lDay;
      tim.tm_mon  = (int)lMonth - 1;
      tim.tm_year = (int)lYear;
      tim.tm_hour = (int)lHour;
      tim.tm_min  = (int)lMinute;
      tim.tm_sec  = (int)lSecond;
      tim.tm_isdst = -1;   // Do not use day light saving feature
                     
      ltime = mktime( &tim );
      if ( ltime == (time_t)-1L ) 
         err = WBERR_DATETIMELIT;
      else {
         if ( lpResult != NULL ) {
            lpResult->var.tLong = (long)ltime;
            lpResult->type = VT_DATE;
         }
      }
   }

   WBPopContext();   
   
   return err;
}            
      
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
static int WBMonthNameToNumber( LPSTR lpszText )
{
   int i, nLen;
   
   // minimum abbreviation of month name is 3 letters
   nLen = max( strlen( lpszText ), 3 );
        
   for ( i = 0; i < 12; i++ ) {
      if ( strncmp( lpszText, aMonthTable[i].npszText, nLen ) == 0 )
         return i + 1;
   }
   
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
BOOL IsTermToken( TOKEN tok, WORD wTerm )
{
   switch ( tok ) {
      
      case TOKEN_ENDIF:
         return ( wTerm & TERM_ENDIF ) == TERM_ENDIF;

      case TOKEN_ELSE:
         return ( wTerm & TERM_ELSE ) == TERM_ELSE;
         
      case TOKEN_ELSEIF:
         return ( wTerm & TERM_ELSEIF ) == TERM_ELSEIF;
         
      case TOKEN_NEXT:
         return ( wTerm & TERM_NEXT ) == TERM_NEXT;
       
      case TOKEN_LOOP:
         return ( wTerm & TERM_LOOP ) == TERM_LOOP;
         
      case TOKEN_WEND:
         return ( wTerm & TERM_WEND ) == TERM_WEND;
         
      default:
         return FALSE;
   }
}
