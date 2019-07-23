/* wupbas.h */

#ifndef RC_INVOKED
/* 4001: Allow single line comments */
/* 4045: Do not warn about inline's being removed */
/* 4704: Allow inline assembly      */
#pragma warning( disable: 4001 4505 4704 )
#endif  /* RC_INVOKED */

#include <stdio.h>
#include <assert.h>
#include <dos.h>
#include "vbapisub.h"
#include "resource.h"

#define UNREFERENCED_PARAM(x)  ((void)(x))

#define MAXTOKEN     40    // Max 39 characters in an identifier
#define MAXLONG      12    // Max characters to represent a long signed int
                           // including the terminating null character
#define MAXTEXT      256
#define MAXSTRING    65500

#define MAXARGS      32    // Maximum number of arguments that can be
                           // passed to a user-defined function

#define WBERR_INITIALIZATION        100
#define WBERR_TASK_NOT_REGISTERED   101

#define VT_OBJECT       256
#define VT_SLIST        VT_OBJECT+1
#define VT_FILE         VT_OBJECT+2

typedef enum tagTOKEN {
   TOKEN_ALIAS,
   TOKEN_AND,
   TOKEN_ANY,
   TOKEN_AS, 
   TOKEN_BYVAL,
   TOKEN_COLON,
   TOKEN_COMMA,
   TOKEN_CONCAT,
   TOKEN_CONST,
   TOKEN_DIM,
   TOKEN_DIV,
   TOKEN_DATETIMELIT,
   TOKEN_DECLARE,
   TOKEN_DO,
   TOKEN_DOLLAR,
   TOKEN_ELSE,
   TOKEN_ELSEIF,
   TOKEN_END,
   TOKEN_ENDFUNCTION,
   TOKEN_ENDIF,
   TOKEN_ENDSUB,
   TOKEN_ENDTYPE,
   TOKEN_EOF,
   TOKEN_EOL,
   TOKEN_EQU,
   TOKEN_EXIT,
   TOKEN_EXITDO,
   TOKEN_EXITFOR,
   TOKEN_EXITFUNCTION,
   TOKEN_EXITSUB,
   TOKEN_FALSE,
   TOKEN_FOR,
   TOKEN_FUNCTION,
   TOKEN_GE,
   TOKEN_GLOBAL,
   TOKEN_GT,
   TOKEN_HASHSIGN,
   TOKEN_IDENT,
   TOKEN_INTEGER,
   TOKEN_IF,
   TOKEN_LE,
   TOKEN_LIB,
   TOKEN_LOADDLL,
   TOKEN_LOADLIB,
   TOKEN_LOCAL,
   TOKEN_LONG,
   TOKEN_LONGINT,
   TOKEN_LOOP,
   TOKEN_LPAREN,
   TOKEN_LT,
   TOKEN_MINUS,
   TOKEN_MOD,
   TOKEN_MULT,
   TOKEN_NE,
   TOKEN_NEW,
   TOKEN_NEXT,
   TOKEN_NOT,
   TOKEN_OPTION,
   TOKEN_OR,
   TOKEN_PERCENT,
   TOKEN_PERIOD,
   TOKEN_PLUS,
   TOKEN_PRINT,
   TOKEN_RPAREN,
   TOKEN_SHIFTLEFT,
   TOKEN_SHIFTRIGHT,
   TOKEN_STEP,
   TOKEN_STOP,
   TOKEN_STRING,
   TOKEN_STRINGLIT,
   TOKEN_SUB,
   TOKEN_THEN,
   TOKEN_TO,
   TOKEN_TRUE,
   TOKEN_TYPE,
   TOKEN_UNDEFINED,
   TOKEN_UNKNOWN,
   TOKEN_UNTIL,
   TOKEN_VARIANT,
   TOKEN_WEND,
   TOKEN_WHILE,
   TOKEN_XOR
} TOKEN;

typedef enum tagTOKENCLASS {
   TC_USERIDENT,
   TC_KEYWORD,
   TC_NUMCONST,
   TC_SYMBOLCONST,
   TC_STRINGCONST,
   TC_ADDOP,            // Addition operators (+, -)
   TC_MULTOP,           // Multiplication operator (*, /, MOD)
   TC_LOP,              // Dyadic logical operators (AND, OR, XOR) but not NOT
   TC_RELOP,            // Relational operators (=, <, >, <=, >=, <>)
   TC_SHIFTOP,          // Shift operators ( <<, >> )
   TC_OTHER
} TOKENCLASS;

#define TERM_ENDIF   0x0001
#define TERM_ELSE    0x0002
#define TERM_ELSEIF  0x0004
#define TERM_LOOP    0x0008
#define TERM_NEXT    0x0010
#define TERM_WEND    0x0020

#define VAR_TABLE_SIZE      23
#define DECL_TABLE_SIZE     97

// Manifest constants for variable attributes
#define VA_ARGBYVAL     0x0001
#define VA_ARGBYREF     0x0002
#define VA_CONST        0x0004
#define VA_VARIANT      0x0008
#define VA_USERTYPE     0x0010

typedef struct tagVARDEF {
   struct tagVARDEF NEAR* npNext;
   UINT fuAttribs;
   VARIANT value;
   char szName[1];
} VARDEF;
typedef VARDEF NEAR* NPVARDEF;

typedef struct tagLOCALVARTABINFO {
   NPVARDEF NEAR *npvtLocal;
   struct tagLOCALVARTABINFO NEAR* npNext;
} LOCALVARTABINFO;
typedef LOCALVARTABINFO NEAR* NPLOCALVARTABINFO;


// Maximum number of temporary HLSTR's
#define MAXTEMPHLSTR          32

typedef LPVOID CALLBACK FNNEXTLINE( LPSTR, UINT );
typedef FNNEXTLINE FAR* LPFNNEXTLINE;
typedef void CALLBACK FNGOTOLINE( LPVOID );
typedef FNGOTOLINE FAR* LPFNGOTOLINE;
typedef int CALLBACK FNEXTCMD( LPSTR );
typedef FNEXTCMD FAR* LPFNEXTCMD;
typedef void CALLBACK FNBROADCASTHANDLER( UINT, WORD, LONG );
typedef FNBROADCASTHANDLER FAR* LPFNBROADCASTHANDLER;

// Date order values stored in fuDateOrder member of OPTIONSTRUCT
#define DO_RESET     0
#define DO_DMY       1
#define DO_MDY       2
#define DO_YMD       3

typedef struct tagLOCALHEAPINFO {
   HANDLE   hpool;
   HGLOBAL  hglobal;
   WORD     wSeg;
   BOOL     fAvailable;
   struct tagLOCALHEAPINFO NEAR* npNext;
   struct tagLOCALHEAPINFO NEAR* npPrev;
   struct tagLOCALHEAPINFO NEAR* npLink;
#ifdef _DEBUG   
   USHORT   cObjects;
#endif   
} LOCALHEAPINFO;
typedef LOCALHEAPINFO NEAR* NPLOCALHEAPINFO;

typedef ERR (PASCAL FAR* LPFUNCTION)( int, LPVAR, LPVAR );
typedef struct tagFUNCTIONINFO {
   char szName[MAXTOKEN];
   char szTemplate[MAXTOKEN];
   LPFUNCTION lpfn;
} FUNCTIONINFO;
typedef FUNCTIONINFO NEAR* NPFUNCTIONINFO;

typedef enum tagOPENMODE {
   OM_INPUT,
   OM_OUTPUT,
   OM_APPEND,
   OM_BINARY
} OPENMODE;

typedef enum tagACCESSMODE {
   AM_READ,
   AM_WRITE,
   AM_READWRITE
} ACCESSMODE;

typedef struct tagOPENFILESTRUCT {
   char szFileName[MAXTEXT];
   OPENMODE    enumOpenMode;
   ACCESSMODE  enumAccessMode;
   int nRecLen;
   HFILE hFile;
   BOOL fOpen;
   BOOL fEof;
   HGLOBAL hglbBuffer;
   LPBYTE lpBuffer;
   UINT uBufIndex;
   UINT uByteCount;
} OPENFILESTRUCT;

// Linked list structure for Library module
typedef struct tagLIBLINE {
   struct tagLIBLINE FAR* lpNext;
   struct tagLIBLINE FAR* lpPrev;
   HMEM hmem;
   UINT uLineNum;
   char szText[1];
} LIBLINE;
typedef LIBLINE FAR* LPLIBLINE;

typedef struct tagSHELLWAITINFO {
   HINSTANCE hinstChild;
   HMODULE hmodChild;
   HWND hwndWait;
   long lExitCode;
} SHELLWAITINFO;

typedef struct tagDLLLISTENTRY {
   HINSTANCE hinstDLL;
   HMODULE hmodDLL;
   char szFileName[13];
   struct tagDLLLISTENTRY NEAR* npNext;
} DLLLISTENTRY;
typedef DLLLISTENTRY NEAR* NPDLLLISTENTRY;


/*--------------------------------------------------------------------------*/
/* WBCONTEXT Structure                                                      */
/*                                                                          */
/* Holds the current execution state of a program module. WUPBAS maintains  */
/* a list of execution contexts which it treates as a LIFO stack. The       */
/* top of the stack (i.e. first in the list) represents the context of the  */
/* currently running program module (SUB or FUNCTION procedure).            */
/* The npContext member variable of the WBTASK structure points to the      */
/* top of the stack. Upon entering a new SUB or FUNCTION procedure a new    */
/* WBCONTEXT structure is created and pushed onto the stack. Upon leaving   */
/* the procedure the context popped off the stack and destroyed.            */
/*--------------------------------------------------------------------------*/
typedef struct tagWBCONTEXT {            

   // State variables for lexical scanner
      char              szProgText[MAXTEXT]; // Holds a copy of the text of
                                       // the macro line being interpreted
      NPSTR             npchNextPos;   // Ptr to start position in szProgText
                                       // for extracting next token
      NPSTR             npchPrevPos;   // Ptr to position in szProgText where
                                       // scanning for current token began
      char              szToken[MAXTEXT];// Textual representation of 
                                       // current token
      TOKEN             tokType;       // Type of current token
      TOKENCLASS        tokClass;      // Class of token
      long              lTokenVal;     // Value of current token if numeric
      int               nLeadingSpaces;// Nr of leading spaces that preceded 
                                       // current token
      BOOL              fEol;          // TRUE when end of line has been 
                                       // parsed
                                       
   // State variables for handling ExitFor and ExitDo statements                                       
      int               nForNextLevel; // For-Next loop nesting level
      int               nDoLoopLevel;  // Do-Loop nestring level
      BOOL              fExitBlock;    // If TRUE rest of block statement to be skipped
      
      struct _find_t    fileinfo;      // Used by DIR function
      
      LPLIBLINE         lpLibCurrent;  // Current source file line descriptor
      UINT              uLineNum;      // Current source file line number
      
      LPFNNEXTLINE      lpfnNextLine;  // Ptr to function that returns text 
                                       // of next line
      LPFNGOTOLINE      lpfnGotoLine;  // Ptr to function that returns text 
                                       // of a specific line
      LPVOID            lpLine;        // Handle of current line
      
      NPVARDEF NEAR*    npvtLocal;     // Ptr to local variable table
      
      NPSTR npszLibName;   // Name of file of currently running procedure
      
      struct tagWBCONTEXT NEAR* npNext;// Ptr to next context
} WBCONTEXT;
typedef WBCONTEXT NEAR* NPWBCONTEXT;

#define DECL_NOTFOUND    0
#define DECL_INTFUNC     1       // Intrinsic function (built-in) or in DLL
#define DECL_USERFUNC    2       // User defined function
#define DECL_USERTYPE    3       // User defined data type
#define DECL_DECLARE     4       // Declared external DLL function

typedef struct tagWBINTFUNC {
   NPFUNCTIONINFO npIntFunc;
} WBINTFUNC;
typedef WBINTFUNC NEAR* NPWBINTFUNC;
   
typedef struct tagWBUSERFUNC {
   LPSTR  lpszName;
   LPVOID lpLine;       // Handle of line containing start of procedure
   int    cArg;         // Number of expected arguments
   NPSTR  npszLibName;   // Name of file containing procedure
} WBUSERFUNC;
typedef WBUSERFUNC NEAR* NPWBUSERFUNC;

typedef struct tagWBLOADFILEENTRY {
   struct tagWBLOADFILEENTRY NEAR* npNext;
   char szLibName[1];
} WBLOADFILEENTRY;
typedef WBLOADFILEENTRY NEAR* NPWBLOADFILEENTRY;

typedef struct tagWBTYPEMEMBER {
   UINT fMemberType;
   UINT uSize;
   UINT uOffset;
   struct tagWBTYPEMEMBER NEAR* npNext;
   struct tagWBUSERTYPE NEAR* npChild;
   char szMemberName[1];
} WBTYPEMEMBER;
typedef WBTYPEMEMBER NEAR* NPWBTYPEMEMBER;

typedef struct tagWBUSERTYPE {
   UINT uSize;
   NPWBTYPEMEMBER npMemberList;
} WBUSERTYPE;
typedef WBUSERTYPE NEAR* NPWBUSERTYPE;

typedef struct tagWBARGATTRIB {
   UINT fArgType;
   NPWBUSERTYPE npUserType;
   BOOL fByVal;
   struct tagWBARGATTRIB NEAR* npNext;
} WBARGATTRIB;
typedef WBARGATTRIB NEAR* NPWBARGATTRIB;

// System DLLs that are always loaded
#define SYSLIB_USER     0x0001
#define SYSLIB_GDI      0x0002
#define SYSLIB_KERNEL   0x0004

typedef struct tagWBDECLARE {
   char szTrueName[MAXTOKEN];
   FARPROC lpfn;
   UINT fuSysLib;
   UINT fReturnType;
   int cArg;
   NPWBARGATTRIB npArgList;
   char szPathName[1];
} WBDECLARE;
typedef WBDECLARE NEAR* NPWBDECLARE;

typedef void NEAR* NPVOID;

typedef union tagNPWBHASHENTRY {
   NPVOID       VoidPtr;
   NPWBINTFUNC  IntFunc;
   NPWBUSERFUNC UserFunc;
   NPWBUSERTYPE UserType;
   NPWBDECLARE  Declare;
} NPWBHASHENTRY;

typedef struct tagWBDECLARATION {
   struct tagWBDECLARATION NEAR* npNext;
   UINT fuDeclType;
   NPWBHASHENTRY npExtra;     // Points to additional data directly
                              // following szName member
   char szName[1];      
} WBDECLARATION;
typedef WBDECLARATION NEAR* NPWBDECLARATION;

// Task structure. Each DLL client calling the DLL's initialization routine
// gets assigned a task structure. This structure records the complete state
// of the DLL for the given task.              
typedef struct tagWBTASK {
      HTASK                htaskClient;   // Task handle of client app
      HWND                 hwndClient;    // Main window handle of client app
      NPWBCONTEXT          npContext;     // Pointer to top of execution 
                                          // context stack
      LPFNEXTCMD           lpfnExtCmd;
      LPFNBROADCASTHANDLER lpfnBroadcastHandler;
      NPVARDEF NEAR        *npvtGlobal;   // Global variable table
      NPWBDECLARATION      DeclarationTable[DECL_TABLE_SIZE];
      NPDLLLISTENTRY       npDLLList;
      NPLOCALHEAPINFO      npLocalHeapList;
      BOOL                 fNoLocalCompact;
      HANDLE               hStrDescPool;
      HANDLE               hStrPool;
      HANDLE               hTypePool;
      HANDLE               hLibPool;
      HLSTR                TempHlstrTable[MAXTEMPHLSTR];
      SHELLWAITINFO        ShellWaitInfo;
      long                 lWakeupTime;
      long                 lOption;
      HLSTR                hlstrInputBox;
      HCURSOR              hcurPrev;
      int                  nWaitCursorCount;
      CATCHBUF             CatchBuf;   
      LPLIBLINE            lpLibHead;
      ERR                  nCustomError;
      LPSTR                lpszCustomErrorText;
      BOOL                 fTrueVar;
      NPVARDEF             npTrueVar;
      BOOL                 fLargeDialogs;

      NPWBLOADFILEENTRY    npLoadFileList;      
      
      BOOL fCompareText;   // If TRUE all text comparisons are case-insensitive
      char szDatimFmt[64]; // Holds current date and time format
      UINT fuDateOrder;    // Indicates current DMY order
      

#ifdef _DEBUG
      UINT                 cObjects;
#endif      
      struct tagWBTASK NEAR* npNext;
} WBTASK;
typedef WBTASK NEAR* NPWBTASK;

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
TOKEN NextToken( void );
TOKEN PeekToken( void );
LPVOID GetCodePointer( void );
void SkipLine( void );
void ResetLine( void );
ERR WBParseDateTimeString( LPSTR lpszDatimString, LPVAR lpResult );
BOOL IsTermToken( TOKEN tok, WORD wTerm );

/* var.c */
UINT WBHashName( LPCSTR lpszName, UINT uTableSize );
NPVARDEF WBInsertLocalVariable( LPCSTR lpszName );
NPVARDEF WBInsertGlobalVariable( LPCSTR lpszName );
NPVARDEF WBLookupVariable( LPSTR pszVarName, BOOL fCreate );
ERR WBGetVariableValue( NPVARDEF npvar, LPVAR lpVariant );
void WBSetVariable( NPVARDEF npvar, LPVAR lpVariant );
ERR WBMakeLongVariant( LPVAR lpvar );
ERR WBVariantToString( LPVAR lpvar );
ERR WBVariantStringToOther( LPVAR lpVar, int nVarType );
ERR WBVariantDateToString( LPVAR lpVariant );
void WBDestroyVarTable( NPVARDEF NEAR* pVarTable);
void WBDestroyVariantValue( LPVAR lpVar );
void WBDumpVarTable( void );

/* statmnt.c */
void WBStatement( void );
ERR WBGlobalStatement( void );

/* wbexpr.c */
ERR WBExpression( LPVAR lpResult );

/* wbdeclar.c */
NPVOID WBLookupDeclaration( LPSTR lpszName, UINT FAR* lpfuDeclType );
void WBAddDeclaration( NPWBDECLARATION npDecl, LPSTR lpszName, UINT fuDeclType );
void WBDestroyDeclarationTable( void );

/* wbfunc.h */
void WBInitFunctionTable( void );
NPFUNCTIONINFO WBLookupFunction( USHORT usHash, LPCSTR lpszFuncName );
ERR WBBuildArgList( LPSTR lpszTemplate, int FAR* lpnarg, LPVAR FAR* lplparg );
void WBFreeArgList( int narg, LPVAR lparg );

void WBStopExecution( ERR err );
void WBRuntimeError( ERR err );

LPVOID WBNextLine( LPSTR lpszBuffer );
void WBGotoLine( LPVOID lpLine );
int WBExtCmd( LPSTR lpszCommand );
void WBBroadcastEvent( UINT uEvent, WORD wParam, LONG lParam );
ERR WBRestoreTaskState( void );
NPWBCONTEXT WBPushContext( void );
void WBPopContext( void );

// wbmisc.c
int WBMessageBox( HWND hwndParent, LPCSTR lpszText, LPCSTR lpszTitle, UINT fuStyle );
void WBShowWaitCursor( void );
void WBRestoreCursor( void );
LPSTR WBReplaceMetaChars(LPSTR lpszText, BOOL fExpand);

// wbmem.c
#ifdef _DEBUG
#define WBLocalAlloc(p,q)  WBLocalAllocDebug(p,q)
#define WBLocalReAlloc(p,q,r)  WBLocalReAllocDebug(p,q,r)
#define WBLocalFree(p)     WBLocalFreeDebug(p)
HLOCAL WBLocalAllocDebug( UINT fuAllocFlags, UINT fuAlloc );
HLOCAL WBLocalReAllocDebug( HLOCAL hloc, UINT fuNewSize, UINT fuFlags );
HLOCAL WBLocalFreeDebug( HLOCAL hloc );
#else
#define WBLocalAlloc(p,q)  LocalAlloc(p,q)
#define WBLocalReAlloc(p,q,r)  LocalReAlloc(p,q,r)
#define WBLocalFree(p)     LocalFree(p)
#endif



// wbstmnt.h
void WBCheckEndOfLine( void );

// wbslist.c
BOOL WBRegisterModel_SListObject( void );

// wbccemul.c
BOOL WBRegisterModel( HANDLE hmodDLL, LPMODEL lpmodel );
ERR WBCreateObject( LPCSTR lpszClassName, LPVAR lpResult );
ERR WBDestroyObject( HCTL hctl );
ERR WBDestroyObjectIfNoRef( HCTL hctl );
void WBObjectRefIncrement( HCTL hctl );
void WBObjectRefDecrement( HCTL hctl );
void WBWireObject( HCTL hctl );
void WBUnWireObject( HCTL hctl );
LPMODEL WBGetControlModel( HCTL hctl );
ERR WBDefControlProc( HCTL hctl, USHORT msg, USHORT wp, LONG lp );
LPMETHODINFO WBLookupMethod( HCTL hctl, LPCSTR lpszMethodName );
USHORT WBLookupProperty( HCTL hctl, LPCSTR lpszProp );
ERR WBSetPropertyVariant( HCTL hctl, USHORT iProp, LONG lIndex, LPVAR lpvar );
ERR WBGetPropertyVariant( HCTL hctl, USHORT iProp, LONG lIndex, LPVAR lpResult );
LPVOID WBDerefControl( HCTL hctl );
ERR WBSetControlProperty( HCTL hctl, USHORT iProp, LONG Data );
ERR WBGetControlProperty( HCTL hctl, USHORT iProp, LPVOID lpdata );

// defined in wnfedit.c
USHORT WBStrPos( USHORT usPos, LPSTR lpsz1, USHORT usLen1, 
                 LPSTR lpsz2, USHORT usLen2, BOOL fCompareText );

// wblib.c
ERR WBLoadProgram( LPSTR lpszFileName );
void WBLoadLibStatement( void );
ERR WBCallUserFunction( NPWBUSERFUNC npfun, int narg, LPVAR lparg, LPVAR lpResult );
void WBDestroyLoadLibList( void );
#ifdef _DEBUG
void WBDestroyLib( LPLIBLINE lpHead );
#endif

int WBDestroyOrphanTempHlstr( void );

void WBLoadDLLStatement( void );
void WBUnloadExplicitLinks( void );

extern NPWBTASK g_npTask;
extern NPWBCONTEXT g_npContext;

__inline long LastInteger( void )
{
   return g_npContext->lTokenVal;
}

__inline LPSTR GetTokenText( void )
{
   return g_npContext->szToken;
}

__inline TOKEN LastToken( void )
{
   return g_npContext->tokType;
}

__inline TOKENCLASS LastTokenClass( void )
{
   return g_npContext->tokClass;
}


