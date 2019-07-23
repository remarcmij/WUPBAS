#include <windows.h>

typedef struct tagWBSYMBOL {
   struct tagWBSYMBOL NEAR* npNext;
   BYTE fSymbolType;
   BYTE cbSymbolName;
   char szSymbolName[1];
} WBSYMBOL;
typedef WBSYMBOL NEAR* NPWBSYMBOL;

typedef struct tagWBVARIABLE {
   VARIANT value;
} WBVARIABLE;
typedef WBVARIABLE NEAR* NPWBVARIABLE;

typedef struct tagWBUSERFUNC {
      LPVOID lpLine;       // Handle of line containing start of procedure
      int    cArg;         // Number of expected arguments
} WBUSERFUNC;
typedef WBUSERFUNC NEAR* NPWBUSERFUNC;

typedef struct tagWBUSERTYPE {
   WORD wSize;
   NPWBTYPEMEMBER npMemberList;
} WBUSERTYPE;
typedef WBUSERTYPE NEAR* NPWBUSERTYPE;

typedef struct tagWBDECLARE {
   char szTrueName[MAXTOKEN];
   FARPROC lpfn;
   UINT fReturnType;
   int cArg;
   NPARGATTRIB npArgList;
   char szPathName[1];
} WBDECLARE;
typedef WBDECLARE NEAR* NPWBDECLARE;

typedef void NEAR* NPVOID;
            
typedef union tagNPWBHASHENTRY {
   NPVOID       npVoid;
   NPWBVARIABLE npVar;
   NPWBUSERFUNC npUserFunc;
   NPWBUSERTYPE npUserType;
   NPWBDECLARE  npDeclare;
} NPWBHASHENTRY;

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
NPVOID WBLookupSymbol( LPSTR lpszName, LPWORD lpwSymbolType )
{
   NPWBSYMBOL npSymbol;
   
   for ( npSymbol = g_npTask->SymbolTable[ WBHashName( lpszName ) ];
         npSymbol != NULL; 
         npSymbol = npSymbol->npNext ) {
               
      if ( strcmp( npSymbol->szSymbolName, lpszName ) == 0 ) {
         npByte = (BYTE NEAR*)npSymbol;
         *lpwSymbolType = (WORD)npSymbol->fSymbolType;
         return (NPVOID)(npSymbol->szSymbolName + npSymbol->cbSymbolName);
   }
   
   return NULL;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
UINT WBHashName( LPCSTR lpszName )
{
   register UINT uHash = 0;
   
   while ( *lpszName )
      uHash = (uHash<<5) + uHash + *lpszName++;
      
   return ( uHash % HASHTABLESIZE );
}
