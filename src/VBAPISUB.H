#ifdef __cplusplus 
extern "C" { 
#endif 
//---------------------------------------------------------------------------
//
//  VBAPI.H -
// Copyright (c) 1991-1992, Microsoft Corp.  All Rights Reserved.
//
// Include file for the Visual Basic Custom Control and DLL
// application programming interface.
// 
// Modified for use by WUPBAS.DLL
//---------------------------------------------------------------------------

//---------------------------------------------------------------------------
// General types and modifiers
//---------------------------------------------------------------------------
#define WB_FALSE     0
#define WB_TRUE      -1

#define DISKSPACE_CHANGE  1

#define FP              FAR  *
typedef   signed short  SHORT;      // ad lib
typedef unsigned short  USHORT;     // ad lib
typedef unsigned int    FSHORT;     // fs
typedef   signed char   CHAR;       // ch
typedef unsigned char   ENUM;       // enum
typedef unsigned long   ULONG;      // ul
typedef unsigned long   FLONG;      // fl
typedef SHORT  FP       LPSHORT;
typedef USHORT          IPROP;          // iprop
typedef HANDLE          HFORMFILE;  // hformfile
typedef WORD            HPIC;    // hpic
//typedef LPSTR           HSZ;     // hsz
typedef LPVOID          HLSTR;      // hlstr
typedef LPVOID          HAD;     // hAD - handle to Basic array desc.
typedef LPVOID          HCTL;    // hctl
typedef int             ERR;
typedef const void FAR* HANDLE32;
typedef HANDLE32        HMEM;

typedef ERR (*PCTLPROC)(HCTL hctl, USHORT msg, USHORT wp, LONG lp); // pctlproc

//---------------------------------------------------------------------------
// language Variant structures
//---------------------------------------------------------------------------

// variant type constants. these are the return value of GetVariantType.
// these are also defined in Constant.Txt for return value of VarType()

#define VT_EMPTY     0  // Empty
#define VT_NULL      1  // Null
#define VT_I2        2  // Integer
#define VT_I4        3  // Long
#define VT_R4        4  // Single
#define VT_R8        5  // Double
#define VT_CURRENCY  6  // Currency
#define VT_DATE      7  // Date
#define VT_STRING    8  // String
#define VT_USERTYPE  9  // User defined Type

//---------------------------------------------------------------------------
// Property data structure
//---------------------------------------------------------------------------
typedef struct tagDATASTRUCT
{
    LONG   data;      // Data for Get/Set
    USHORT cindex;       // Number of indecies
    struct
    {
   USHORT datatype;    // Type of nth index (Currently always DT_SHORT)
   LONG   data;       // Value of nth index
    } index[1];       // Currently, only 1-dim arrays supported
} DATASTRUCT;

typedef DATASTRUCT FAR  *LPDATASTRUCT; // pds
typedef DATASTRUCT NEAR *NPDATASTRUCT; // npds


//---------------------------------------------------------------------------
// VISUAL BASIC variable types
//---------------------------------------------------------------------------
#define ET_I2     1   // 16 bit signed integer scalar or array variable
#define ET_I4     2   // 32 bit signed integer scalar or array variable
#define ET_R4     3
#define ET_R8     4   // 64 bit real scalar or array variable
#define ET_CY     5   // 64 bit currency scalar or array variable
#define ET_HLSTR  6   // string scalar or array variable
#define ET_SD     6   // older, unpreferred equivalent to ET_HLSTR
#define ET_FS     7   // fixed-length string variable

//---------------------------------------------------------------------------
// Control Property definitions and structures.
//---------------------------------------------------------------------------

#define PF_datatype        0x000000FFL
#define PF_fPropArray      0x00000100L
#define PF_fSetData        0x00000200L
#define PF_fSetMsg         0x00000400L
#define PF_fNoShow         0x00000800L
#define PF_fNoRuntimeW     0x00001000L
#define PF_fGetData        0x00002000L
#define PF_fGetMsg         0x00004000L
#define PF_fSetCheck       0x00008000L
#define PF_fSaveData       0x00010000L
#define PF_fSaveMsg        0x00020000L
#define PF_fLoadDataOnly   0x20010000L
#define PF_fLoadMsgOnly    0x20020000L
#define PF_fGetHszMsg      0x00040000L
#define PF_fUpdateOnEdit   0x00080000L
#define PF_fEditable       0x00100000L
#define PF_fPreHwnd        0x00200000L
#define PF_fDefVal         0x00400000L
#define PF_fNoInitDef      0x00800000L
#define PF_fNoRuntimeR     0x02000000L
#define PF_fNoMultiSelect  0x04000000L

typedef struct tagPROPINFO
  {
  PSTR   npszName;
  FLONG  fl;               // PF_ flags
  BYTE   offsetData;       // Offset into static structure
  } PROPINFO;

typedef PROPINFO  FAR  * LPPROPINFO;
typedef PROPINFO  NEAR * NPPROPINFO;
typedef PROPINFO  NEAR *  PPROPINFO;
typedef PPROPINFO FAR  * LPPROPLIST;
typedef PPROPINFO NEAR * NPPROPLIST;
typedef PPROPINFO NEAR *  PPROPLIST;

// Values for PROP PF_datatype

#define DT_HSZ          0x01
#define DT_SHORT        0x02
#define DT_LONG         0x03
#define DT_BOOL         0x04
#define DT_HLSTR        0x0D

//---------------------------------------------------------------------------
// Custom method definitions
//---------------------------------------------------------------------------

union VARIANTUNION {
   HLSTR tString;
   long tLong;
   HCTL tCtl;
   LPVOID tlpVoid;
};

typedef struct tagVARIANT {
   int type;
   union VARIANTUNION var;
} VARIANT;

typedef VARIANT FAR* LPVAR;

typedef struct tagPARAMINFO {
   struct tagPARAMINFO NEAR* pNext;
   VARIANT value;
} PARAMINFO;
typedef PARAMINFO NEAR* NPPARAMINFO;


typedef ERR (FAR *LPFNMETHOD)( HCTL, int, LPVAR, LPVAR );

typedef struct tagMETHODINFO {
   NPSTR npszName;
   NPSTR npszTemplate;
   LPFNMETHOD lpfn;
} METHODINFO;

typedef METHODINFO  FAR  * LPMETHODINFO;
typedef METHODINFO  NEAR * NPMETHODINFO;
typedef METHODINFO  NEAR *  PMETHODINFO;
typedef PMETHODINFO FAR  * LPMETHODLIST;
typedef PMETHODINFO NEAR * NPMETHODLIST;
typedef PMETHODINFO NEAR *  PMETHODLIST;

//---------------------------------------------------------------------------
// Control MODEL structure
//---------------------------------------------------------------------------
typedef struct tagMODEL
  {
  PCTLPROC     pctlproc;      // The control proc.
  USHORT       cbCtlExtra;    // # bytes alloc'd for HCTL structure
  NPPROPLIST   npproplist;    // Property list
  NPMETHODLIST npmethodlist;  // Method list
  PSTR         npszClassName; // Visual Basic class name
  int          nObjectType;
  BOOL         fInitialized;
  } MODEL;

typedef MODEL FAR  *  LPMODEL;

//---------------------------------------------------------------------------
// MODELINFO structure
//---------------------------------------------------------------------------
typedef struct tagMODELINFO
  {
  USHORT usVersion;     // VB version used by control
  LPMODEL FAR  *lplpmodel;    // pointer to null-terminated
  } MODELINFO;          //    list of LPMODELS

typedef MODELINFO FAR *LPMODELINFO;

//---------------------------------------------------------------------------
// Visual Basic messages
//---------------------------------------------------------------------------
#define VBM__BASE    ( WM_USER + 0x0C00 )

// Pre-hwnd properties are not yet loaded (no hwnd exists for control yet).
// Must set MODEL_fInitMsg to receive.
#define VBM_INITIALIZE     ( VBM__BASE + 0x02 )

// Get a property value.
// [wp=iprop, lp=pdata], RetVal=ERR
#define VBM_GETPROPERTY    ( VBM__BASE + 0x03 )

// Verify a property value.
// [wp=iprop, lp=data], RetVal=ERR
#define VBM_CHECKPROPERTY  ( VBM__BASE + 0x04 )

// Set a property value.
// [wp=iprop, lp=data], RetVal=ERR
#define VBM_SETPROPERTY    ( VBM__BASE + 0x05 )


HLSTR WINAPI FAR __export WBCreateHlstr( LPVOID pb, USHORT cbLen );
HLSTR WINAPI FAR __export WBCreateHlstrFromTemp( HLSTR hlstr );
HLSTR WINAPI FAR __export WBCreateTempHlstr( LPVOID pb, USHORT cbLen );
HLSTR WINAPI FAR __export WBCreateTempHlstrFromTemp( HLSTR hlstr );
void WINAPI FAR __export WBDestroyHlstr( HLSTR hlstr );
BOOL WINAPI FAR __export WBDestroyHlstrIfTemp( HLSTR hlstr );
USHORT WINAPI FAR __export WBGetHlstrLen( HLSTR hlstr );
USHORT WINAPI FAR __export WBGetHlstr( HLSTR hlstr, LPVOID pb, USHORT cbLen );
ERR WINAPI FAR __export WBSetHlstr( HLSTR FAR* phlstr, LPVOID pb, USHORT cbLen );
ERR WINAPI FAR __export WBResizeHlstr( HLSTR hlstr, USHORT newCbLen );
LPSTR WINAPI FAR __export WBDerefHlstr( HLSTR hlstr );
LPSTR WINAPI FAR __export WBDerefHlstrLen( HLSTR hlstr, USHORT FAR* lpCbLen );
LPSTR WINAPI FAR __export WBDerefZeroTermHlstr( HLSTR hlstr );
LPSTR WINAPI FAR __export WBLockHlstr( HLSTR hlstr );
LPSTR WINAPI FAR __export WBLockHlstrLen( HLSTR hlstr, USHORT FAR* pcbLen );
void WINAPI FAR __export WBUnlockHlstr( HLSTR hlstr );
ERR WINAPI __export WBSetErrorMessage( ERR err, LPSTR lpszErrorText );
HWND WINAPI __export WBGetHostWindow( void );

#ifdef __cplusplus 
} 
#endif 
