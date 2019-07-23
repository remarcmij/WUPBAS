#define STRICT
#include <windows.h>
#include <string.h>
#include "wupbas.h"
#include "wbfnstr.h"
#include "wbfndos.h"
#include "wbfnwin.h"
#include "wbfntime.h"
#include "wbfnfile.h"
#include "wbfnothr.h"
#include "wbdllcal.h"

extern NPWBTASK g_npTask;

static char npszNull[] = "";

FUNCTIONINFO FnTable[] = {
   "ASC",            "S",     WBFnAsc,
   "BEEP",           "I",     WBFnBeep,
   "CHDIR",          "S",     WBFnChDir,
   "CHDRIVE",        "S",     WBFnChDrive,
   "CHECKBOX",       "SSSIII",WBFnCheckBox,
   "CHR",            "I",     WBFnChr,
   "CPUTYPE",        "",      WBFnCPUType,
   "CURDIR",         "",      WBFnCurDir,
   "DAY",            "",      WBFnDay,
   "DIR",            "SI",    WBFnDir,
   "DIREXIST",       "S",     WBFnDirExist,
   "DIRLIST",        "SI",    WBFnFileList,     // Obsoleted
   "ENABLELARGEDIALOGS", "I", WBFnEnableLargeDialogs,
   "ENVIRON",        "S",     WBFnEnviron,
   "ENVIRONLIST",    "",      WBFnEnvironList,
   "EXITWINDOWS",    "I",     WBFnExitWindows,
   "EXITWINDOWSEXEC","SS",    WBFnExitWindowsExec,
   "FILECOPY",       "SS",    WBFnFileCopy,
   "FILEDATE",       "S",     WBFnFileDate,
   "FILEDELETE",     "S",     WBFnFileDelete,
   "FILEEXIST",      "S",     WBFnFileExist,
   "FILEISOPEN",     "S",     WBFnFileIsOpen,
   "FILELEN",        "S",     WBFnFileLen,
   "FILELIST",       "SI",    WBFnFileList,
   "FILERENAME",     "S",     WBFnFileRename,
   "FILESEARCH",     "SS",    WBFnFileSearch,
   "FILESELECT",     "SSSS",  WBFnFileSelect,
   "FINDMODULE",     "S",     WBFnFindModule,
   "FINDWINDOW",     "S",     WBFnFindWindow,
   "GETATTR",        "S",     WBFnGetAttr,
   "GETDOSVERSION",  "",      WBFnGetDosVersion,
   "GETDISKFREESPACE","S",    WBFnGetDiskFreeSpace,
   "GETDISKPARMS",   "S",     WBFnGetDiskParms,
   "GETDRIVETYPE",   "S",     WBFnGetDriveType,
   "GETEXTMEMSIZE",  "",      WBFnGetExtMemSize,
   "GETINSTANCE",    "I",     WBFnGetInstance,
   "GETSYSTEMMETRICS", "",    WBFnGetSystemMetrics,
   "GETWINDIR",      "",      WBFnGetWinDir,
   "GETWINVERSION",  "",      WBFnGetWinVersion,
   "HEX",            "I",     WBFnHex,
   "INICAPS",        "S",     WBFnIniCaps,
   "INIDELETE",      "SSS",   WBFnIniDelete,
   "INIDELETEPVT",   "SSS",   WBFnIniDeletePvt, // Obsoleted
   "INIFLUSH",       "S",     WBFnIniFlush,
   "INIFLUSHPVT",    "S",     WBFnIniFlushPvt,  // Obsoleted
   "INILIST",        "SS",    WBFnIniList,
   "INIREAD",        "SSSS",  WBFnIniRead,
   "INIREADPVT",     "SSSS",  WBFnIniReadPvt,   // Obsoleted
   "INIWRITE",       "SSSS",  WBFnIniWrite,
   "INIWRITEPVT",    "SSSS",  WBFnIniWritePvt,  // Obsoleted
   "INPUTBOX",       "SSSII", WBFnInputBox,
   "INSTR",          "",      WBFnInStr,
   "INTERPRET",      "S",     WBFnInterpret,
   "ISNUMERIC",      "",      WBFnIsNumeric,
   "LCASE",          "S",     WBFnLcase,
   "LEFT",           "SI",    WBFnLeft,
   "LEFTPART",       "SS",    WBFnLeftPart,
   "LEN",            "",      WBFnLen,
   "LOADFILE",       "S",     WBFnLoadFile,
   "LTRIM",          "S",     WBFnLTrim,
   "MID",            "SII",   WBFnMid,
   "MKDIR",          "S",     WBFnMkDir,
   "MONTH",          "",      WBFnMonth,
   "MSGBOX",         "SIS",   WBFnMsgBox,
   "NOW",            "",      WBFnNow,
   "OPTIONBOX",      "SSSIII",WBFnOptionBox,
   "OVERLAY",        "SISII", WBFnOverlay,
   "READROMBIOS",    "II",    WBFnReadROMBios,
   "RESTORECURSOR",  "",      WBFnRestoreCursor,
   "RIGHT",          "SI",    WBFnRight,
   "RIGHTPART",      "SS",    WBFnRightPart,
   "RMDIR",          "S",     WBFnRmDir,
   "RTRIM",          "S",     WBFnRTrim,
   "SETATTR",        "SI",    WBFnSetAttr,
   "SHELL",          "SSI",   WBFnShell,
   "SHELLWAIT",      "I",     WBFnShellWait,
   "SHOWWAITCURSOR", "",      WBFnShowWaitCursor,
   "SPACE",          "I",     WBFnSpace,
   "STR",            "",      WBFnStr,
   "STRING",         "I",     WBFnString,
   "TIMER",          "",      WBFnTimer,
   "TOKEN",          "SIS",   WBFnToken,
   "TOKENCOUNT",     "SS",    WBFnTokenCount,
   "TOKENEXTRACT",   "ISS",   WBFnTokenExtract,
   "TOKENS",         "SS",    WBFnTokenCount,
   "TREEDELETE",     "SI",    WBFnTreeDelete,
   "TRIM",           "S",     WBFnTrim,
   "TRUNCSTRING",    "S",     WBFnTruncString,
   "UCASE",          "S",     WBFnUcase,
   "VARTYPE",        "",      WBFnVarType,
   "WAKEUP",         "",      WBFnWakeup,
   "WEEKDAY",        "",      WBFnWeekDay,
   "WORD",           "SI",    WBFnWord,
   "WORDS",          "S",     WBFnWords,
   "YEAR",           "",      WBFnYear
};

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBInitFunctionTable( void )
{
   int i;
   NPWBDECLARATION npDecl;
   NPWBINTFUNC npIntFunc;
   UINT uSize, cb;
   
   for ( i = 0; i < sizeof(FnTable) / sizeof(FUNCTIONINFO); i++ ) {
   
      cb = (UINT)strlen( FnTable[i].szName );
      uSize = sizeof(WBDECLARATION) + cb;
      
      // Adjust size for word alignment of extra data
      if ( uSize & 0x0001 )
         uSize++;
         
      uSize += sizeof(WBINTFUNC);
      
      npDecl = (NPWBDECLARATION)WBLocalAlloc( LPTR, uSize + sizeof(WBINTFUNC) );
      if ( npDecl == NULL )
         break;

      npIntFunc = (NPWBINTFUNC)(((NPSTR)npDecl) + uSize);
      npDecl->npExtra.IntFunc = npIntFunc;
      npIntFunc->npIntFunc = &(FnTable[i]);
   
      WBAddDeclaration( npDecl, FnTable[i].szName, DECL_INTFUNC );
   }
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBBuildArgList( LPSTR lpszTemplate, int FAR* lpnarg, LPVAR FAR* lplparg )
{
   TOKEN tok;
   VARIANT vResult;
   VARIANT NEAR* nparg;
   int narg;
   ERR err;

   nparg = (VARIANT NEAR*)WBLocalAlloc( LPTR, sizeof(VARIANT) * MAXARGS );
   if ( nparg == NULL )
      return WBERR_OUTOFMEMORY;
      
   if ( lpszTemplate == NULL )
      lpszTemplate = npszNull;
            
   narg = 0;   
   err  = 0;
   tok  = LastToken();
   
   while ( narg < MAXARGS && err == 0 ) {

      if ( tok == TOKEN_COMMA ) {
         vResult.var.tLong = 0;
         vResult.type = VT_EMPTY;
      }
      else {
         if ( ( err = WBExpression( &vResult ) ) != 0 )
            break; 
      }
      
      switch ( *lpszTemplate ) {
      
         case 'I':
            switch ( vResult.type ) {
               case VT_I4:
               case VT_EMPTY:
                  nparg[narg++] = vResult;
                  lpszTemplate++;
                  break;
                  
               case VT_STRING:
                  if ( (err = WBMakeLongVariant( &vResult )) == 0 ) {
                     nparg[narg++] = vResult;
                     lpszTemplate++;   
                  }
                  break;
                  
               default:
                  nparg[narg++] = vResult;
                  err = WBERR_ARGTYPEMISMATCH;
            }
            break; 
         
         case 'S':
            switch ( vResult.type ) {
               case VT_STRING:
               case VT_EMPTY:
                  nparg[narg++] = vResult;
                  lpszTemplate++;
                  break;
                  
               case VT_I4:
               case VT_DATE:
                  if ( (err = WBVariantToString( &vResult )) == 0 ) {
                     nparg[narg++] = vResult;
                     lpszTemplate++;
                  }
                  break;
                  
               default:
                  nparg[narg++] = vResult;
                  err = WBERR_ARGTYPEMISMATCH;
            }
            break;
         
         case '\0':
            nparg[narg++] = vResult;
            break;
         
         default:
            assert( FALSE );
            err = WBERR_INTERNAL_ERROR;
      }

      if ( err != 0 )
         break;

      tok = LastToken();
      if ( tok == TOKEN_COMMA ) {
         tok = NextToken();
         if ( tok == TOKEN_EOL )
            tok = NextToken();
      }
      else
         break;
   }
   
   if ( narg == MAXARGS && err == 0 ) 
      err = WBERR_TOOMANYARGS;
   
   if ( err == 0 ) {
      *lpnarg = narg;
      *lplparg = (LPVAR)nparg;
   }
   else {
      WBFreeArgList( narg, nparg );
      *lpnarg = 0;
      *lplparg = NULL;
   }
   
   return err;
}
         
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
void WBFreeArgList( int narg, LPVAR lparg )
{
   int i;

   if ( lparg != NULL ) {
         
      for ( i = 0; i < narg; i++ ) {
         if ( lparg[i].type == VT_STRING )
            WBDestroyHlstrIfTemp( lparg[i].var.tString );
         else if ( lparg[i].type > VT_OBJECT ) 
            WBDestroyObjectIfNoRef( lparg[i].var.tCtl );
      }
      
      WBLocalFree( (HLOCAL)LOWORD( lparg ) );
   }
}         
   
