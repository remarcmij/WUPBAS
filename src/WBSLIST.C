/* wbslist.c */
#define STRICT
#include <windows.h>
#include <windowsx.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <io.h>
#include <limits.h>
#include <sys\types.h>
#include <sys\stat.h>
#include "wupbas.h"
#include "wbmem.h"
#include "wbfunc.h"
#include "wbslist.h"
#include "wbfileio.h"

#define LIST_HEAP_CHUNK       1024L     // space for 500 lines per chunk
#define MAX_SEGMENT_SIZE      65000L
#define IOBUFFERSIZE          2048

extern HINSTANCE g_hinstDLL;
extern char __near g_szAppTitle[];
extern HBRUSH g_hbrPrompt;

static LPSTR lpszNull = "";

typedef struct tagSLIST 
{
// Property data
   HLSTR hlstrFileName;
   HLSTR hlstrBakExt;
   BOOL  fBackup;
   BOOL  fModified;
   LONG  lListIndex;
   LONG  lListCount;
   HLSTR hlstrText;
   HLSTR FAR* ListAr;      // List property array
// Private data
   HGLOBAL hglbListAr;
} SLIST;
typedef SLIST FAR* LPSLIST;
   
typedef struct tagLISTBOXSTRUCT 
{
   LPSTR lpszPrompt;
   LPSTR lpszTitle;
   int   nDefault;
   UINT xPos, yPos;
   BOOL fSorted;
   LPSLIST lpslist;
} LISTBOXSTRUCT;
typedef LISTBOXSTRUCT FAR* LPLISTBOXSTRUCT;

enum property_indices 
{
   IPROP_SLIST_FILENAME,
   IPROP_SLIST_BAKEXT,
   IPROP_SLIST_BACKUP,
   IPROP_SLIST_MODIFIED,
   IPROP_SLIST_LISTINDEX,
   IPROP_SLIST_INDEX,
   IPROP_SLIST_LISTCOUNT,
   IPROP_SLIST_COUNT,
   IPROP_SLIST_TEXT,
   IPROP_SLIST_LIST
};

#define OFFSETIN(struc, field) ((USHORT)LOWORD(&(((struc *)0)->field)))

PROPINFO Property_FileName = 
{
   "FileName",
   DT_HLSTR | PF_fGetData | PF_fSetData,
   OFFSETIN(SLIST, hlstrFileName)
};

PROPINFO Property_BakExt = 
{
   "BakExt",
   DT_HLSTR | PF_fGetData | PF_fSetData,
   OFFSETIN(SLIST, hlstrBakExt)
};

PROPINFO Property_Backup = 
{
   "Backup",
   DT_BOOL | PF_fGetData | PF_fSetData,
   OFFSETIN(SLIST, fBackup)
};

PROPINFO Property_Modified = 
{
   "Modified",
   DT_BOOL | PF_fGetData | PF_fNoRuntimeW,
   OFFSETIN(SLIST, fModified)
};

PROPINFO Property_ListIndex = 
{
   "ListIndex",
   DT_LONG | PF_fGetData | PF_fNoRuntimeW,
   OFFSETIN(SLIST, lListIndex)
};

// Index is an alias for ListIndex
PROPINFO Property_Index = 
{
   "Index",
   DT_LONG | PF_fGetData | PF_fNoRuntimeW,
   OFFSETIN(SLIST, lListIndex)
};

PROPINFO Property_ListCount = 
{
   "ListCount",
   DT_LONG | PF_fGetData | PF_fNoRuntimeW,
   OFFSETIN(SLIST, lListCount)
};

// Count is an alias for ListCount
PROPINFO Property_Count = 
{
   "Count",
   DT_LONG | PF_fGetData | PF_fNoRuntimeW,
   OFFSETIN(SLIST, lListCount)
};

PROPINFO Property_Text = 
{
   "Text",
   DT_HLSTR | PF_fGetData | PF_fNoRuntimeW,
   OFFSETIN(SLIST, hlstrText)
};

PROPINFO Property_List = 
{
   "List",
   DT_HLSTR | PF_fGetMsg | PF_fSetMsg | PF_fPropArray,
   OFFSETIN(SLIST, hglbListAr)
};

PPROPINFO SList_Properties[] = 
{
   &Property_FileName,
   &Property_BakExt,
   &Property_Backup,
   &Property_Modified,
   &Property_ListIndex,
   &Property_Index,
   &Property_ListCount,
   &Property_Count,
   &Property_Text,
   &Property_List,
   NULL
};

ERR WBSList_MethodAddItem(HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult);
ERR WBSList_MethodRemoveItem(HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult);
ERR WBSList_MethodFind(HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult);
ERR WBSList_MethodFindEx(HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult);
ERR WBSList_MethodLocate(HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult);
ERR WBSList_MethodReplace(HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult);
ERR WBSList_MethodSave(HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult);
ERR WBSList_MethodSaveAs(HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult);
ERR WBSList_MethodClear(HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult);
ERR WBSList_MethodListBox(HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult);

METHODINFO Method_AddItem    = { "ADDITEM",    "",     WBSList_MethodAddItem };
METHODINFO Method_RemoveItem = { "REMOVEITEM", "I",    WBSList_MethodRemoveItem };
METHODINFO Method_Find       = { "FIND",       "SII",  WBSList_MethodFind };
METHODINFO Method_FindEx     = { "FINDEX",     "SII",  WBSList_MethodFindEx };
METHODINFO Method_Locate     = { "LOCATE",     "SIII", WBSList_MethodLocate };
METHODINFO Method_Replace    = { "REPLACE",    "SSIII",WBSList_MethodReplace };
METHODINFO Method_Save       = { "SAVE",       "S",    WBSList_MethodSave };
METHODINFO Method_SaveAs     = { "SAVEAS",     "S",    WBSList_MethodSaveAs };
METHODINFO Method_Clear      = { "CLEAR",      "",     WBSList_MethodClear };
METHODINFO Method_ListBox    = { "LISTBOX",    "SSIIII",WBSList_MethodListBox };

// Keep this list ordered from most frequently to least frequently
// used method for best performance
PMETHODINFO SList_Methods[] = 
{
   &Method_AddItem,
   &Method_RemoveItem,
   &Method_Find,
   &Method_FindEx,
   &Method_Locate,
   &Method_Replace,
   &Method_Save,
   &Method_SaveAs,
   &Method_Clear,
   &Method_ListBox,
   NULL
};

ERR WBSListCtlProc (HCTL hctl, USHORT msg, USHORT wp, LONG lp);

MODEL modelSLIST = 
{
   (PCTLPROC)WBSListCtlProc,        // Control procedure
   sizeof(SLIST),                 // Size of SLIST structure
   SList_Properties,                // Property information table
   SList_Methods,                   // Method information table
   "SLIST",                       // Object class name
   VT_SLIST,                      // Object variant type
   FALSE,
};

#define lpSListDEREF(hctl)  ((LPSLIST)WBDerefControl(hctl))


ERR WBSList_OnCreate(HCTL hctl);
ERR WBSList_OnDestroy(HCTL hctl);
ERR WBSList_OnGetListProperty(HCTL hctl, LPDATASTRUCT lpDs);
ERR WBSList_OnSetListProperty(HCTL hctl, LPDATASTRUCT lpDs);
ERR WBSList_AddItemHelper(HCTL hctl, LPSLIST lpslist, HLSTR hlstr, long lIndex);

BOOL CALLBACK __export 
WBListBoxDlgProc(HWND hwndDlg, UINT msg, WPARAM wParam, LPARAM lParam);

/*--------------------------------------------------------------------------*/
/* Note: called by LibMain                                                  */
/*--------------------------------------------------------------------------*/
BOOL WBRegisterModel_SListObject(void)
{
   return WBRegisterModel(NULL, &modelSLIST);
}

//---------------------------------------------------------------------------
// SList Control Procedure
//---------------------------------------------------------------------------
ERR WBSListCtlProc(HCTL hctl, USHORT msg, USHORT wp, LONG lp)
{
   switch (msg) 
   {
      case WM_CREATE:
         WBSList_OnCreate(hctl);
         break;
      
      case VBM_GETPROPERTY:
         if (wp == IPROP_SLIST_LIST) 
            return WBSList_OnGetListProperty(hctl, (LPDATASTRUCT)lp);
         break;
         
      case VBM_SETPROPERTY:
         if (wp == IPROP_SLIST_LIST) 
            return WBSList_OnSetListProperty(hctl, (LPDATASTRUCT)lp);
         break;
         
      case WM_DESTROY:
         WBSList_OnDestroy(hctl);
         break;
   }
               
   return WBDefControlProc(hctl, msg, wp, lp);
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBSList_OnCreate(HCTL hctl)
{
   LPSLIST lpslist;
             
   lpslist = lpSListDEREF(hctl);

   // set private variables
   lpslist->hglbListAr = NULL;
   
   // set property defaults
   lpslist->hlstrFileName = WBCreateHlstr(NULL, 0);
   lpslist->hlstrBakExt   = WBCreateHlstr("BAK", 3);
   lpslist->fBackup       = TRUE;
   lpslist->fModified     = 0;
   lpslist->lListIndex    = -1;
   lpslist->lListCount    = 0L;
   lpslist->hlstrText     = WBCreateHlstr(NULL, 0);
   lpslist->ListAr        = NULL;
   
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBSList_OnDestroy(HCTL hctl)
{
   LPSLIST lpslist;
   USHORT i;
   
   lpslist = lpSListDEREF(hctl);
   
   for (i = 0; i < (USHORT)lpslist->lListCount; i++) 
   {
      assert(lpslist->ListAr[i] != NULL);
      WBDestroyHlstr(lpslist->ListAr[i]);
   }

   if (lpslist->hglbListAr != NULL) 
   {
      GlobalUnlock(lpslist->hglbListAr);
      GlobalFree(lpslist->hglbListAr);
   }
   
   WBDestroyHlstr(lpslist->hlstrFileName);
   WBDestroyHlstr(lpslist->hlstrBakExt);
   WBDestroyHlstr(lpslist->hlstrText);
   
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBSList_OnGetListProperty(HCTL hctl, LPDATASTRUCT lpDs)
{
   LPSLIST lpslist;
   LONG i;
   HLSTR hlstr;
   ERR err;
         
   lpslist = lpSListDEREF(hctl);
   i = lpDs->index[0].data;
   if (i < 0L || i >= lpslist->lListCount)
      return WBERR_INDEXRANGE;
   
   // return the string
   hlstr = WBCreateTempHlstr(NULL, 0);
   if (hlstr == NULL)
      return WBERR_STRINGSPACE;

   if (lpslist->ListAr[i] != NULL) 
   {
      err = WBSetHlstr(&hlstr, (HLSTR)(lpslist->ListAr[i]), (USHORT)-1);
      if (err != 0)
         return err; 
   }
   
   lpDs->data = (LONG)hlstr;
   return 0; 
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBSList_OnSetListProperty(HCTL hctl, LPDATASTRUCT lpDs)
{
   LPSLIST lpslist;
   LONG i;
   HLSTR hlstr;
   ERR err;
         
   lpslist = lpSListDEREF(hctl);
   i = lpDs->index[0].data;
   if (i < 0L || i >= lpslist->lListCount)
      return WBERR_INDEXRANGE;
   
   // replace string by new value
   hlstr = lpslist->ListAr[i];
   if (hlstr == NULL) 
   {
      hlstr = WBCreateHlstr(NULL, 0);
      if (hlstr == NULL)
         return WBERR_STRINGSPACE;
   }
   
   if ((err = WBSetHlstr(&hlstr, (LPVOID)(HLSTR)(lpDs->data), (USHORT)-1)) != 0)
      return err;
   lpslist->ListAr[i] = hlstr;
   
   // set the Modified property
   lpslist->fModified = 1;
   
   return 0;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBSList_MethodAddItem(HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult)
{
   LPSLIST lpslist;
   LONG lIndex;
   VARIANT vTemp;
   ERR err;
   
   lpslist = lpSListDEREF(hctl);
   
   if (narg < 1)
      return WBERR_ARGTOOFEW;
   else if (narg > 2)
      return WBERR_ARGTOOMANY;

   if (lparg[0].type == VT_EMPTY)
      return WBERR_ARGRANGE;
      
   vTemp = lparg[0];
   
   if (narg == 2) 
   {
      if ((err = WBMakeLongVariant(&lparg[1])) != 0)
         return err;
      lIndex = lparg[1].var.tLong;
   }
   else 
      lIndex = lpslist->lListCount;
   
   // validate index
   if (lIndex < 0 || lIndex > lpslist->lListCount)
      return WBERR_INDEXRANGE;
      
   err = 0;
   
   if (vTemp.type < VT_OBJECT) 
   {
      // Simple variant type
      if ((err = WBVariantToString(&vTemp)) == 0)
         err = WBSList_AddItemHelper(hctl, lpslist, vTemp.var.tString, lIndex);         
   }
   else if (vTemp.type == VT_SLIST) 
   {
      // SList object
      long lIndex2;
      LPSLIST lpslist2 = lpSListDEREF(vTemp.var.tCtl);
      
      // Cannot add an SList to itself
      if (hctl == vTemp.var.tCtl)
         return WBERR_SLIST_SELFASSIGN;
      
      // Insert each string from the source SList object into the
      // destination object
      for (lIndex2 = 0; lIndex2 < lpslist2->lListCount; lIndex2++) 
      {
         err = WBSList_AddItemHelper(hctl, 
                                      lpslist, 
                                      lpslist2->ListAr[lIndex2],
                                      lIndex++);
         if (err != 0)
            break;
      }                                     
      
   }
   else 
      err = WBERR_TYPEMISMATCH;
   
   if (err == 0) 
   {
      // set the Modified property
      lpslist->fModified = TRUE;
      // reset ListIndex and Text properties
      lpslist->lListIndex = -1L;
   }
   
   lpResult->var.tLong = WB_TRUE;
   lpResult->type = VT_I4;
   return err;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBSList_AddItemHelper(HCTL hctl, LPSLIST lpslist, HLSTR hlstr, long lIndex)
{
   HLSTR hlstrNew;
   ERR err;

   hlstrNew = WBCreateHlstrFromTemp(hlstr);
   if (hlstrNew == NULL)
      return WBERR_STRINGSPACE;
      
   // add the new string to end of the list
   if ((err = WBSListAppendLine(hctl, (LPVOID)hlstrNew, (USHORT)-1)) != 0)
      return err;
            
   // put element in correct position 
   // note: lListCount was incremented by 1 by WBSListAppendLine()
   if (lIndex != lpslist->lListCount - 1L) 
   {
      hlstrNew = lpslist->ListAr[lpslist->lListCount - 1L];
      // move elements after index value one position down
      _fmemmove(&(lpslist->ListAr[lIndex + 1L]),
                 &(lpslist->ListAr[lIndex]),
                 (size_t)(lpslist->lListCount-lIndex-1L) * sizeof(HLSTR));
      // assign the string to the indexed position                 
      lpslist->ListAr[lIndex] = hlstrNew;
   }
   
   return 0;
}
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBSList_MethodRemoveItem(HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult)
{
   LPSLIST lpslist;
   LONG lIndex;
   
   lpslist = lpSListDEREF(hctl);

   if (narg < 1)
      return WBERR_ARGTOOFEW;
   else if (narg > 1)
      return WBERR_ARGTOOMANY;

   lIndex = lparg[0].var.tLong;

   // validate index
   if (lIndex < 0L || lIndex >= lpslist->lListCount)
      return WBERR_INDEXRANGE;

   // destroy the target string      
   WBDestroyHlstr(lpslist->ListAr[lIndex]);
   lpslist->lListCount--;
   
   // move elements as required to close the gap
   if (lIndex < lpslist->lListCount) 
      _fmemmove(&(lpslist->ListAr[lIndex]),
                 &(lpslist->ListAr[lIndex+1L]),
                 (size_t)(lpslist->lListCount-lIndex) * sizeof(HLSTR));
   
   // set the Modified property
   lpslist->fModified = TRUE;
   
   // reset ListIndex
   lpslist->lListIndex = -1L;
   
   lpResult->var.tLong = WB_TRUE;
   lpResult->type = VT_I4;
   return 0;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBSList_MethodClear(HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult)
{
   LPSLIST lpslist;
   USHORT i;
   
   UNREFERENCED_PARAM(lparg);
   
   if (narg != 0)
      return WBERR_EXPECT_NOARG;
   
   lpslist = lpSListDEREF(hctl);
   
   for (i = 0; i < (USHORT)lpslist->lListCount; i++) 
   {
      assert(lpslist->ListAr[i] != NULL);
      WBDestroyHlstr(lpslist->ListAr[i]);
   }

   if (lpslist->hglbListAr != NULL) 
   {
      GlobalUnlock(lpslist->hglbListAr);
      GlobalFree(lpslist->hglbListAr);
   }
   
   lpslist->ListAr      = NULL;
   lpslist->hglbListAr  = NULL;
   lpslist->lListCount  = 0;
   lpslist->lListIndex  = -1L;
   lpslist->fModified   = FALSE;

   lpResult->var.tLong = WB_TRUE;
   lpResult->type = VT_I4;
   return 0;
}   
      
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBSList_MethodSave(HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult)
{
   LPSLIST lpslist;
   LPSTR lpch;
   USHORT usSize;
   VARIANT vTemp;
   ERR err;
   
   UNREFERENCED_PARAM(lparg);
   
   if (narg != 0)
      return WBERR_EXPECT_NOARG;
      
   lpslist = lpSListDEREF(hctl);
   assert(lpslist->hlstrFileName != NULL);
   
   lpch = WBLockHlstrLen(lpslist->hlstrFileName, &usSize);
   if (usSize == 0) 
   {
      WBUnlockHlstr(lpslist->hlstrFileName);
      return WBERR_FILENAME_NOT_SET;
   }
   
   vTemp.var.tString = WBCreateTempHlstr(lpch, usSize);
   WBUnlockHlstr(lpslist->hlstrFileName);
   
   if (vTemp.var.tString == NULL)
      return WBERR_STRINGSPACE;
            
   vTemp.type = VT_STRING;

   if ((err = WBSList_MethodSaveAs(hctl, 1, &vTemp, lpResult)) != 0)
      return err;
   
   WBDestroyHlstrIfTemp(vTemp.var.tString);
   return 0;
}   
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBSList_MethodSaveAs(HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult)
{
   LPSLIST lpslist;
   char szFileName[MAXTEXT];
   char szBackup1[MAXTEXT];
   OFSTRUCT of;
   OPENFILESTRUCT FAR* lpof;
   USHORT usSize;
   LPSTR lpch1, lpch2;
   USHORT i;
   ERR err;

   if (narg < 1)
      return WBERR_ARGTOOFEW;
   else if (narg > 1)
      return WBERR_ARGTOOMANY;
      
   lpslist = lpSListDEREF(hctl);

   assert(lparg[0].type == VT_STRING);
   lpch1 = WBDerefHlstrLen(lparg[0].var.tString, &usSize);
   usSize = min(usSize, sizeof(szFileName) - 1);
   memcpy(szFileName, lpch1, usSize);
   szFileName[usSize] = '\0';
      
   // Rename original file if so requested
   if (lpslist->fBackup) 
   {
      // make sure that the bakext property has been set
      assert(lpslist->hlstrBakExt != NULL);
      if (WBGetHlstrLen(lpslist->hlstrBakExt) == 0)
         return WBERR_BAKEXT_NOT_SET;      

      strcpy(szBackup1, szFileName);         
                                 
      lpch1 = strrchr(szBackup1, '\\');
      if (lpch1 == NULL)
         lpch1 = szBackup1;

      lpch2 = strchr(lpch1, '.');
      if (lpch2 == NULL) 
      {
         strcat(szBackup1, ".");
         lpch2 = strchr(szBackup1, '.');
      }
      lpch2++;

      lpch1 = WBDerefHlstrLen(lpslist->hlstrBakExt, &usSize);
      usSize = min(usSize, 3);
      memcpy(lpch2, lpch1, usSize);
      lpch2 += usSize - 1;
      while (usSize != 0 && isspace(*lpch2)) 
      {
         lpch2--;
         usSize--;
      }
      *(lpch2 + 1) = '\0';

      // if file to be saved exist already rename it 
      if (OpenFile(szFileName, &of, OF_EXIST) != HFILE_ERROR) 
      {
         // First try with default extension.
         size_t npos;
         int n;
         
         if (rename(szFileName, szBackup1) != 0) 
         {
            char szBackup2[MAXTEXT];
            char szBackup3[MAXTEXT];
            // If unsuccessful, use a series of up to 10 backup files
            // as follows:
            // The last character of the backup file extention is replaced
            // by a digit. File '9' is deleted. Then file '8' is renamed
            // to file '9', file '7' to '8' etc. Finally the orginal
            // backup file name is renamed to file '0', which makes this
            // original backup file name available for the targetted backup.
            npos = strlen(szBackup1) - 1; 
            strcpy(szBackup2, szBackup1);
            strcpy(szBackup3, szBackup1);
            szBackup3[npos] = '9';
            _chmod(szBackup3, _S_IREAD | _S_IWRITE);
            remove(szBackup3);
            
            for (n = 8; n >= 0; n--) 
            {
               szBackup2[npos] = (char)('0' + n);
               _chmod(szBackup2, _S_IREAD | _S_IWRITE);
               rename(szBackup2, szBackup3);
               szBackup3[npos] = (char)('0' + n);
            }
            
            _chmod(szBackup1, _S_IREAD | _S_IWRITE);
            if ((rename(szBackup1, szBackup2)) != 0)
               return WBERR_RENAME_FILE;
            if (rename(szFileName, szBackup1) != 0) 
               return WBERR_RENAME_FILE;
         }
      }
   }

   lpof = (OPENFILESTRUCT FAR*)WBLocalAlloc(LPTR, sizeof(OPENFILESTRUCT));
   if (lpof == NULL)
      return WBERR_OUTOFMEMORY;

   strcpy(lpof->szFileName, szFileName);
   lpof->enumOpenMode = OM_OUTPUT;
   lpof->nRecLen = IOBUFFERSIZE;
   
   if ((err = WBOpenFile(lpof)) != 0) 
      return err;

   WBShowWaitCursor();
  
   for (i = 0, err = 0; i < (USHORT)lpslist->lListCount && err == 0; i++) 
      err = WBWriteString(lpof, lpslist->ListAr[i]);
   
   WBCloseFile(lpof);
   WBLocalFree((HLOCAL)LOWORD(lpof));
   WBRestoreCursor();

   if (err == 0)
      lpslist->fModified = FALSE;
   else 
   {
      // if an error occcurred, restore backup file (if backup made)
      if (lpslist->fBackup) 
      {
         remove(lpof->szFileName);
         rename(szBackup1, lpof->szFileName);
      }
   }

   lpResult->var.tLong = WB_TRUE;
   lpResult->type = VT_I4;
   return err;
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBSListAppendLine(HCTL hctl, LPVOID lpv, USHORT cb)
{
   LPSLIST lpslist;
   HLSTR hlstrText;
   DWORD dwSize;
   HGLOBAL hglb;
                
   lpslist = lpSListDEREF(hctl);
                   
   // set up initial memory for HLSTR array, if not set up
   // already
   if (lpslist->hglbListAr == NULL) 
   {
      lpslist->hglbListAr = GlobalAlloc(GMEM_MOVEABLE, LIST_HEAP_CHUNK);
      if (lpslist->hglbListAr == NULL)
         return WBERR_OUTOFMEMORY;
         
      // lock the global memory (kept locked until SLIST object is cleared 
      // or destroyed)   
      lpslist->ListAr = (HLSTR FAR*)GlobalLock(lpslist->hglbListAr);
   }
      
   // see if there is room to add another array element
   dwSize = GlobalSize(lpslist->hglbListAr);
   if ((DWORD)(lpslist->lListCount+1L) * (DWORD)sizeof(HLSTR) > dwSize) 
   {
      dwSize += LIST_HEAP_CHUNK;
      if (dwSize > MAX_SEGMENT_SIZE)
         return WBERR_OUTOFMEMORY;
         
      // grow the list array memory block
      GlobalUnlock(lpslist->hglbListAr);
      hglb = GlobalReAlloc(lpslist->hglbListAr, dwSize, GMEM_MOVEABLE);
      
      // re-lock memory, irrespective of wether realloc was successful
      lpslist->ListAr = (HLSTR FAR*)GlobalLock(lpslist->hglbListAr);
      if (hglb == NULL)
         return WBERR_OUTOFMEMORY;
   }

   // If cb is -1 a HLSTR is passed in pb rather than a text buffer
   if (cb == (USHORT)-1) 
      lpslist->ListAr[ lpslist->lListCount++ ] = (HLSTR)lpv;
   else 
   {
      // copy the line to a HLSTR string   
      hlstrText = WBCreateHlstr(lpv, cb);
      if (hlstrText == NULL) 
         return WBERR_STRINGSPACE;
      lpslist->ListAr[ lpslist->lListCount++ ] = hlstrText;
   }
   
   return 0;
}

/*--------------------------------------------------------------------------*/
/* slist.Find(sNeedle [,nStart [,nEnd ] ])                                */
/*--------------------------------------------------------------------------*/
ERR WBSList_MethodFind(HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult)
{
   LPSLIST lpslist;
   LPSTR lpsz1, lpsz2;
   USHORT usLen1, usLen2;
   LONG lStart, lEnd, lTemp;
   int i, nStart, nEnd;
   BOOL fInRange;
   BOOL fForward;

   if (narg < 1)
      return WBERR_ARGTOOFEW;
   else if (narg > 3)
      return WBERR_ARGTOOMANY;

   lpslist = lpSListDEREF(hctl);

   if (lparg[0].type == VT_EMPTY)
      return WBERR_ARGRANGE;
   
   // Get optional nStart
   if (narg >= 2) 
   {
      if (lparg[1].type == VT_EMPTY)
         lStart = 0;
      else 
         lStart = lparg[1].var.tLong;
   }
   else 
      lStart = 0;
   
   // Get optional nEnd
   if (narg == 3) 
   {
      if (lparg[2].type == VT_EMPTY)
         lEnd = lpslist->lListCount - 1;
      else 
         lEnd = lparg[2].var.tLong;
   }
   else 
      lEnd = lpslist->lListCount - 1;
   
   if (lStart > (LONG)INT_MAX || lStart < (LONG)INT_MIN ||
        lEnd > (LONG)INT_MAX   || lEnd < (LONG)INT_MIN)
      return WBERR_ARGRANGE;

   if (lEnd >= lStart) 
      fForward = TRUE;
   else 
   {
      fForward = FALSE;
      lTemp = lStart;
      lStart = lEnd;
      lEnd = lTemp;
   }
   
   if (lStart < 0 && lEnd < 0)
      fInRange = FALSE;
   else if (lStart >= lpslist->lListCount && lEnd >= lpslist->lListCount)
      fInRange = FALSE;
   else
      fInRange = TRUE;
      
   if (fInRange) 
   {
      if (fForward) 
      {
         nStart = (int)max(lStart, 0L);
         nEnd   = (int)min(lEnd, lpslist->lListCount - 1);
      }
      else 
      {
         nStart = (int)min(lEnd, lpslist->lListCount - 1);
         nEnd   = (int)max(lStart, 0L);
      }
   }
   
   lpsz1 = WBDerefHlstrLen(lparg[0].var.tString, &usLen1);
   
   // Strip leading blanks from Needle
   while (usLen1 > 0 && isspace(*lpsz1)) 
   {
      usLen1--;
      lpsz1++;
   }
    
   // If start and end indices are not in range or if needle consists of 
   // white space only we need not look further
   if (!fInRange || usLen1 == 0) 
   {
      lpResult->var.tLong = WB_FALSE;
      lpResult->type = VT_I4;
      lpslist->lListIndex = -1;
      return WBSetHlstr(&(lpslist->hlstrText), NULL, 0);
   }
   
   for (i = nStart; ; fForward ? i++ : i--) 
   {
      if (fForward) 
      {
         if (i > nEnd)
            break;
      }
      else 
      {
         if (i < nEnd)
            break;
      }

      lpsz2 = WBDerefHlstrLen(lpslist->ListAr[i], &usLen2);
      // Strip off leading blanks from line
      while (usLen2 > 0 && isspace(*lpsz2)) 
      {
         usLen2--;
         lpsz2++;
      }
      
      if (usLen2 >= usLen1 && _memicmp(lpsz1, lpsz2, (size_t)usLen1) == 0) 
      {
         lpResult->var.tLong = WB_TRUE;
         lpResult->type = VT_I4;
         lpslist->lListIndex = (long)i;
         return WBSetHlstr(&(lpslist->hlstrText), (LPVOID)lpslist->ListAr[i], (USHORT)-1);
      }
   }
   
   lpslist->lListIndex = -1;
   lpResult->var.tLong = WB_FALSE;
   lpResult->type = VT_I4;
   return WBSetHlstr(&(lpslist->hlstrText), NULL, 0);
}

/*--------------------------------------------------------------------------*/
/*                                                                          */
/* slist.FindEx(sTarget [,nStart [,nEnd ] ])                              */
/*                                                                          */
/*--------------------------------------------------------------------------*/
ERR WBSList_MethodFindEx(HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult)
{
   LPSLIST lpslist;
   LPSTR lpszTarget;
   LPSTR lpszWork;
   LPSTR lpszText;
   LONG lStart, lEnd, lTemp;
   int i, nStart, nEnd;
   USHORT usSize, usSize1, usSize2; 
   BOOL fInRange;
   BOOL fForward;

   if (narg < 1)
      return WBERR_ARGTOOFEW;
   else if (narg > 3)
      return WBERR_ARGTOOMANY;

   lpslist = lpSListDEREF(hctl);

   if (lparg[0].type == VT_EMPTY)
      return WBERR_ARGRANGE;
   
   // Get optional nStart
   if (narg >= 2) 
   {
      if (lparg[1].type == VT_EMPTY)
         lStart = 0;
      else 
         lStart = lparg[1].var.tLong;
   }
   else 
      lStart = 0;
   
   // Get optional nEnd
   if (narg == 3) 
   {
      if (lparg[2].type == VT_EMPTY)
         lEnd = lpslist->lListCount - 1;
      else 
         lEnd = lparg[2].var.tLong;
   }
   else 
      lEnd = lpslist->lListCount - 1;

   if (lStart > (LONG)INT_MAX || lStart < (LONG)INT_MIN ||
        lEnd > (LONG)INT_MAX   || lEnd < (LONG)INT_MIN)
      return WBERR_ARGRANGE;

   if (lEnd >= lStart) 
      fForward = TRUE;
   else 
   {
      fForward = FALSE;
      lTemp = lStart;
      lStart = lEnd;
      lEnd = lTemp;
   }
   
   if (lStart < 0 && lEnd < 0)
      fInRange = FALSE;
   else if (lStart >= lpslist->lListCount && lEnd >= lpslist->lListCount)
      fInRange = FALSE;
   else
      fInRange = TRUE;
      
   if (fInRange) 
   {
      if (fForward) 
      {
         nStart = (int)max(lStart, 0L);
         nEnd   = (int)min(lEnd, lpslist->lListCount - 1);
      }
      else 
      {
         nStart = (int)min(lEnd, lpslist->lListCount - 1);
         nEnd   = (int)max(lStart, 0L);
      }
   }
   
   lpszTarget = WBDerefHlstrLen(lparg[0].var.tString, &usSize);
   
   // ignore leading spaces in search argument
   while (usSize != 0 &&  isspace(*lpszTarget)) 
   {
      usSize--;
      lpszTarget++;
   }
    
   // If start and end indices are not in range or if target consists of 
   // white space only we need not look further
   if (!fInRange || usSize == 0) 
   {
      lpResult->var.tLong = WB_FALSE;
      lpResult->type = VT_I4;
      lpslist->lListIndex = -1;
      return WBSetHlstr(&(lpslist->hlstrText), NULL, 0);
   }

   for (i = nStart; ; fForward ? i++ : i--) 
   {
      if (fForward) 
      {
         if (i > nEnd)
            break;
      }
      else 
      {
         if (i < nEnd)
            break;
      }

      lpszWork = lpszTarget;
      usSize1  = usSize;
      lpszText = WBDerefHlstrLen(lpslist->ListAr[i], &usSize2);

      // ignore leading spaces in SLIST text lines
      while (usSize2 != 0 && isspace(*lpszText)) 
      {
         usSize2--;
         lpszText++;
      }

      if (usSize2 != 0) 
      {
         for (; ;)  
         {
            while (usSize1 != 0 && isspace(*lpszWork)) 
            {
               usSize1--;            
               lpszWork++;
            }

            if (usSize1 == 0) 
            {
               lpslist->lListIndex = (long)i;
               lpResult->var.tLong = WB_TRUE;
               lpResult->type = VT_I4;
               return WBSetHlstr(&(lpslist->hlstrText), (LPVOID)lpslist->ListAr[i], (USHORT)-1);
            }

            while (usSize2 != 0 && isspace(*lpszText)) 
            {
               usSize2--;
               lpszText++;
            }

            if (usSize2 == 0)
               break;

            if (toupper(*lpszText) != toupper(*lpszWork))
               break;

            lpszWork++;
            usSize1--;
            
            lpszText++;
            usSize2--;
         }
      }
   }

   lpslist->lListIndex = -1;
   lpResult->var.tLong = WB_FALSE;
   lpResult->type = VT_I4;
   return WBSetHlstr(&(lpslist->hlstrText), NULL, 0);
}

/*--------------------------------------------------------------------------*/
/* slist.Locate(sTarget, nStart, nEnd, fCompare)                          */
/*--------------------------------------------------------------------------*/
ERR WBSList_MethodLocate(HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult)
{
   LPSLIST lpslist;
   LPSTR lpsz1, lpsz2;
   USHORT usLen1, usLen2;
   LONG lStart, lEnd, lTemp;
   int i, nStart, nEnd;
   BOOL fInRange;
   BOOL fForward;
   BOOL fCompareText;

   if (narg < 1)
      return WBERR_ARGTOOFEW;
   else if (narg > 4)
      return WBERR_ARGTOOMANY;

   lpslist = lpSListDEREF(hctl);

   if (lparg[0].type == VT_EMPTY)
      return WBERR_ARGRANGE;
   
   // Get optional nStart
   if (narg >= 2) 
   {
      if (lparg[1].type == VT_EMPTY)
         lStart = 0;
      else 
         lStart = lparg[1].var.tLong;
   }
   else 
      lStart = 0;
   
   // Get optional nEnd
   if (narg == 3) 
   {
      if (lparg[2].type == VT_EMPTY)
         lEnd = lpslist->lListCount - 1;
      else 
         lEnd = lparg[2].var.tLong;
   }
   else 
      lEnd = lpslist->lListCount - 1;

   if (lStart > (LONG)INT_MAX || lStart < (LONG)INT_MIN ||
        lEnd > (LONG)INT_MAX   || lEnd < (LONG)INT_MIN)
      return WBERR_ARGRANGE;

   if (lEnd >= lStart) 
      fForward = TRUE;
   else 
   {
      fForward = FALSE;
      lTemp = lStart;
      lStart = lEnd;
      lEnd = lTemp;
   }
   
   if (lStart < 0 && lEnd < 0)
      fInRange = FALSE;
   else if (lStart >= lpslist->lListCount && lEnd >= lpslist->lListCount)
      fInRange = FALSE;
   else
      fInRange = TRUE;
      
   if (fInRange) 
   {
      if (fForward) 
      {
         nStart = (int)max(lStart, 0L);
         nEnd   = (int)min(lEnd, lpslist->lListCount - 1);
      }
      else 
      {
         nStart = (int)min(lEnd, lpslist->lListCount - 1);
         nEnd   = (int)max(lStart, 0L);
      }
   }
   
   // Get optional fCompare   
   if (narg == 4) 
   {
      if (lparg[3].var.tLong == 0)
         fCompareText = FALSE;
      else
         fCompareText = TRUE;
   }
   else 
      fCompareText = g_npTask->fCompareText;
   
   lpsz1 = WBDerefHlstrLen(lparg[0].var.tString, &usLen1);
   
   // If start and end indices are not in range or if needle consists of 
   // white space only we need not look further
   if (!fInRange || usLen1 == 0) 
   {
      lpResult->var.tLong = WB_FALSE;
      lpResult->type = VT_I4;
      lpslist->lListIndex = -1;
      return WBSetHlstr(&(lpslist->hlstrText), NULL, 0);
   }
   
   for (i = nStart; ; fForward ? i++ : i--) 
   {
   
      if (fForward) 
      {
         if (i > nEnd)
            break;
      }
      else 
      {
         if (i < nEnd)
            break;
      }

      lpsz2 = WBDerefHlstrLen(lpslist->ListAr[i], &usLen2);
      if (usLen2 >= usLen1 && 
           WBStrPos(0, lpsz2, usLen2, lpsz1, usLen1, fCompareText) != (USHORT)-1) 
      {
         lpslist->lListIndex = (long)i;
         lpResult->var.tLong = WB_TRUE;
         lpResult->type = VT_I4;
         return WBSetHlstr(&(lpslist->hlstrText), (LPVOID)lpslist->ListAr[i], (USHORT)-1);
      }
   }
   
   lpslist->lListIndex = -1;
   lpResult->var.tLong = WB_FALSE;
   lpResult->type = VT_I4;
   return WBSetHlstr(&(lpslist->hlstrText), NULL, 0);
}

/*--------------------------------------------------------------------------*/
/* slist.Replace(sOld, sNew ,nStart, nEnd, fCompare)                      */
/*--------------------------------------------------------------------------*/
ERR WBSList_MethodReplace(HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult)
{
   LPSLIST lpslist;
   HLSTR hlstr;
   LPSTR lpsz1, lpsz2, lpsz3;
   USHORT usLen1, usLen2, usLen3;
   LONG lStart, lEnd, lTemp;
   int i, nStart, nEnd;
   USHORT usPos;
   USHORT usChanges;
   USHORT usNewSize;
   int nDiff;
   BOOL fInRange;
   BOOL fForward;
   BOOL fCompareText;
   ERR err;

   if (narg < 2)
      return WBERR_ARGTOOFEW;
   else if (narg > 5)
      return WBERR_ARGTOOMANY;

   lpslist = lpSListDEREF(hctl);

   if (lparg[0].type == VT_EMPTY)
      return WBERR_ARGRANGE;
   if (lparg[1].type == VT_EMPTY)
      return WBERR_ARGRANGE;
   
   // Get optional nStart
   if (narg >= 3) 
   {
      if (lparg[2].type == VT_EMPTY)
         lStart = 0;
      else 
         lStart = lparg[2].var.tLong;
   }
   else 
      lStart = 0;
   
   // Get optional nEnd
   if (narg >= 4) 
   {
      if (lparg[3].type == VT_EMPTY)
         lEnd = lpslist->lListCount - 1;
      else 
         lEnd = lparg[3].var.tLong;
   }
   else 
      lEnd = lpslist->lListCount - 1;

   if (lStart > (LONG)INT_MAX || lStart < (LONG)INT_MIN ||
        lEnd > (LONG)INT_MAX   || lEnd < (LONG)INT_MIN)
      return WBERR_ARGRANGE;

   if (lEnd >= lStart) 
      fForward = TRUE;
   else 
   {
      fForward = FALSE;
      lTemp = lStart;
      lStart = lEnd;
      lEnd = lTemp;
   }
   
   if (lStart < 0 && lEnd < 0)
      fInRange = FALSE;
   else if (lStart >= lpslist->lListCount && lEnd >= lpslist->lListCount)
      fInRange = FALSE;
   else
      fInRange = TRUE;
      
   if (fInRange) 
   {
      if (fForward) 
      {
         nStart = (int)max(lStart, 0L);
         nEnd   = (int)min(lEnd, lpslist->lListCount - 1);
      }
      else 
      {
         nStart = (int)min(lEnd, lpslist->lListCount - 1);
         nEnd   = (int)max(lStart, 0L);
      }
   }
   
   // Get optional fCompare   
   if (narg == 5) 
   {
      if (lparg[4].var.tLong == 0)
         fCompareText = FALSE;
      else
         fCompareText = TRUE;
   }
   else 
      fCompareText = g_npTask->fCompareText;
   
   lpsz1 = WBLockHlstrLen(lparg[0].var.tString, &usLen1);
   lpsz2 = WBLockHlstrLen(lparg[1].var.tString, &usLen2);
   
   if (!fInRange || usLen1 == 0) 
   {
      WBUnlockHlstr(lparg[0].var.tString);
      WBUnlockHlstr(lparg[1].var.tString);
      lpResult->var.tLong = 0;
      lpResult->type = VT_I4;
      return 0;
   }
   
   nDiff = (int)usLen2 - (int)usLen1;
   usChanges = 0;
   err = 0;
   
   WBShowWaitCursor();
   WBEnableLocalCompact(FALSE);
   
   for (i = nStart; ; fForward ? i++ : i--) 
   {
   
      if (fForward) 
      {
         if (i > nEnd)
            break;
      }
      else 
      {
         if (i < nEnd)
            break;
      }

      lpsz3 = WBDerefHlstrLen(lpslist->ListAr[i], &usLen3);
      
      if (usLen3 >= usLen1) 
      { 
         usPos = 0;
         
         for (;;) 
         {
            usPos = WBStrPos(usPos, lpsz3, usLen3, lpsz1, usLen1, fCompareText);
            if (usPos == (USHORT)-1)
               break;
               
            // Resize string, shift characters and deref pointer again
            // if old and new string differ in size
            if (nDiff != 0) 
            {
               usNewSize = usLen3 - usLen1 + usLen2;
               
               if (nDiff < 0) 
               {
                  // Shrinking: move data before resize
                  memmove(lpsz3 + usPos + usLen2, 
                           lpsz3 + usPos + usLen1, 
                           (size_t)(usLen3 - usPos - usLen1));
                  usLen3 -= (USHORT)(-nDiff);
               }
               
               hlstr = lpslist->ListAr[i];
               if ((err = WBResizeHlstr(hlstr, usNewSize)) != 0)
                  break;
                  
               lpsz3 = WBDerefHlstr(hlstr);
               lpslist->ListAr[i] = hlstr;
               
               if (nDiff > 0) 
               {
                  // Growing: move data after resize
                  memmove(lpsz3 + usPos + usLen2, 
                           lpsz3 + usPos + usLen1, 
                           (size_t)(usLen3 - usPos - usLen1));
                  usLen3 += (USHORT)nDiff;
               }
            }
            
            // Copy new string into target line
            memcpy(lpsz3 + usPos, lpsz2, (size_t)usLen2);
            
            // Restart search after string just inserted
            usPos += usLen2;

            // Increment change count            
            usChanges++;
         }
      }
   }
   
   WBEnableLocalCompact(TRUE);
   WBRestoreCursor();
   
   WBUnlockHlstr(lparg[0].var.tString);
   WBUnlockHlstr(lparg[1].var.tString);
      
   if (err != 0)
      return err;
   
   lpslist->lListIndex = -1;
   
   if (usChanges > 0)
      lpslist->fModified = TRUE;
   
   lpResult->var.tLong = (long)usChanges;
   lpResult->type = VT_I4;
   return 0;
}

/*--------------------------------------------------------------------------*/
/* slist.ListBox(prompt [, [title] [,[default][,[sortflag][, xpos, ypos]]]])*/
/*--------------------------------------------------------------------------*/
ERR WBSList_MethodListBox(HCTL hctl, int narg, LPVAR lparg, LPVAR lpResult)
{
   LISTBOXSTRUCT ListBoxStruct;
   LPSLIST lpslist;
   int idDialog;
   int ret;
   LPSTR lpszTemp;
   
   lpslist = lpSListDEREF(hctl);
   
   if (narg < 1)
      return WBERR_ARGTOOFEW;
   else if (narg > 6)
      return WBERR_ARGTOOMANY;
   
   memset(&ListBoxStruct, 0 , sizeof(LISTBOXSTRUCT));
   
   if (lparg[0].type == VT_EMPTY)
      ListBoxStruct.lpszPrompt = lpszNull;
   else 
   {
      lpszTemp = WBDerefZeroTermHlstr( lparg[0].var.tString );
      ListBoxStruct.lpszPrompt = WBReplaceMetaChars(lpszTemp, TRUE);
      if (ListBoxStruct.lpszPrompt == NULL)
         return WBERR_OUTOFMEMORY;
   }
   
   if (narg >= 2 && lparg[1].type != VT_EMPTY) 
      ListBoxStruct.lpszTitle = WBDerefZeroTermHlstr(lparg[1].var.tString);
   else 
      ListBoxStruct.lpszTitle = g_szAppTitle;
   
   if (narg >= 3 && lparg[2].type != VT_EMPTY) 
   {
      if (lparg[2].var.tLong < -1 || lparg[2].var.tLong >= lpslist->lListCount) 
         return WBERR_INDEXRANGE;
      ListBoxStruct.nDefault = (int)lparg[2].var.tLong;
   }
   else 
      ListBoxStruct.nDefault = -1;
   
   if (narg >= 4 && lparg[3].type != VT_EMPTY) 
   {
      if (lparg[3].var.tLong == 0) 
      {
         idDialog = IDD_LISTBOX_UNSORT;
         ListBoxStruct.fSorted = FALSE;
      }
      else 
      {
         idDialog = IDD_LISTBOX_SORT;
         ListBoxStruct.fSorted = TRUE;
      }
   }
   else 
   {
      idDialog = IDD_LISTBOX_UNSORT;
      ListBoxStruct.fSorted = FALSE;
   }
   
   if (narg >= 5) 
   {
      if (narg != 6)
         return WBERR_ARGTOOFEW;
               
      if (lparg[4].type == VT_EMPTY)
         return WBERR_ARGRANGE;
      ListBoxStruct.xPos = (UINT)lparg[4].var.tLong;
      ListBoxStruct.yPos = (UINT)lparg[5].var.tLong;
   }
      
   ListBoxStruct.lpslist = lpslist;

   if (g_npTask->fLargeDialogs)
      idDialog++;
         
   ret = DialogBoxParam(g_hinstDLL, 
                         MAKEINTRESOURCE(idDialog),
                         g_npTask->hwndClient,
                         (DLGPROC)WBListBoxDlgProc,
                         (LPARAM)(LPLISTBOXSTRUCT)&ListBoxStruct);
                         
   WBRestoreTaskState();                         
   WBLocalFree((HLOCAL)LOWORD(ListBoxStruct.lpszPrompt));
   
   lpResult->var.tLong = (long)ret;     
   lpResult->type = VT_I4;
   return 0;
}                         
      
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
BOOL CALLBACK __export 
WBListBoxDlgProc(HWND hwndDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
   WBRestoreTaskState();
   
   switch (msg) 
   {
      case WM_INITDIALOG:
         {             
         LPLISTBOXSTRUCT lpListBoxStruct;
         int nScreenWidth, nScreenHeight;
         int xPos, yPos;
         int nLogPixelsX, nLogPixelsY;
         RECT rcParent, rcDialog;
         HDC hdc;
         LPSLIST lpslist;
         HWND hwndList;
         LPSTR lpszItem;
         char szBuffer[128];
         USHORT usLen;
         LRESULT lresult;
         UINT i;
         WORD cxExtent, cxMaxExtent;
         TEXTMETRIC tm;
         HFONT hfontOld;
         
         lpListBoxStruct = (LPLISTBOXSTRUCT)lParam;
         SetWindowText(hwndDlg, lpListBoxStruct->lpszTitle);
            
         SetDlgItemText(hwndDlg, IDC_EDIT, lpListBoxStruct->lpszPrompt);
         
         lpslist = lpListBoxStruct->lpslist;
         hwndList = GetDlgItem(hwndDlg, IDC_LISTBOX);

         hdc = GetDC(hwndList);
         hfontOld = SelectFont(hdc, GetWindowFont(hwndDlg));
         cxMaxExtent = 0;
         for (i = 0; i < (UINT)lpslist->lListCount; i++) 
         {
            lpszItem = WBDerefHlstrLen(lpslist->ListAr[i], &usLen);
            usLen = min(usLen, sizeof(szBuffer) - 10);
            memcpy(szBuffer, lpszItem, (size_t)usLen);
            szBuffer[usLen] = '\0';
            cxExtent = LOWORD(GetTextExtent(hdc, szBuffer, (int)usLen));
            cxMaxExtent = max(cxExtent, cxMaxExtent);
            lresult = ListBox_AddString(hwndList, szBuffer);
            if (lresult == LB_ERR || lresult == LB_ERRSPACE)
               break;
            ListBox_SetItemData(hwndList, lresult, i);
         }
         GetTextMetrics(hdc, &tm);
         cxMaxExtent += (WORD)tm.tmAveCharWidth;
         SelectFont(hdc, hfontOld);
         ReleaseDC(hwndList, hdc);
         ListBox_SetHorizontalExtent(hwndList, cxMaxExtent);               

         i = lpListBoxStruct->nDefault;
         if (i != -1 && lpListBoxStruct->fSorted) 
         {
            lpszItem = WBDerefZeroTermHlstr(lpslist->ListAr[i]);
            i = (int)SendMessage(hwndList, LB_FINDSTRING, (WPARAM)-1, (LPARAM)lpszItem);
            if (i == LB_ERR)
               i = 0;
         }
         
         // Set default selection if list non-empty
         if (lpslist->lListCount > 0) 
            SendMessage(hwndList, LB_SETCURSEL, (WPARAM)i, 0L);
         
         // Disable OK button if list empty or no default select
         if (lpslist->lListCount == 0 || lpListBoxStruct->nDefault == -1) 
            EnableWindow(GetDlgItem(hwndDlg, IDOK), FALSE);
         
         GetWindowRect(hwndDlg, &rcDialog);
         nScreenWidth = GetSystemMetrics(SM_CXSCREEN);
         nScreenHeight = GetSystemMetrics(SM_CYSCREEN);
         
         if (lpListBoxStruct->xPos != 0) 
         {         
            hdc = GetDC(hwndDlg);
            nLogPixelsX = GetDeviceCaps(hdc, LOGPIXELSX);
            nLogPixelsY = GetDeviceCaps(hdc, LOGPIXELSY);
            ReleaseDC(hwndDlg, hdc);
            xPos = (int)((long)lpListBoxStruct->xPos * 
                        (long)nLogPixelsX / 1440L);
            yPos = (int)((long)lpListBoxStruct->yPos * 
                        (long)nLogPixelsY / 1440L);
         }
         else 
         {
            xPos = (nScreenWidth - (rcDialog.right - rcDialog.left)) / 2;
            yPos = (nScreenHeight - (rcDialog.bottom - rcDialog.top)) / 3;
         }
         
         xPos = max(xPos, 0);
         xPos = min(xPos, nScreenWidth - (rcDialog.right - rcDialog.left));
         yPos = max(yPos, 0);
         yPos = min(yPos, nScreenHeight - (rcDialog.bottom - rcDialog.top));
         
         GetClientRect(g_npTask->hwndClient, &rcParent);
         xPos -= rcParent.left;
         yPos -= rcParent.top;
         
         MoveWindow(hwndDlg, 
                     xPos, 
                     yPos, 
                     rcDialog.right - rcDialog.left,
                     rcDialog.bottom - rcDialog.top,
                     FALSE);
         
         return TRUE;
         }
         
      case WM_COMMAND:
         switch (wParam) 
         {
         
            case IDC_LISTBOX:
               if (HIWORD(lParam) == LBN_DBLCLK)
                  goto got_selection;
               else if (HIWORD(lParam) == LBN_SELCHANGE)
                  EnableWindow(GetDlgItem(hwndDlg, IDOK), TRUE);
               break;
         
            case IDOK:
got_selection:
               {
                  int nIndex;
                  HWND hwndList;
                  
                  hwndList = GetDlgItem(hwndDlg, IDC_LISTBOX);
                  nIndex = ListBox_GetCurSel(hwndList);
                  assert(nIndex != LB_ERR);
                  nIndex = (int)ListBox_GetItemData(hwndList, nIndex);
                  EndDialog(hwndDlg, nIndex);
               }
               return TRUE;
               
            case IDCANCEL:
               EndDialog(hwndDlg, -1);
               return TRUE;
         }
         break;
         
      case WM_CTLCOLOR:
         if (g_npTask->fLargeDialogs)
         {
            if ((HWND)LOWORD(lParam) == GetDlgItem(hwndDlg, IDC_EDIT)) {
               SetBkColor((HDC)wParam, RGB(192, 192, 192));
               return (BOOL)g_hbrPrompt;
            }
         }
         break;
   }
   
   return FALSE;
}
   
/*--------------------------------------------------------------------------*/
/*                                                                          */
/*--------------------------------------------------------------------------*/
USHORT WBStrPos(USHORT usPos, LPSTR lpsz1, USHORT usLen1, 
                 LPSTR lpsz2, USHORT usLen2, BOOL fCompareText)
{
   LPSTR lpPos, lpPos2;
   int chUpper, chLower;
   USHORT usRange;
   BOOL fIsAlpha;
   BOOL fFound;
   int nComp;

   if (usLen1 == 0 || usLen2 == 0 || usLen2 + usPos > usLen1)
      return (USHORT)-1;

   chUpper = chLower = *lpsz2;
   fIsAlpha = isalpha(chUpper);
   if (fCompareText) 
   {
      chUpper = toupper(chUpper);
      chLower = tolower(chUpper);
   }
      
   lpPos = NULL;
   fFound = FALSE;
      
   while (usPos <= usLen1 - usLen2)  
   {
      usRange = usLen1 - usLen2 - usPos + 1;
      
      // perform memory search for ch 
      lpPos = memchr(lpsz1 + usPos, chUpper, usRange);
      if (fCompareText && fIsAlpha) 
      {
         lpPos2 = memchr(lpsz1 + usPos, chLower, usRange);

         if (lpPos == NULL && lpPos2 == NULL) 
            break;
         else if (lpPos == NULL) 
            lpPos = lpPos2;
         else if (lpPos2 != NULL) 
            lpPos = (LPSTR)min((DWORD)lpPos, (DWORD)lpPos2);
      }
      else if (lpPos == NULL) 
         break;

      if (fCompareText) 
         nComp = memicmp(lpPos, lpsz2, usLen2);
      else
         nComp = memcmp(lpPos, lpsz2, usLen2);
         
      if (nComp == 0) 
      {
         fFound = TRUE;
         break;
      }
        
      usPos = (USHORT)(lpPos - lpsz1 + 1);
   }
     
   // at this point, if lpPos == NULL then string was not found,
   // otherwise lpPos is the location of string2 in string1
   if (fFound)
      return (USHORT)(lpPos - lpsz1);
   else
      return (USHORT)-1;
}     

/*--------------------------------------------------------------------------*/   
/*                                                                          */
/*--------------------------------------------------------------------------*/   
ERR WBSListAppendItem(HCTL hctl, LPVAR lpvar)
{
   HLSTR hlstr;
   ERR err;
   
   if ((err = WBVariantToString(lpvar)) != 0)
      return err;

   hlstr = WBCreateHlstrFromTemp(lpvar->var.tString);
   if (hlstr == NULL)
      return WBERR_STRINGSPACE;
      
   return WBSListAppendLine(hctl, (LPVOID)hlstr, (USHORT)-1);
}
   
/*--------------------------------------------------------------------------*/   
/*                                                                          */
/*--------------------------------------------------------------------------*/   
void WBSListSetModifiedFlag(HCTL hctl, BOOL fModified)
{
   lpSListDEREF(hctl)->fModified = fModified;
}

   
