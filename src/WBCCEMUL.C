#define STRICT
#include <windows.h>
#include <string.h>
#include "wupbas.h"
#include "wbmem.h"

typedef struct tagOBJECTINSTANCEINFO {
   LPMODEL lpmodel;
   int nRefCount;
   struct tagOBJECTINSTANCEINFO FAR* lpNext;
   struct tagOBJECTINSTANCEINFO FAR* lpPrev;
} OBJECTINSTANCEINFO;
typedef OBJECTINSTANCEINFO FAR* LPOBJECTINSTANCEINFO;
   
typedef struct tagMODELLISTENTRY {
   LPMODEL lpmodel;
   LPOBJECTINSTANCEINFO lpInstanceList;
   struct tagMODELLISTENTRY NEAR* npNext;
} MODELLISTENTRY;
typedef MODELLISTENTRY NEAR* NPMODELLISTENTRY;

static NPMODELLISTENTRY npmodelList;

/*--------------------------------------------------------------------------*/   
/*                                                                          */
/*--------------------------------------------------------------------------*/   
BOOL WBRegisterModel( HANDLE hmodDLL, LPMODEL lpmodel )
{
   NPMODELLISTENTRY npEntry;
   LPOBJECTINSTANCEINFO lpInstance;
   
   UNREFERENCED_PARAM( hmodDLL );

   // Check for duplicate model registration   
   for ( npEntry = npmodelList; npEntry != NULL; npEntry = npEntry->npNext ) {
      if ( _fstricmp( npEntry->lpmodel->npszClassName, lpmodel->npszClassName ) == 0 )
         return FALSE;
   }
   
   npEntry = (NPMODELLISTENTRY)LocalAlloc( LPTR, sizeof(MODELLISTENTRY));
   if ( npEntry == NULL )
      return FALSE;
      
   lpInstance = (LPOBJECTINSTANCEINFO)LocalAlloc( LPTR, sizeof(OBJECTINSTANCEINFO));
   if ( lpInstance == NULL )
      return FALSE;

   // Cross-link object instance info headnode to itself
   lpInstance->lpNext = lpInstance->lpPrev = lpInstance;
   
   // Add model list entry to list
   npEntry->lpmodel = lpmodel;
   npEntry->lpInstanceList = lpInstance;
   npEntry->npNext = npmodelList;
   npmodelList = npEntry;
   
   return TRUE;
}
   
/*--------------------------------------------------------------------------*/   
/*                                                                          */
/*--------------------------------------------------------------------------*/   
ERR WBCreateObject( LPCSTR lpszClassName, LPVAR lpResult )
{
   NPMODELLISTENTRY npEntry;
   LPOBJECTINSTANCEINFO lpInstance;
   LPMODEL lpmodel;
   HCTL hctl;
   
   // Check for duplicate model registration   
   for ( npEntry = npmodelList; npEntry != NULL; npEntry = npEntry->npNext ) {
      if ( _fstricmp( npEntry->lpmodel->npszClassName, lpszClassName ) == 0 )
         break;
   }
   
   if ( npEntry == NULL )
      return WBERR_CREATEOBJECT;
   
   lpmodel = (LPMODEL)WBLocalAlloc( 
                        LPTR, 
                        sizeof(MODEL) + npEntry->lpmodel->cbCtlExtra );
   if ( lpmodel == NULL )
      return WBERR_OUTOFMEMORY;
                           
   // Copy model class data into object model structure
   memcpy( lpmodel, npEntry->lpmodel, sizeof(MODEL) );

   lpInstance = (LPOBJECTINSTANCEINFO)WBLocalAlloc( 
                     LPTR,
                     sizeof(OBJECTINSTANCEINFO) );
   if ( lpInstance == NULL ) {
      WBLocalFree( (HLOCAL)LOWORD(lpmodel) );
      return WBERR_OUTOFMEMORY;
   }
   
   lpInstance->lpmodel = lpmodel;
   
   // Add the new object instance at the begin of the model's instance
   // list   
   lpInstance->lpNext = npmodelList->lpInstanceList->lpNext;
   lpInstance->lpPrev = npmodelList->lpInstanceList;
   npmodelList->lpInstanceList->lpNext->lpPrev = lpInstance;
   npmodelList->lpInstanceList->lpNext = lpInstance;
   
   // Control handle is the address of the object instance structure
   hctl = (HCTL)lpInstance;

   // Initialize the object through its control procedure      
   if ( (lpmodel->pctlproc( hctl, WM_CREATE, 0, 0L )) != 0 )
      return WBERR_CREATEOBJECT;

   // Flag control as initialized   
   lpmodel->fInitialized = TRUE;
   
#ifdef DEBUGOUTPUT
   {
   char szMessage[128];
   wsprintf( szMessage, 
             "t WUBBAS: object %s created, handle %lX\r\n",
             (LPSTR)lpmodel->npszClassName,
             hctl );
   OutputDebugString( szMessage );
   }
#endif   

   lpResult->var.tCtl = hctl;
   lpResult->type = lpmodel->nObjectType;
   return 0;
}
   
/*--------------------------------------------------------------------------*/   
/*                                                                          */
/*--------------------------------------------------------------------------*/   
ERR WBDestroyObject( HCTL hctl )
{
   LPOBJECTINSTANCEINFO lpInstance;
   
   lpInstance = (LPOBJECTINSTANCEINFO)hctl;
   
   return lpInstance->lpmodel->pctlproc( hctl, WM_DESTROY, 0, 0L );
}

/*--------------------------------------------------------------------------*/   
/*                                                                          */
/*--------------------------------------------------------------------------*/   
ERR WBDestroyObjectIfNoRef( HCTL hctl )
{
   LPOBJECTINSTANCEINFO lpInstance;
   
   lpInstance = (LPOBJECTINSTANCEINFO)hctl;
   
   if ( lpInstance->nRefCount <= 0 )
      return WBDestroyObject( hctl );
}

/*--------------------------------------------------------------------------*/   
/*                                                                          */
/*--------------------------------------------------------------------------*/   
void WBObjectRefIncrement( HCTL hctl )
{
   LPOBJECTINSTANCEINFO lpInstance;
   
   lpInstance = (LPOBJECTINSTANCEINFO)hctl;
   lpInstance->nRefCount++;
}

/*--------------------------------------------------------------------------*/   
/*                                                                          */
/*--------------------------------------------------------------------------*/   
void WBObjectRefDecrement( HCTL hctl )
{
   LPOBJECTINSTANCEINFO lpInstance;
   
   lpInstance = (LPOBJECTINSTANCEINFO)hctl;
   lpInstance->nRefCount--;
}

/*--------------------------------------------------------------------------*/   
/*                                                                          */
/*--------------------------------------------------------------------------*/   
void WBWireObject( HCTL hctl )
{
   LPOBJECTINSTANCEINFO lpInstance;
   
   lpInstance = (LPOBJECTINSTANCEINFO)hctl;
   lpInstance->nRefCount++;
}

/*--------------------------------------------------------------------------*/   
/*                                                                          */
/*--------------------------------------------------------------------------*/   
void WBUnWireObject( HCTL hctl )
{
   LPOBJECTINSTANCEINFO lpInstance;
   
   lpInstance = (LPOBJECTINSTANCEINFO)hctl;
   
   if ( --lpInstance->nRefCount <= 0 )
      WBDestroyObject( hctl );
}

/*--------------------------------------------------------------------------*/   
/*                                                                          */
/*--------------------------------------------------------------------------*/   
ERR WBDefControlProc( HCTL hctl, USHORT msg, USHORT wp, LONG lp )
{
   LPOBJECTINSTANCEINFO lpInstance;
   
   UNREFERENCED_PARAM( wp );
   UNREFERENCED_PARAM( lp );
   
   lpInstance = (LPOBJECTINSTANCEINFO)hctl;
   
   switch ( msg ) {
   
      case WM_DESTROY:
#ifdef DEBUGOUTPUT
         {
            char szMessage[80];
            wsprintf( szMessage, 
                  "t WUPBAS: destroying %s object, handle %lX\r\n",
                  (LPSTR)lpInstance->lpmodel->npszClassName,
                  hctl );
            OutputDebugString( szMessage );
}
#endif          
         // Destroy model structure (including programmer defined bit)
         WBLocalFree( (HLOCAL)LOWORD(lpInstance->lpmodel) );

         // Keep instance info structure until no more references
         // to the object exist
         if ( lpInstance->nRefCount <= 0 ) {         
            // Unlink instance info structure from instance list and
            // destroy
            lpInstance->lpPrev->lpNext = lpInstance->lpNext;
            lpInstance->lpNext->lpPrev = lpInstance->lpPrev;
            WBLocalFree( (HLOCAL)LOWORD(lpInstance) );
         }
         else {
            lpInstance->lpmodel = NULL;
            lpInstance->nRefCount--;  
         }
         break;

   }
   
   return 0;
}
         
/*--------------------------------------------------------------------------*/   
/*                                                                          */
/*--------------------------------------------------------------------------*/   
LPMODEL WBGetControlModel( HCTL hctl )                           
{
   LPOBJECTINSTANCEINFO lpInstance;
   
   lpInstance = (LPOBJECTINSTANCEINFO)hctl;
   
   return lpInstance->lpmodel;
}

/*--------------------------------------------------------------------------*/   
/*                                                                          */
/*--------------------------------------------------------------------------*/   
LPVOID WBDerefControl( HCTL hctl )
{
   LPOBJECTINSTANCEINFO lpInstance;
   
   lpInstance = (LPOBJECTINSTANCEINFO)hctl;
   return (LPVOID)(((LPBYTE)(lpInstance->lpmodel)) + sizeof(MODEL));
}

/*--------------------------------------------------------------------------*/   
/*                                                                          */
/*--------------------------------------------------------------------------*/   
LPMETHODINFO WBLookupMethod( HCTL hctl, LPCSTR lpszMethodName )
{
   LPMODEL lpmodel;
   LPMETHODLIST lpMethodList;
   WORD wSeg;
   LPSTR lpszName;
   
   lpmodel = ((LPOBJECTINSTANCEINFO)hctl)->lpmodel;
   wSeg = HIWORD( lpmodel );
   lpMethodList = MAKELP( wSeg, lpmodel->npmethodlist );
   
   while ( *lpMethodList != NULL ) {
      lpszName = MAKELP( wSeg, (*lpMethodList)->npszName );
      if ( _fstrcmp( lpszName, lpszMethodName ) == 0 ) {
         return *lpMethodList;
      }
      lpMethodList++;
   }
   return NULL;
}
   
/*--------------------------------------------------------------------------*/   
/*                                                                          */
/*--------------------------------------------------------------------------*/   
USHORT WBLookupProperty( HCTL hctl, LPCSTR lpszProp )
{
   LPMODEL lpmodel;
   NPPROPINFO npPropInfo;
   USHORT iProp;
   
   lpmodel = (LPMODEL)((LPOBJECTINSTANCEINFO)hctl)->lpmodel;
   
   // Find property name
   for ( iProp = 0; lpmodel->npproplist[iProp] != NULL; iProp++ ) {
      npPropInfo = lpmodel->npproplist[iProp];
      if ( _fstricmp( npPropInfo->npszName, lpszProp ) == 0 ) {
         return iProp;
      }
   }
   
   return (USHORT)-1;
}

/*--------------------------------------------------------------------------*/   
/*                                                                          */
/*--------------------------------------------------------------------------*/   
ERR WBSetPropertyVariant( HCTL hctl, USHORT iProp, LONG lIndex, LPVAR lpvar )
{
   LPMODEL lpmodel;
   NPPROPINFO npPropInfo;
   BOOL fIsPropArray;
   DATASTRUCT data;
   ERR err = 0;
   
   lpmodel = (LPMODEL)((LPOBJECTINSTANCEINFO)hctl)->lpmodel;
   npPropInfo = lpmodel->npproplist[iProp];
   
   // If PF_fPropArray flag is set for the property prepare a DATASTRUCT
   if ( npPropInfo->fl & PF_fPropArray ) {
      if ( lIndex == -1 )
         return WBERR_PROPARRAY_INDEXNEEDED;
         
      fIsPropArray = TRUE;
      data.cindex = 1;
      data.index[0].datatype = DT_SHORT;
      data.index[0].data = lIndex;
   }
   else {
      fIsPropArray = FALSE;
      if ( lIndex != -1L )
         return WBERR_NOTAPROPARRAY;
   }
   
   switch ( npPropInfo->fl & (LONG)PF_datatype ) {
   
      case DT_HLSTR:
         if ( (err = WBVariantToString( lpvar )) != 0 )
            break;
         if ( fIsPropArray ) {
            data.data = (LONG)(lpvar->var.tString);
            err = WBSetControlProperty( hctl, iProp, (LONG)&data );
         }
         else {
            err = WBSetControlProperty( hctl, iProp, (LONG)(lpvar->var.tString) );
         }
         break; 
      
      case DT_BOOL:
      case DT_SHORT:
      case DT_LONG:
         if ( (err = WBMakeLongVariant( lpvar )) != 0 )
            break;
         if ( fIsPropArray ) {
            data.data = lpvar->var.tLong;
            err = WBSetControlProperty( hctl, iProp, (LONG)&data );
         }
         else {
            err = WBSetControlProperty( hctl, iProp, lpvar->var.tLong );
         }
         break;
         
         
      default:
         assert( FALSE );
         return WBERR_INTERNAL_ERROR;
   }
   
   return err;
}

/*--------------------------------------------------------------------------*/   
/*                                                                          */
/*--------------------------------------------------------------------------*/   
ERR WBGetPropertyVariant( HCTL hctl, USHORT iProp, LONG lIndex, LPVAR lpResult )
{
   LPMODEL lpmodel;
   NPPROPINFO npPropInfo;
   DATASTRUCT data;
   LONG lValue;
   ERR err;
   
   lpmodel = (LPMODEL)((LPOBJECTINSTANCEINFO)hctl)->lpmodel;
   npPropInfo = lpmodel->npproplist[iProp];
   
   // If PF_fPropArray flag is set for the property prepare a DATASTRUCT
   if ( npPropInfo->fl & PF_fPropArray ) {
      if ( lIndex == -1 )
         return WBERR_PROPARRAY_INDEXNEEDED;
         
      data.cindex = 1;
      data.index[0].datatype = DT_SHORT;
      data.index[0].data = lIndex;
      if ( (err = WBGetControlProperty( hctl, iProp, (LPVOID)&data )) != 0 )
         return err;
      lValue = data.data;
   }
   else {
      if ( lIndex != -1L )
         return WBERR_NOTAPROPARRAY;
      if ( (err = WBGetControlProperty( hctl, iProp, &lValue )) != 0 )
         return err;
   }
   
      
   switch ( npPropInfo->fl & (LONG)PF_datatype ) {
   
      case DT_HLSTR:
         {  
            HLSTR hlstr;
            hlstr = WBCreateTempHlstr( NULL, 0 );
            err = WBSetHlstr( &hlstr, (LPVOID)(HLSTR)lValue, (USHORT)-1);
            if ( err != 0 )
               return err;
            lpResult->var.tString = hlstr;
            lpResult->type = VT_STRING;
         }
         break;
      
      case DT_BOOL:
         lpResult->var.tLong = ( lValue == 0 ? WB_FALSE : WB_TRUE );
         lpResult->type = VT_I4;
         break;
         
      case DT_SHORT:
      case DT_LONG:
         lpResult->var.tLong = lValue;
         lpResult->type = VT_I4;
         break;
         
      default:
         assert( FALSE );
         return WBERR_INTERNAL_ERROR;
   }
   
   return 0;
}

/*--------------------------------------------------------------------------*/   
/*                                                                          */
/*--------------------------------------------------------------------------*/   
ERR WBSetControlProperty( HCTL hctl, USHORT i, LONG Data )
{
   LPMODEL lpmodel;
   NPPROPINFO npPropInfo;
   LPVOID lpv;
   ERR err;
   
   lpmodel = (LPMODEL)((LPOBJECTINSTANCEINFO)hctl)->lpmodel;
   npPropInfo = lpmodel->npproplist[i];
   
   // Check wether setting the property is permitted
   if ( lpmodel->fInitialized && (npPropInfo->fl & (LONG)PF_fNoRuntimeW) )
      return WBERR_READONLYPROPERTY;
   
   // Send to control procedure if PF_fSetMsg is set for the property
   if ( npPropInfo->fl & PF_fSetMsg ) 
      return lpmodel->pctlproc( hctl, VBM_SETPROPERTY, i, Data );
   
   // Must always use PF_fSetMsg for property arrays
   assert( (npPropInfo->fl & PF_fPropArray) == 0 );
   
   // Set directly if PF_fSetData flag is set for the property
   if ( npPropInfo->fl & PF_fSetData ) {
   
      // point to property value
      lpv = (LPVOID)((LPBYTE)lpmodel + sizeof(MODEL) + npPropInfo->offsetData); 
      
      switch ( npPropInfo->fl & (LONG)PF_datatype ) {
      
         case DT_HLSTR:
            {
               HLSTR hlstr;
               // Destroy old HLSTR if property value != NULL
               hlstr = *((HLSTR FAR*)lpv);
               if ( hlstr != NULL )
                  WBDestroyHlstr( hlstr );
               // Make copy of new HLSTR
               hlstr = WBCreateHlstr( NULL, 0 );
               if ( hlstr == NULL )
                  return WBERR_STRINGSPACE;
               if ( (err = WBSetHlstr( &hlstr, (LPVOID)(HLSTR)Data, (USHORT)-1)) != 0 )
                  return err;
               // Assign to property
               *((HLSTR FAR*)lpv) = hlstr;
            }
            break;
         
         case DT_BOOL:
            *((BOOL FAR*)lpv) = (BOOL)Data;
            break;
            
         case DT_SHORT:
            *((USHORT FAR*)lpv) = (USHORT)Data;
            break;
            
         case DT_LONG:
            *((LPLONG)lpv) = (LONG)Data;
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
ERR WBGetControlProperty( HCTL hctl, USHORT i, LPVOID lpdata )
{
   LPMODEL lpmodel;
   NPPROPINFO npPropInfo;
   LPVOID lpv;
   
   lpmodel = (LPMODEL)((LPOBJECTINSTANCEINFO)hctl)->lpmodel;
   npPropInfo = lpmodel->npproplist[i];
   
   if ( lpmodel->fInitialized && (npPropInfo->fl & (LONG)PF_fNoRuntimeR) )
      return WBERR_WRITEONLYPROPERTY;
   
   // Send to control procedure if PF_fGetMsg is set for the property
   if ( npPropInfo->fl & PF_fGetMsg ) 
      return lpmodel->pctlproc( hctl, VBM_GETPROPERTY, i, (LONG)lpdata );

   // Must always use PF_fGetMsg for property arrays
   assert( (npPropInfo->fl & PF_fPropArray) == 0 );
   
   // Set directly if PF_fGetData flag is set for the property
   if ( npPropInfo->fl & PF_fGetData ) {
      
      // point to property value
      lpv = (LPVOID)((LPBYTE)lpmodel + sizeof(MODEL) + npPropInfo->offsetData); 
      
      switch ( npPropInfo->fl & (LONG)PF_datatype ) {
      
         case DT_HLSTR:
            *((LPLONG)lpdata) = (LONG)*((HLSTR FAR*)lpv);
            break;
         
         case DT_BOOL:
            *((LPLONG)lpdata) = (LONG)*((BOOL FAR*)lpv);
            break;
            
         case DT_SHORT:
            *((LPLONG)lpdata) = (LONG)*((USHORT FAR*)lpv);
            break;
            
         case DT_LONG:
            *((LPLONG)lpdata) = (LONG)*((LPLONG)lpv);
            break;
            
         default:
            assert( FALSE );
            return WBERR_INTERNAL_ERROR;
      }
   }
   else {
      lpdata = NULL;
   }
   
   return 0;
}
      
