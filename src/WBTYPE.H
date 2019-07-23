/* wbtype.h */

typedef struct tagWBTYPEMEM {
   NPWBUSERTYPE npUserType;
   HMEM hmem;
   BYTE data[1];
} WBTYPEMEM;
typedef WBTYPEMEM FAR* LPWBTYPEMEM;

ERR WBTypeStatement( void );
NPWBTYPEMEMBER WBFindTypeMember( NPWBUSERTYPE npUserType, LPSTR lpszName );
LPWBTYPEMEM WBAllocTypeMem( NPWBUSERTYPE npUserType, BOOL fInitData );
void WBFreeTypeMem( LPWBTYPEMEM lpTypeMem );
ERR WBGetUserTypeData( LPWBTYPEMEM lpTypeMem, 
                       NPWBUSERTYPE FAR* lpnpUserType, 
                       UINT FAR* lpfMemberType,
                       LPBYTE FAR* lplpData,
                       UINT FAR* lpuSize );
                       
