/* wbapi.h */

#ifdef __cplusplus
extern "C"
{
#endif

typedef LPVOID (CALLBACK* LPNEXTLINECALLBACK)( LPSTR, UINT );
typedef void (CALLBACK* LPGOTOLINECALLBACK)( LPVOID );
typedef int (CALLBACK* LPCMDHANDLERCALLBACK)( LPSTR );
typedef void (CALLBACK* LPBROADCASTCALLBACK)( UINT, WORD, LONG );

int WINAPI WBInitialize( HWND );
UINT WINAPI WBAboutString( LPSTR, UINT );
int WINAPI WBRegisterLineCallbacks( LPNEXTLINECALLBACK, LPGOTOLINECALLBACK );
int WINAPI WBRegisterCmdCallback( LPCMDHANDLERCALLBACK );
int WINAPI WBRun( void );
int WINAPI WBRunFile( LPSTR lpszFileName );
int WINAPI WBDeInitialize( void );
BOOL WINAPI WBSetLongVariable( LPCSTR, long );
BOOL WINAPI WBSetStringVariable( LPCSTR, LPCSTR );
void WINAPI WBRegisterBroadcastHandler( LPBROADCASTCALLBACK );

#ifdef __cplusplus
}
#endif
