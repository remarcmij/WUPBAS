/* wbfileio.h */

BOOL WBRegisterModel_FileObject( void );
ERR WBOpenFile( OPENFILESTRUCT FAR* lpof );
void WBCloseFile( OPENFILESTRUCT FAR* lpof );
BOOL WBEof( OPENFILESTRUCT FAR* lpof );
ERR WBReadString( OPENFILESTRUCT FAR* lpof, HLSTR FAR* phlstr );
ERR WBWriteString( OPENFILESTRUCT FAR* lpof, HLSTR hlstr );
