/* wbdllcal.h */

ERR WBDeclareStatement( void );
ERR WBCallDeclaredFunction( NPWBDECLARE npfun, LPVAR lpResult );
void WBDestroyDeclareDefs( void );
ERR FAR PASCAL WBFnTruncString( int narg, LPVAR lparg, LPVAR lpResult );
