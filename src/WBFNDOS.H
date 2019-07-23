/* wbfndos.h */

ERR FAR PASCAL WBFnChDir( int narg, LPVAR lparg, LPVAR lpResult );
ERR FAR PASCAL WBFnChDrive( int narg, LPVAR lparg, LPVAR lpResult );
ERR FAR PASCAL WBFnCurDir( int narg, LPVAR lparg, LPVAR lpResult );
ERR FAR PASCAL WBFnDir( int narg, LPVAR lparg, LPVAR lpResult );
ERR FAR PASCAL WBFnDirExist( int narg, LPVAR lparg, LPVAR lpResult );
ERR FAR PASCAL WBFnFileList( int narg, LPVAR lparg, LPVAR lpResult );
ERR FAR PASCAL WBFnEnviron( int narg, LPVAR lparg, LPVAR lpResult );
ERR FAR PASCAL WBFnEnvironList( int narg, LPVAR lparg, LPVAR lpResult );
ERR FAR PASCAL WBFnFileSearch( int narg, LPVAR lparg, LPVAR lpResult );
ERR FAR PASCAL WBFnGetDiskFreeSpace( int narg, LPVAR lparg, LPVAR lpResult );
ERR FAR PASCAL WBFnGetDiskParms( int narg, LPVAR lparg, LPVAR lpResult );
ERR FAR PASCAL WBFnGetDriveType( int narg, LPVAR lparg, LPVAR lpResult );
ERR FAR PASCAL WBFnGetDosVersion( int narg, LPVAR lparg, LPVAR lpResult );
ERR FAR PASCAL WBFnGetExtMemSize( int narg, LPVAR lparg, LPVAR lpResult );
ERR FAR PASCAL WBFnGetWinVersion( int narg, LPVAR lparg, LPVAR lpResult );
ERR FAR PASCAL WBFnMkDir( int narg, LPVAR lparg, LPVAR lpResult );
ERR FAR PASCAL WBFnReadROMBios( int narg, LPVAR lparg, LPVAR lpResult );
ERR FAR PASCAL WBFnRmDir( int narg, LPVAR lparg, LPVAR lpResult );
ERR FAR PASCAL WBFnTreeDelete( int narg, LPVAR lparg, LPVAR lpResult );
