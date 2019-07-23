/* wbslist.h */

ERR WBSListAppendItem( HCTL hctl, LPVAR lpvar ); 
ERR WBSListAppendLine( HCTL hctl, LPVOID lpv, USHORT cb );
void WBSListSetModifiedFlag( HCTL hctl, BOOL fModified );
