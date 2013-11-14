/* NBook.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/11/14 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngmodel;

/**
 * @author dennis
 *
 */
public interface NBook {

//	NBookSeries getBookSeries();
	
	NSheet getSheetAt(int i);
	
	int getNumOfSheet();
	
	NSheet getSheetByName(String name);
	//editable 
	NSheet createSheet(NSheet src);
	void setSheetName(NSheet sheet, String newname);
	void deleteSheet(NSheet sheet);
	void moveSheetTo(NSheet sheet, int index);
}
