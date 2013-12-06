/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngapi;

import java.util.Locale;

import org.zkoss.zss.ngmodel.NSheet;
/**
 * The most useful api to manipulate a book
 *  
 * @author dennis
 * @since 3.5.0
 */
public interface NRange {

	NSheet getSheet();
	int getRow();
	int getColumn();
	int getLastRow();
	int getLastColumn();
	
	void setEditText(String editText);
	void setValue(Object value);
	void clear();
	void setLocale(Locale locale);
	Locale getLocale();
}
