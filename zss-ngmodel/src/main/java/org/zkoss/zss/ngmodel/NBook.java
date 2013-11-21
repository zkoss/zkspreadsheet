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

import org.zkoss.zss.ngmodel.util.CellStyleMatcher;

/**
 * @author dennis
 *
 */
public interface NBook {

//	NBookSeries getBookSeries();
	/**
	 * get sheet at the index
	 * @param idx the sheet index
	 * @return the sheet at the index
	 */
	NSheet getSheet(int idx);
	
	/**
	 * get the number of sheet inside this book
	 * @return the number of sheet
	 */
	int getNumOfSheet();
	
	/**
	 * get the sheet by name
	 * @param name the name of sheet
	 * @return
	 */
	NSheet getSheetByName(String name);
	
	//editable
	NSheet createSheet(String name);
	NSheet createSheet(String name, NSheet src);
	void setSheetName(NSheet sheet, String newname);
	void deleteSheet(NSheet sheet);
	void moveSheetTo(NSheet sheet, int index);
	
	NCellStyle getDefaultCellStyle();

	/**
	 * create a cell style
	 * @param inStyleTable if true, the new created style will be stored inside this book, 
	 * then you can use {@link #searchCellStyle(CellStyleMatcher)} to search and reuse this style.
	 * @return 
	 */
	NCellStyle createCellStyle(boolean inStyleTable);
	
	/**
	 * create a cell style and copy the style from the src style.
	 * @param src the source style to copy from.
	 * @param inStyleTable if true, the new created style will be stored inside this book, 
	 * then you can use {@link #searchCellStyle(CellStyleMatcher)} to search and reuse this style.
	 * @return 
	 */
	NCellStyle createCellStyle(NCellStyle src,boolean inStyleTable);
	
	/**
	 * Search the style table and return the first matched style. 
	 * @param matcher the style matcher
	 * @return the matched style.
	 */
	NCellStyle searchCellStyle(CellStyleMatcher matcher);
	
	
	int getMaxRowSize();
	
	int getMaxColumnSize();
}
