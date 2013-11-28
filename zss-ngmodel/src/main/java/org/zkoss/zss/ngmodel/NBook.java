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

import org.zkoss.zss.ngmodel.NChart.NChartType;
import org.zkoss.zss.ngmodel.chart.NChartData;
import org.zkoss.zss.ngmodel.util.CellStyleMatcher;

/**
 * @author dennis
 *
 */
public interface NBook {

	public String getBookName();
	
	public NBookSeries getBookSeries();
	/**
	 * get sheet at the index
	 * @param idx the sheet index
	 * @return the sheet at the index
	 */
	public NSheet getSheet(int idx);
	
	/**
	 * get the number of sheet inside this book
	 * @return the number of sheet
	 */
	public int getNumOfSheet();
	
	/**
	 * get the sheet by name
	 * @param name the name of sheet
	 * @return
	 */
	public NSheet getSheetByName(String name);
	
	//editable
	public NSheet createSheet(String name);
	public NSheet createSheet(String name, NSheet src);
	public void setSheetName(NSheet sheet, String newname);
	public void deleteSheet(NSheet sheet);
	public void moveSheetTo(NSheet sheet, int index);
	
	public NCellStyle getDefaultCellStyle();

	/**
	 * create a cell style
	 * @param inStyleTable if true, the new created style will be stored inside this book, 
	 * then you can use {@link #searchCellStyle(CellStyleMatcher)} to search and reuse this style.
	 * @return 
	 */
	public NCellStyle createCellStyle(boolean inStyleTable);
	
	/**
	 * create a cell style and copy the style from the src style.
	 * @param src the source style to copy from.
	 * @param inStyleTable if true, the new created style will be stored inside this book, 
	 * then you can use {@link #searchCellStyle(CellStyleMatcher)} to search and reuse this style.
	 * @return 
	 */
	public NCellStyle createCellStyle(NCellStyle src,boolean inStyleTable);
	
	/**
	 * Search the style table and return the first matched style. 
	 * @param matcher the style matcher
	 * @return the matched style.
	 */
	public NCellStyle searchCellStyle(CellStyleMatcher matcher);
	
	
	public int getMaxRowSize();
	
	public int getMaxColumnSize();
	
	public void addEventListener(ModelEventListener listener);
	
	public void removeEventListener(ModelEventListener listener);
}
