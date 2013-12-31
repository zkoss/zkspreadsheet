/* NBook.java

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
package org.zkoss.zss.ngmodel;

import java.util.List;
import java.util.Map;

import org.zkoss.zss.ngmodel.util.CellStyleMatcher;
import org.zkoss.zss.ngmodel.util.FontMatcher;

/**
 * @author dennis
 * @since 3.5.0
 */
public interface NBook {

	/**
	 * Get the book name, a book name is unique for book in {@link NBookSeries}
	 * @return book name;
	 */
	public String getBookName();
	
	/**
	 * Get the book series, it contains a group of book that might refer to other by book name
	 * @return book series
	 */
	public NBookSeries getBookSeries();
	/**
	 * Get sheet at the index
	 * @param idx the sheet index
	 * @return the sheet at the index
	 */
	public NSheet getSheet(int idx);
	
	/**
	 * Get the index of sheet
	 * @param sheet the sheet
	 * @return the index
	 */
	public int getSheetIndex(NSheet sheet);
	
	/**
	 * Get the number of sheet
	 * @return the number of sheet
	 */
	public int getNumOfSheet();
	
	/**
	 * Get the sheet by name
	 * @param name the name of sheet
	 * @return the sheet, or null if not found
	 */
	public NSheet getSheetByName(String name);
	
	/**
	 * Get the sheet by id
	 * @param id the id of sheet
	 * @return the sheet, or null if not found
	 */
	public NSheet getSheetById(String id);
	
	/**
	 * Create a sheet
	 * @param name the name of sheet
	 * @return the sheet
	 */
	public NSheet createSheet(String name);
	
	/**
	 * Get all sheets
	 * @return an unmodifiable sheet list
	 */
	public List<NSheet> getSheets();
	
	/**
	 * Create a sheet and copy the contain form the sheet sheet
	 * @param name the name of sheet
	 * @param src the source sheet to copy
	 * @return the sheet
	 */
	public NSheet createSheet(String name, NSheet src);
	
	/**
	 * Set the sheet to a new name
	 * @param sheet the sheet
	 * @param newname the new name
	 */
	public void setSheetName(NSheet sheet, String newname);
	
	/**
	 * Delete the sheet
	 * @param sheet the sheet
	 */
	public void deleteSheet(NSheet sheet);
	
	/**
	 * Move the sheet to new position
	 * @param sheet the sheet
	 * @param index the new position
	 */
	public void moveSheetTo(NSheet sheet, int index);
	
	/**
	 * Get the default style of this book
	 * @return
	 */
	public NCellStyle getDefaultCellStyle();

	/**
	 *Create a cell style
	 * @param inStyleTable if true, the new created style will be stored inside this book, 
	 * then you can use {@link #searchCellStyle(CellStyleMatcher)} to search and reuse this style.
	 * @return 
	 */
	public NCellStyle createCellStyle(boolean inStyleTable);
	
	/**
	 * Create a cell style and copy the style from the src style.
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
	
	
	public NFont getDefaultFont();

	public NFont createFont(boolean inFontTable);
	
	public NFont createFont(NFont src,boolean inFontTable);
	
	public NFont searchFont(FontMatcher matcher);
	
	public NColor createColor(byte r, byte g, byte b);
	public NColor createColor(String htmlColor);
	
	/**
	 * Get the max row size of this book
	 */
	public int getMaxRowSize();
	
	/**
	 * Get the max column size of this book
	 */
	public int getMaxColumnSize();
	
	/**
	 * add event listener to this book
	 * @param listener the listener
	 */
	public void addEventListener(ModelEventListener listener);
	
	/**
	 * remove event listener from this book
	 * @param listener the listener
	 */
	public void removeEventListener(ModelEventListener listener);
	
	/**
	 * Get the runtime custom attribute that stored in this book
	 * @param name the attribute name
	 * @return the value, or null if not found
	 */
	public Object getAttribute(String name);
	
	/**
	 * Set the runtime custom attribute to stored in this book, the attribute is only use for developer to stored runtime data in the book,
	 * values will not stored to excel when exporting.
	 * @param name name the attribute name
	 * @param value the attribute value
	 */
	public Object setAttribute(String name,Object value);
	
	/**
	 * Get the unmodifiable runtime attributes map
	 * @return
	 */
	public Map<String,Object> getAttributes();
	
	public NName createName(String name,String applyToSheetName);
	public NName createName(String name);
	
	public void setNameName(NName name,String newname, String applyToSheetName);
	public void setNameName(NName name,String newname);
	
	public void deleteName(NName name);
	
	public int getNumOfName();
	
	public NName getName(int idx);
	
	public NName getNameByName(String namename, String sheetName);
	public NName getNameByName(String namename);
	
	public List<NName> getNames();

	public void setShareScope(String scope);

	public String getShareScope();

	/**
	 * Optimize CellStyle, usually called when export book or after many style operation 
	 * @return
	 */
	public void optimizeCellStyle();
}
