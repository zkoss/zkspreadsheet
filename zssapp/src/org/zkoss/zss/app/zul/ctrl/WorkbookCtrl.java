/* WorkbookCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 25, 2010 10:39:37 AM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul.ctrl;

import java.io.ByteArrayOutputStream;
import java.util.List;

import org.zkoss.util.media.Media;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zss.app.file.SpreadSheetMetaInfo;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * @author Sam
 *
 */
public interface WorkbookCtrl {
	
	public final static int HEADER_TYPE_ROW = 0;
	public final static int HEADER_TYPE_COLUMN = 1;
	
	public void reGainFocus();
	
	public void renameSelectedSheet(String name);
	
	public void setSelectedSheet(String name);
	
	public void insertSheet();
	
	public void insertImage(Media media);
	
	public void insertFormula(String formula);
	
	//TODO: use clip board interface, return  
	public void cutSelection();
	
	public void copySelection();
	
	public void pasteSelection();
	
	public void clearSelectionContent();
	
	public void clearSelectionStyle();
	
	//TODO: insert row below
	public void insertRowAbove();
	
	public void deleteRow();
	
	//TODO: insert column right
	public void insertColumnLeft();
	
	public void deleteColumn();
	
	public void hide(boolean hide);
	
	public void sort(boolean isSortDescending);
	
	public void shiftCell(int direction);
	
	public void setRowFreeze(int rowfreeze);
	
	public void setColumnFreeze(int columnFreeze);
	
	public void addEventListener(String evtnm, EventListener listener);
	
	public boolean removeEventListener(String evtnm, EventListener listener);
	
	public String getCurrentCellPosition();
	
	public void setDataFormat(String format);
	
	public List<String> getSheetNames();
	
	/**
	 * Sets {@link Spreadsheet} book src
	 * @param src
	 */
	public void setBookSrc(String src);
	
	/**
	 * Returns whether current sheet is protected or not
	 * @return
	 */
	public boolean isSheetProtect();
	
    /**
     * Sets the protection enabled as well as the password
     * @param password to set for protection. Pass <code>null</code> to remove protection
     */
	public void protectSheet(String password);
	
	/**
	 * Save book
	 */
	public void save();
	
	/**
	 * Close book
	 */
	public void close();
	
	public String getSrc();
	
	public void setSrcName(String src);
	
	/**
	 * Open a new empty work sheet of current {@link Spreadsheet}
	 */
	public void newBook();
	
	/**
	 * Open a work book of current {@link Spreadsheet}
	 * @param spreadSheetMetaInfo
	 */
	public void openBook(SpreadSheetMetaInfo spreadSheetMetaInfo);
	
	/**
	 * Returns whether {@link Spreadsheet} has book or not
	 * @return
	 */
	public boolean hasBook();
	
	/**
	 * Returns whether {@link Spreadsheet} has file name or not
	 * <p> if {@link Spreadsheet} open new empty work sheet will return false
	 * @return boolean
	 */
	public boolean hasFileExtentionName();

	/**
	 * Exports spreadsheet to excel file
	 * @return
	 */
	public ByteArrayOutputStream exportToExcel();

	/**
	 * Returns {@link Spreadsheet} book name
	 * @return
	 */
	public String getBookName();
	
	/**
	 * Add a {@link #EventListener} to this book
	 * @param listener
	 */
	public void addBookEventListener(EventListener listener);
	
	/**
	 * Remove {@link #EventListener} the specified listener from listening to this book
	 * @param listener
	 */
	public void removeBookEventListener(EventListener listener);

	/**
	 * Sets column width of selected columns
	 * @param width
	 */
	public void setColumnWidthInPx(int width);

	/**
	 * Sets row height of selected rows
	 * @param height
	 */
	public void setRowHeightInPx(int height);
	
	public int getDefaultCharWidth();
	
	/**
	 * Shifts current selected sheet left, returns shifted sheet index
	 * <p> Returns -1 if fail to shift sheet
	 * @return index
	 */
	public int shiftSheetLeft();
	
	/**
	 * Shifts current selected sheet right, returns shifted sheet index
	 * <p> Returns -1 if fail to shift sheet
	 * @return index
	 */
	public int shiftSheetRight();
	
	/**
	 * Deletes current selected sheet, returns next sheet index
	 * <> Returns -1 if delete sheet fail
	 * @return index
	 */
	public int deleteSheet();
}