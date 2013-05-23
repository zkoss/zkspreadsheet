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
import java.io.IOException;
import java.util.List;

import org.zkoss.image.AImage;
//import org.zkoss.poi.ss.usermodel.Cell;
//import org.zkoss.poi.ss.usermodel.charts.ChartType;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Chart;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.app.file.SpreadSheetMetaInfo;
//import org.zkoss.zss.model.sys.XBook;
//import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.ui.Position;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * @author Sam
 *
 */
public interface WorkbookCtrl {
	
	public final static int HEADER_TYPE_ROW = 0;
	public final static int HEADER_TYPE_COLUMN = 1;
	
	public void reGainFocus();
	/**
	 * Retrieve client side spreadsheet focus. 
	 * 
	 * @param row
	 * @param column
	 */
	public void focusTo(int row, int column, boolean fireFocusEvent);
	
	/**
	 * Return current cell(row,column) focus position
	 */
	public Position getCellFocus();
	
	/**
	 * Add and move other editor's focus
	 * 
	 * @param id
	 * @param name
	 * @param color
	 * @param row
	 * @param col
	 */
	public void moveEditorFocus(String id, String name, String color, int row ,int col);
	
	/**
	 * Remove editor's focus on specified name
	 * 
	 * @param name
	 */
	public void removeEditorFocus(String name);
	
	/**
	 * Return current selection rectangle only if onCellSelection event listener is registered.
	 * @return
	 */
	public Rect getSelection();
	
	/**
	 * Returns the maximum visible number of columns of this spreadsheet
	 * 
	 * @return
	 */
	public int getMaxcolumns();
	
	/**
	 * Returns the maximum visible number of rows of this spreadsheet
	 * 
	 * @return
	 */
	public int getMaxrows();
	
	/**
	 * Clear copy source or cut source of workbook
	 */
	public void clearClipbook();
	
	public void renameSelectedSheet(String name);
	
	public Sheet getSelectedSheet();
	
	public void setSelectedSheet(String name);
	
	public void insertSheet();
	
	public void addImage(int row, int col, AImage image);
	
	public void insertFormula(int rowIdx, int colIdx, String formula);
		
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
	
	/**
	 * Add an event listener.
	 * @param evtnm
	 * @param listener
	 */
	public void addEventListener(String evtnm, EventListener listener);
	
	/**
	 * Removes an event listener.
	 * @param evtnm
	 * @param listener
	 * @return
	 */
	public boolean removeEventListener(String evtnm, EventListener listener);
	
	public String getCurrentCellPosition();
	
	/**
	 * Returns the reference 
	 * @return
	 */
	public String getReference (int row, int column);
	
//	/**
//	 * 
//	 * @param cell
//	 * @param text
//	 */
//	public void escapeAndUpdateText(int row, int column, String text);
	
//	public void updateText(int row, int column, String text);

	public void setDataFormat(String format);
	
	public List<String> getSheetNames();
	
	/**
	 * Sets {@link Spreadsheet} book src
	 * @param src
	 */
	public void setBookSrc(String src);
	
	public void setBook(Book book);
	
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
	
//	public void setSrcName(String src);
	
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
	public ByteArrayOutputStream exportToExcel()  throws IOException ;

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
	public void setColumnWidthInPx(int width, Rect selection);

	/**
	 * Sets row height of selected rows
	 * @param height
	 */
	public void setRowHeightInPx(int height, Rect selection);
	
//	public int getDefaultCharWidth();
	
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
	 * <p> Returns -1 if delete sheet fail
	 * @return index
	 */
	public int deleteSheet();
	
	public void addChart(int row, int col, Chart.Type type);
	
	public String getColumnTitle(int col);
	
	public String getRowTitle(int row);

	public Rect getVisibleRect();
	
//	public boolean setEditTextWithValidation(Sheet sheet, int row, int col, String txt, EventListener callback);
}