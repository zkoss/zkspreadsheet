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

import org.zkoss.util.media.Media;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zss.app.file.SpreadSheetMetaInfo;

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
	
	public void setBookSrc(String src);
	
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
	
	public void newBook();
	
	public void openBook(SpreadSheetMetaInfo info);
	
	public boolean hasBook();
	
	public boolean hasFileName();

	public ByteArrayOutputStream exportToExcel();

	public String getBookName();
	
	/**
	 * Subscribe book event listener
	 * @param listener
	 */
	public void subscribe(EventListener listener);
	
	/**
	 * Unsubscribe book event listener
	 * @param listener
	 */
	public void unsubscribe(EventListener listener);

	/**
	 * @param width
	 */
	public void setColumnWidthInPx(int width);

	/**
	 * @param height
	 */
	public void setRowHeightInPx(int height);
	
	public int getDefaultCharWidth();
	
}
