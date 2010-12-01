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

import org.zkoss.util.media.Media;
import org.zkoss.zk.ui.event.EventListener;

/**
 * @author Sam
 *
 */
public interface WorkbookCtrl {
	
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
}
