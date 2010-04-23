/* WidgetHandler.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Apr 21, 2008 2:16:21 PM     2008, Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui.sys;

import java.io.IOException;
import java.io.Writer;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.Widget;


/**
 * WidgetHandler control when widgets should display on client, when should remove.
 * WidgetHandler and Widget are implementation-sensitive. 
 * @author Dennis.Chen
 *
 */
public interface WidgetHandler{

	/**
	 * initial a widget handler. 
	 * this method will be invoked by spreadsheet, you should not call this method directly.
	 * @param spreadsheet
	 */
	public void init(Spreadsheet spreadsheet);
	
	/**
	 * get spreadsheet of this handler
	 * @return spreadsheet
	 */
	public Spreadsheet getSpreadsheet();
	
	/**
	 * Add widget to a handler, 
	 * notice : WidgetHandler and Widget are implementation-sensitive. 
	 * @param widget a widget
	 * @return true if success to add a widget
	 */
	public boolean addWidget(Widget widget);
	
	/**
	 * Remove widget from a handler
	 * notice : WidgetHandler and Widget are implementation-sensitive.
	 * @param widget
	 * @return true if success to remove a widget
	 */
	public boolean removeWidget(Widget widget);
	
	/**
	 * call when spreadsheet try to load a block of cell to client side. 
	 * handler should take care this method and load corresponding widgets, which in the block , to client side.
	 * this method will be invoked by spreadsheet, you should not call this method directly.
	 */
	public void onLoadOnDemand(Sheet sheet,int left,int top,int right,int bottom);
	
	/**
	 * invalidate this handle, which means all widget in client side will be remove, and need to re-drew.
	 * this method will be invoked when spreadsheet invalidate only.
	 */
	public void invaliate();
}
