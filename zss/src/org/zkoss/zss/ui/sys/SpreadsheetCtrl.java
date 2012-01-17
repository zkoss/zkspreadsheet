/* SpreadsheetCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Dec 18, 2007 12:18:09 PM     2007, Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui.sys;


//import org.zkoss.zss.model.Sheet;
import org.zkoss.json.JSONObject;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Widget;
import org.zkoss.zss.ui.impl.HeaderPositionHelper;
import org.zkoss.zss.ui.impl.MergeMatrixHelper;


/**
 * Special controller interface .
 * Only spreadsheet developer need to use this interface. 
 * @author Dennis.Chen
 *
 */
public interface SpreadsheetCtrl {

	final public static String CHILD_PASSING_KEY = "zsschildren";
	
	//TODO: measure best load size
	public static final int DEFAULT_LOAD_COLUMN_SIZE = 35;
	public static final int DEFAULT_LOAD_ROW_SIZE = 45;
	
	public enum CellAttribute {
		ALL(1), TEXT(2), STYLE(3), SIZE(5), MERGE(5);
		
		int value;
		CellAttribute(int value) {
			this.value = value;
		}
		
		public String toString() {
			return "" + value;
		}
	}
	
	public enum Header {
		NONE, ROW, COLUMN, BOTH;
		
	}
	
	public HeaderPositionHelper getRowPositionHelper(String sheetId);
	
	public HeaderPositionHelper getColumnPositionHelper(String sheetId);
	
	public MergeMatrixHelper getMergeMatrixHelper(Worksheet sheet);
	
	
	public Rect getSelectionRect();
	public Rect getFocusRect();
	public Rect getLoadedRect();
	public Rect getVisibleRect();
	
	public WidgetHandler getWidgetHandler();
	
	public JSONObject getRowHeaderAttrs(Worksheet sheet, int rowStart, int rowEnd);
	
	public JSONObject getColumnHeaderAttrs(Worksheet sheet, int colStart, int colEnd);
	
	public JSONObject getRangeAttrs(Worksheet sheet, Header containsHeader, CellAttribute type, int left, int top, int right, int bottom);
	
	public String getDataPanelAttrs();
	
	
	/**
	 * Add widget to the {@link WidgetHandler} of this spreadsheet, 
	 * 
	 * @param widget a widget
	 * @return true if success to add a widget
	 */
	public boolean addWidget(Widget widget);
	
	/**
	 * Remove widget from the {@link WidgetHandler} of this spreadsheet, 
	 * @param widget
	 * @return true if success to remove a widget
	 */
	public boolean removeWidget(Widget widget);
	
	public Boolean getTopHeaderHiddens(int col);
	
	public Boolean getLeftHeaderHiddens(int row);
}
