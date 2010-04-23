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
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Widget;
import org.zkoss.zss.ui.impl.HeaderPositionHelper;
import org.zkoss.zss.ui.impl.MergeMatrixHelper;

import org.apache.poi.ss.usermodel.Sheet;

/**
 * Special controller interface .
 * Only spreadsheet developer need to use this interface. 
 * @author Dennis.Chen
 *
 */
public interface SpreadsheetCtrl {

	final public static String CHILD_PASSING_KEY = "zsschildren";
	
	public HeaderPositionHelper getRowPositionHelper(String sheetId);
	
	public HeaderPositionHelper getColumnPositionHelper(String sheetId);
	
	public MergeMatrixHelper getMergeMatrixHelper(Sheet sheet);
	
	
	public Rect getSelectionRect();
	public Rect getFocusRect();
	public Rect getLoadedRect();
	
	public WidgetHandler getWidgetHandler();
	
	public String getCellOuterAttrs(int row,int col);

	public String getCellInnerAttrs(int row,int col);

	public String getRowOuterAttrs(int row);

	public String getTopHeaderOuterAttrs(int col);
	
	public String getTopHeaderInnerAttrs(int col);

	public String getLeftHeaderOuterAttrs(int row);
	public String getLeftHeaderInnerAttrs(int row);
	
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

}
