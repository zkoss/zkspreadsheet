/* ActiveRangeHelper.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 30, 2012 3:44:46 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ui.impl;

import java.util.HashMap;

import org.zkoss.zss.model.sys.Worksheet;
import org.zkoss.zss.ui.Rect;

/**
 * @author sam
 *
 */
public class ActiveRangeHelper {

	private HashMap<Worksheet, Rect> activeRanges = new HashMap<Worksheet, Rect>();
	
	public void setActiveRange(Worksheet sheet, int tRow, int lCol, int bRow, int rCol) {
		Rect rect = activeRanges.get(sheet);
		if (rect == null) {
			activeRanges.put(sheet, rect = new Rect(lCol, tRow, rCol, bRow));
		} else {
			rect.set(lCol, tRow, rCol, bRow);
		}
	}
	
	public Rect getRect(Worksheet sheet) {
		return activeRanges.get(sheet);
	}
	
	public boolean containsSheet(Worksheet sheet) {
		return activeRanges.containsKey(sheet);
	}
	
	public boolean contains(Worksheet sheet, int row, int col) {
		return contains(sheet, row, col, row, col);
	}
	
	public boolean contains(Worksheet sheet, int tRow, int lCol, int bRow, int rCol) {
		Rect rect = activeRanges.get(sheet);
		if (rect == null)
			return false;
		return rect.contains(tRow, lCol, bRow, rCol);
	}
}
