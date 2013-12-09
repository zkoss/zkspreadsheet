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

import org.zkoss.zss.api.AreaRef;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.ngmodel.NSheet;

/**
 * @author sam
 *
 */
public class ActiveRangeHelper {

	private HashMap<NSheet, AreaRef> activeRanges = new HashMap<NSheet, AreaRef>();
	
	public void setActiveRange(NSheet sheet, int tRow, int lCol, int bRow, int rCol) {
		AreaRef rect = activeRanges.get(sheet);
		if (rect == null) {
			activeRanges.put(sheet, rect = new AreaRef(tRow, lCol, bRow, rCol));
		} else {
			rect.setArea(tRow, lCol, bRow, rCol);
		}
	}
	
	public AreaRef getArea(NSheet sheet) {
		return activeRanges.get(sheet);
	}
	
	public boolean containsSheet(NSheet sheet) {
		return activeRanges.containsKey(sheet);
	}
	
	public boolean contains(NSheet sheet, int row, int col) {
		return contains(sheet, row, col, row, col);
	}
	
	public boolean contains(NSheet sheet, int tRow, int lCol, int bRow, int rCol) {
		AreaRef rect = activeRanges.get(sheet);
		if (rect == null)
			return false;
		return rect.contains(tRow, lCol, bRow, rCol);
	}
}
