/* PivotTableManager.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Apr 19, 2012 11:58:47 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.model.impl;

import java.util.List;

import org.zkoss.poi.ss.usermodel.PivotCache;
import org.zkoss.poi.ss.usermodel.PivotTable;
import org.zkoss.poi.ss.util.CellReference;

/**
 * @author sam
 *
 */
public interface PivotTableManager {
	
	public List<PivotTable> getPivotTables();
	
	public PivotTable createPivotTable(CellReference destination, String name, PivotCache pivotCache);
}
