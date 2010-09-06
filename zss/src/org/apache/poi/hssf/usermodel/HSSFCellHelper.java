/* HSSFCellHelper.java

	Purpose:
		
	Description:
		
	History:
		Sep 3, 2010 12:59:35 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.apache.poi.hssf.usermodel;

import org.apache.poi.hssf.record.CellValueRecordInterface;

/**
 * A helper class to make HSSFCell package method visible.
 * @author henrichen
 *
 */
public class HSSFCellHelper {
	private HSSFCell _cell;
	public HSSFCellHelper(HSSFCell cell) {
		_cell = cell;
	}
	public CellValueRecordInterface getCellValueRecord() {
		return _cell.getCellValueRecord();
	}
}
