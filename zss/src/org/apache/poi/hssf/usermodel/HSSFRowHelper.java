/* HSSFRowHelper.java

	Purpose:
		
	Description:
		
	History:
		Jun 4, 2010 10:15:19 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.apache.poi.hssf.usermodel;

import org.apache.poi.hssf.record.CellValueRecordInterface;

/**
 * A helper class to make HSSFRow package method visible.
 * @author henrichen
 *
 */
public class HSSFRowHelper {
	final private HSSFRow _row;
	public HSSFRowHelper(HSSFRow row) {
		_row = row;
	}
	public HSSFRow getRow() {
		return _row;
	}
    public HSSFCell createCellFromRecord(CellValueRecordInterface cellRecord) {
    	return _row.createCellFromRecord(cellRecord);
    }
    public void removeAllCells() {
    	_row.removeAllCells();
    }
}
