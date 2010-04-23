/* DummyHSSFCell.java

	Purpose:
		
	Description:
		
	History:
		Apr 14, 2010 2:18:14 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.apache.poi.hssf.usermodel;

import org.apache.poi.hssf.record.CellValueRecordInterface;

/**
 * Dummy HSSFCell class to expose the under record.
 * @author henrichen
 *
 */
public class DummyHSSFCell extends HSSFCell {
	private final HSSFCell _cell;
	
	public DummyHSSFCell(HSSFCell cell) {
		super((HSSFWorkbook) cell.getSheet().getWorkbook(), (HSSFSheet) cell.getSheet(), cell.getRowIndex(), (short) cell.getColumnIndex());
		_cell = cell;
	}
	
	public CellValueRecordInterface getRecord() {
		return _cell.getCellValueRecord();
	}
}
