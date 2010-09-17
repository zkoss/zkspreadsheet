/* XSSFCellHelper.java

	Purpose:
		
	Description:
		
	History:
		Sep 15, 2010 6:36:04 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.apache.poi.xssf.usermodel;

import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCell;

/**
 * A helper class to make XSSFCell package method visible.
 * @author henrichen
 *
 */
public class XSSFCellHelper {
	private final XSSFCell _cell;
	
	public XSSFCellHelper(XSSFCell cell) {
		_cell = cell;
	}
	
	public void notifyArrayFormulaChanging(String msg){
		_cell.notifyArrayFormulaChanging(msg);
	}
	
	public static XSSFCell createCell(XSSFRow row, CTCell cell) {
		return new XSSFCell(row, cell);
	}
}
