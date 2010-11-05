/* XSSFRowHelper.java

	Purpose:
		
	Description:
		
	History:
		Sep 14, 2010 6:55:33 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.poi.xssf.usermodel;

/**
 * A helper class to make XSSFRow package method visible.
 * @author henrichen
 *
 */
public class XSSFRowHelper {
	private final XSSFRow _row;
	public XSSFRowHelper(XSSFRow row) {
		_row = row;
	}
	
    public void shift(int n) {
    	_row.shift(n);
    }
}
