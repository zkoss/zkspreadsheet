/* Sheet.java

	Purpose:
		
	Description:
		
	History:
		Nov 24, 2010 2:51:25 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model;

import org.zkoss.poi.ss.util.CellRangeAddress;

/**
 * ZK Spreadsheet sheet.
 * @author henrichen
 *
 */
public interface Sheet extends org.zkoss.poi.ss.usermodel.Sheet {
	/**
	 * Returns the associated ZK Spreadsheet {@link Book} of this ZK Spreadsheet Sheet. 
	 * @return the associated ZK Spreadsheet {@link Book} of this ZK Spreadsheet Sheet.
	 */
    public Book getBook();
}
