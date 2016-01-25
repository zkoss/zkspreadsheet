/* PasteCellRegion.java

	Purpose:
		
	Description:
		
	History:
		Jan 22, 2016 6:05:55 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model;

//ZSS-717
/**
 * @author henri
 * @since 3.8.3
 */
public class PasteCellRegion extends CellRegion {
	private final boolean _wholeColumn; 
	public PasteCellRegion(int row, int column, int lastRow, int lastColumn, boolean wholeColumn) {
		super(row, column, lastRow, lastColumn);
		_wholeColumn = wholeColumn;
	}
	public boolean isWholeColumn() {
		return _wholeColumn;
	}
}
