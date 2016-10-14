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
	private final boolean _wholeRow; //ZSS-1277
	//ZSS-1277
	public PasteCellRegion(int row, int column, int lastRow, int lastColumn, boolean wholeColumn, boolean wholeRow) {
		super(row, column, lastRow, lastColumn);
		_wholeColumn = wholeColumn;
		_wholeRow = wholeRow;
	}
	@Deprecated
	public PasteCellRegion(int row, int column, int lastRow, int lastColumn, boolean wholeColumn) {
		this(row, column, lastRow, lastColumn, wholeColumn, false);
	}
	public boolean isWholeColumn() {
		return _wholeColumn;
	}
	//ZSS-1277
	public boolean isWholeRow() {
		return _wholeRow;
	}
}
