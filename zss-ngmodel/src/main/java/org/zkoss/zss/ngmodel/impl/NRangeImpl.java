package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngmodel.NSheet;


public class NRangeImpl implements NRange {
	private final NSheet _sheet;
	private int _column = Integer.MAX_VALUE;
	private int _top = Integer.MAX_VALUE;
	private int _lastColumn = Integer.MIN_VALUE;
	private int _lastRow = Integer.MIN_VALUE;
	public NRangeImpl(NSheet sheet) {
		_sheet = sheet;
		addRef(sheet, 0, 0, sheet.getBook().getMaxRowSize(), sheet.getBook().getMaxColumnSize());
	}
	public NRangeImpl(NSheet sheet, int row, int col) {
		_sheet = sheet;
		addRef(sheet, row, col, row, col);
	}
	
	public NRangeImpl(NSheet sheet, int tRow, int lCol, int bRow, int rCol) {
		_sheet = sheet;
		addRef(sheet, tRow, lCol, bRow, rCol);
	}

	private void addRef(NSheet sheet, int tRow, int lCol, int bRow, int rCol) {
		//TODO possible to have multiple sheets ref as XRange
		_column = Math.min(_column,lCol);
		_top = Math.min(_top, tRow);
		_lastColumn = Math.max(_lastColumn,rCol);
		_lastRow = Math.max(_lastRow, bRow);
	}
	public NSheet getSheet() {
		return _sheet;
	}
	public int getRow() {
		return _top;
	}
	public int getColumn() {
		return _column;
	}
	public int getLastRow() {
		return _lastRow;
	}
	public int getLastColumn() {
		return _lastColumn;
	}
	
	/*package*/
}
