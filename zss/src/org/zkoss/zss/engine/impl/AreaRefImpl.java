/* IndexableRef.java

	Purpose:
		
	Description:
		
	History:
		Mar 6, 2010 6:25:37 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.engine.impl;

import org.zkoss.poi.ss.util.AreaReference;
import org.zkoss.poi.ss.util.CellReference;
import org.zkoss.zss.engine.RefBook;
import org.zkoss.zss.engine.RefSheet;

/**
 * Reference area used with IndexableRefMatrix for hit test.
 * @author henrichen
 *
 */
public class AreaRefImpl extends CellRefImpl {
	private int _bRow;
	private int _rCol;
	
	public AreaRefImpl(int tRow, int lCol, int bRow, int rCol, RefSheet ownerSheet) {
		super(tRow, lCol, ownerSheet);
		_bRow = bRow;
		_rCol = rCol;
	}
	
	@Override
	public int getRightCol() {
		return _rCol;
	}

	@Override
	public int getBottomRow() {
		return _bRow;
	}
	
	@Override
	public void setBottomRow(int row) {
		_bRow = row;
	}

	@Override
	public void setRightCol(int col) {
		_rCol = col;
	}

	@Override
	public boolean isWholeColumn() {
		final RefBook book = getOwnerSheet().getOwnerBook(); 
		return getTopRow() <= 0 && _bRow >= book.getMaxrow();
	}

	@Override
	public boolean isWholeRow() {
		final RefBook book = getOwnerSheet().getOwnerBook(); 
		return getLeftCol() <= 0 && _rCol >= book.getMaxcol();
	}
	
	@Override
	public boolean isWholeSheet() {
		return isWholeColumn() && isWholeRow(); 
	}
	
	@Override
	public int getColumnCount() {
		return _rCol - getLeftCol() + 1;
	}

	@Override
	public int getRowCount() {
		return _bRow - getTopRow() + 1;
	}
	
	public String toString(){
		StringBuilder sb = new StringBuilder();
		sb.append("AreaRef:[")
				.append(getOwnerSheet().getOwnerBook().getBookName())
				.append("]")
				.append(getOwnerSheet().getSheetName())
				.append("!")
				.append(new AreaReference(new CellReference(getTopRow(),
						getLeftCol()), new CellReference(getBottomRow(),
						getRightCol())).formatAsString());

		return sb.toString();
	}
}
