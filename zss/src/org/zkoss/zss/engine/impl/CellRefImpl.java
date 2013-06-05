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
import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.engine.RefSheet;

/**
 * Reference area used with IndexableRefMatrix for hit test.
 * @author henrichen
 *
 */
public class CellRefImpl extends AbstractRefImpl implements Ref {
	private int _tRow;
	private int _lCol;
	private boolean _withIndirectPrecedent;
	
	public CellRefImpl(int row, int col, RefSheet ownerSheet) {
		super(ownerSheet);
		_tRow = row;
		_lCol = col;
	}
	
	@Override
	public int getLeftCol() {
		return _lCol;
	}

	@Override
	public int getTopRow() {
		return _tRow;
	}

	@Override
	public int getRightCol() {
		return _lCol;
	}

	@Override
	public int getBottomRow() {
		return _tRow;
	}
	
	@Override
	public void setTopRow(int row) {
		_tRow = row;
	}
	
	@Override
	public void setLeftCol(int col) {
		_lCol = col;
	}

	@Override
	public void setBottomRow(int row) {
		_tRow = row;
	}

	@Override
	public void setRightCol(int col) {
		_lCol = col;
	}

	@Override
	public boolean isWholeColumn() {
		return false; //A cell impossible to occupy the whole column
	}

	@Override
	public boolean isWholeRow() {
		return false; //A cell impossible to occupy the whole row
	}
	
	@Override
	public boolean isWholeSheet() {
		return false; //A cell impossible to occupy the whole sheet
	}
	
	@Override
	public int getColumnCount() {
		return 1;
	}

	@Override
	public int getRowCount() {
		return 1;
	}
	
	@Override
	public boolean isWithIndirectPrecedent() {
		return _withIndirectPrecedent;
	}
	
	@Override
	public void setWithIndirectPrecedent(boolean b) {
		if (_withIndirectPrecedent != b) {
			_withIndirectPrecedent = b;
			final int tRow = this.getTopRow();
			final int lCol = this.getLeftCol();
			getOwnerSheet().setRefWithIndirectPrecedent(tRow, lCol, b);
		}
	}
	
	//--AbstractRefImpl--//
	@Override
	protected void removeSelf() {
		getOwnerSheet().removeRef(_tRow, _lCol, _tRow, _lCol);
	}
	
	
	public String toString(){
		StringBuilder sb = new StringBuilder();
		sb.append("CellRef:[").append(getOwnerSheet().getOwnerBook().getBookName())
				.append("]").append(getOwnerSheet().getSheetName()).append("!")
				.append(new CellReference(_tRow, _lCol).formatAsString());
		
		return sb.toString();
	}
}
