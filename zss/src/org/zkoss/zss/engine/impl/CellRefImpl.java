/* IndexableRef.java

	Purpose:
		
	Description:
		
	History:
		Mar 6, 2010 6:25:37 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.engine.impl;

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
	public int getColumnCount() {
		return 1;
	}

	@Override
	public int getRowCount() {
		return 1;
	}
	
	//--AbstractRefImpl--//
	@Override
	protected void removeSelf() {
		getOwnerSheet().removeRef(_tRow, _lCol, _tRow, _lCol);
	}
}
