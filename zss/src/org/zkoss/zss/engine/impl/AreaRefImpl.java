/* IndexableRef.java

	Purpose:
		
	Description:
		
	History:
		Mar 6, 2010 6:25:37 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.engine.impl;

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
	public int getColumnCount() {
		return _rCol - getLeftCol() + 1;
	}

	@Override
	public int getRowCount() {
		return _bRow - getTopRow() + 1;
	}
	
	//--Object--//
	@Override
	public int hashCode() {
		return ((_bRow << 14 + _rCol) + (getTopRow() << 14 + getLeftCol())) 
				^ (getOwnerSheet() == null ? 0 : getOwnerSheet().hashCode());
	}
	
	@Override
	public boolean equals(Object o) {
		if (o == this) return true;
		if (!(o instanceof Indexable)) return false;
		final AreaRefImpl ref = (AreaRefImpl) o;
		return ref.getTopRow() == getTopRow()  
			&& ref.getLeftCol() == getLeftCol()
			&& ref._bRow == _bRow  
			&& ref._rCol == _rCol
			&& (getOwnerSheet() == null ?
				getOwnerSheet() == ref.getOwnerSheet() : getOwnerSheet().equals(ref.getOwnerSheet()));  
	}
}
